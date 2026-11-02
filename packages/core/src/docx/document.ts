// WordprocessingML editor: an outline of the body (paragraphs and tables) for the edit view, and
// paragraph-level edits applied to word/document.xml as minimal changes. Paragraphs whose content
// cannot be rewritten as plain runs without losing something (fields, hyperlinks, pictures,
// tracked changes, ...) are reported as locked and refused.
import { EditError } from '../errors';
import { OoxmlPackage } from '../package';
import type { AppliedChange, EditableDocument } from '../session';
import { paragraphInfo, paragraphText, readContext, type DocxContext } from '../text/docx';
import { attr, child, children, createLike, numAttr, path, type Document, type Element } from '../xml';
import type { DocxBlock, DocxEditOp, DocxParagraph, DocxTable } from './types';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
/** Characters XML 1.0 cannot represent; dropped from typed text. */
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

const TRACKED = 'Paragraphs with tracked changes cannot be edited here. Accept or reject the changes in Word first.';
const REASONS: Record<string, string> = {
  hyperlink: 'Paragraphs with hyperlinks cannot be edited here yet.',
  fldSimple: 'Paragraphs with fields (page numbers, references, ...) cannot be edited here yet.',
  ins: TRACKED,
  del: TRACKED,
  moveFrom: TRACKED,
  moveTo: TRACKED,
  moveFromRangeStart: TRACKED,
  moveFromRangeEnd: TRACKED,
  moveToRangeStart: TRACKED,
  moveToRangeEnd: TRACKED,
  sdt: 'Paragraphs with content controls cannot be edited here yet.',
  commentRangeStart: 'Paragraphs with comments cannot be edited here yet.',
  commentRangeEnd: 'Paragraphs with comments cannot be edited here yet.',
  customXml: 'Paragraphs with tagged (custom XML) content cannot be edited here yet.',
  smartTag: 'Paragraphs with smart tags cannot be edited here yet.',
  oMath: 'Paragraphs with equations cannot be edited here yet.',
  oMathPara: 'Paragraphs with equations cannot be edited here yet.',
  permStart: 'Paragraphs with editing permissions cannot be edited here yet.',
  permEnd: 'Paragraphs with editing permissions cannot be edited here yet.',
};
const RUN_REASONS: Record<string, string> = {
  fldChar: REASONS.fldSimple,
  instrText: REASONS.fldSimple,
  delText: TRACKED,
  rPrChange: TRACKED,
  drawing: 'Paragraphs with pictures or shapes cannot be edited here yet.',
  pict: 'Paragraphs with pictures or shapes cannot be edited here yet.',
  object: 'Paragraphs with embedded objects cannot be edited here yet.',
  footnoteReference: 'Paragraphs with footnote or endnote references cannot be edited here yet.',
  endnoteReference: 'Paragraphs with footnote or endnote references cannot be edited here yet.',
  commentReference: REASONS.commentRangeStart,
  ruby: 'Paragraphs with ruby (furigana) cannot be edited here yet.',
};
const BOOKMARK = 'A bookmark starts or ends inside this paragraph, so it cannot be edited here yet.';
/** Markers that may stay at either end of a rewritten paragraph. */
const EDGE_MARKERS = new Set(['bookmarkStart', 'bookmarkEnd']);
/** Spelling and grammar marks; Word recomputes them. */
const DROPPABLE = new Set(['proofErr']);
/** Run children that hold no text but do not stop a rewrite (Word recomputes or ignores them). */
const RUN_IGNORED = new Set(['rPr', 'lastRenderedPageBreak']);

interface Analysis {
  text: string;
  locked?: string;
}

interface Slot {
  /** The body or table cell holding the element. */
  container: Element;
  index: number;
  element?: Element;
}

export class WordDocument implements EditableDocument<DocxEditOp> {
  private modified = false;

  private constructor(
    private readonly pkg: OoxmlPackage,
    private readonly mainPath: string,
    private readonly ctx: DocxContext,
    private readonly original: Uint8Array,
  ) {}

  static async open(data: Uint8Array): Promise<WordDocument> {
    const pkg = await OoxmlPackage.open(data);
    const mainPath = await pkg.mainPart();
    const root = (await pkg.xml(mainPath))?.documentElement;
    if (!root || root.localName !== 'document' || !child(root, 'body')) throw new Error('This file is not a Word document.');
    return new WordDocument(pkg, mainPath, await readContext(pkg, mainPath), data);
  }

  /** The editable outline of the body as it is now. */
  async blocks(): Promise<DocxBlock[]> {
    return this.readBlocks(await this.body(), '');
  }

  /** The file as it would be saved, for the preview. */
  async bytes(): Promise<Uint8Array> {
    return this.modified ? this.pkg.save() : this.original;
  }

  save(): Promise<Uint8Array> {
    return this.pkg.save();
  }

  async apply(op: DocxEditOp): Promise<AppliedChange> {
    const doc = await this.pkg.edit(this.mainPath);
    const result = this.applyTo(doc, op);
    this.modified = true;
    return result;
  }

  private applyTo(doc: Document, op: DocxEditOp): AppliedChange {
    switch (op.type) {
      case 'setParagraphText': {
        const p = this.paragraph(doc, op.id);
        const before = p.cloneNode(true) as Element;
        this.setText(doc, p, op.text);
        return { label: 'Edit Paragraph', revert: async () => this.replace(op.id, before) };
      }
      case 'insertParagraph': {
        const p = this.paragraph(doc, op.after, op.update === undefined);
        const before = p.cloneNode(true) as Element;
        const created = this.newParagraph(doc, p, op.text);
        if (op.update !== undefined) {
          this.setText(doc, p, op.update);
          // Splitting a section's last paragraph: the section break stays at the end, in the new paragraph.
          const sectPr = path(p, 'pPr', 'sectPr');
          if (sectPr) {
            sectPr.parentNode!.removeChild(sectPr);
            ensurePPr(doc, created).appendChild(sectPr);
          }
        }
        p.parentNode!.insertBefore(created, p.nextSibling);
        const newId = siblingId(op.after, 1);
        return {
          label: op.update === undefined ? 'Insert Paragraph' : 'Split Paragraph',
          revert: async () => {
            const slot = this.slot(await this.pkg.edit(this.mainPath), newId);
            if (slot.element) slot.container.removeChild(slot.element);
            await this.replace(op.after, before);
          },
        };
      }
      case 'deleteParagraph': {
        const slot = this.slot(doc, op.id);
        const p = this.paragraph(doc, op.id);
        const reason = this.deleteBlocker(slot, p);
        if (reason) throw new EditError(reason);
        slot.container.removeChild(p);
        return {
          label: 'Delete Paragraph',
          revert: async () => {
            const current = this.slot(await this.pkg.edit(this.mainPath), op.id);
            current.container.insertBefore(p, current.element ?? null);
          },
        };
      }
      default:
        throw new EditError(`Unknown edit "${(op as { type?: string }).type}".`);
    }
  }

  // --- Reading -------------------------------------------------------------------------------

  private async body(): Promise<Element> {
    return child((await this.pkg.xml(this.mainPath))!.documentElement, 'body')!;
  }

  private readBlocks(container: Element, prefix: string): DocxBlock[] {
    return children(container).flatMap((el, i): DocxBlock[] => {
      const id = prefix + i;
      switch (el.localName) {
        case 'p': {
          const { text, locked } = analyzeParagraph(el);
          const info = paragraphInfo(el, this.ctx);
          const block: DocxParagraph = { type: 'paragraph', id, text: locked ? paragraphText(el) : text, ...info };
          if (locked) block.locked = locked;
          return [block];
        }
        case 'tbl':
          return [this.readTable(el, id)];
        case 'sdt':
          return [{ type: 'other', id, label: 'Content control', text: blockText(child(el, 'sdtContent')) }];
        case 'customXml':
          return [{ type: 'other', id, label: 'Custom XML', text: blockText(el) }];
        case 'altChunk':
          return [{ type: 'other', id, label: 'Embedded document', text: '' }];
        default:
          return [];
      }
    });
  }

  private readTable(tbl: Element, id: string): DocxTable {
    return {
      type: 'table',
      id,
      rows: children(tbl, 'tr').map((tr, r) => ({
        cells: children(tr, 'tc').map((tc, c) => {
          const cellId = `${id}/${r}/${c}`;
          const vMerge = path(tc, 'tcPr', 'vMerge');
          return {
            id: cellId,
            blocks: this.readBlocks(tc, `${cellId}/`),
            colSpan: numAttr(path(tc, 'tcPr', 'gridSpan'), 'val') ?? 1,
            merged: !!vMerge && attr(vMerge, 'val') !== 'restart',
          };
        }),
      })),
    };
  }

  // --- Locating ------------------------------------------------------------------------------

  /** Resolves an id to its container and position; the element is absent past the end. */
  private slot(doc: Document, id: string): Slot {
    const segments = id.split('/').map(Number);
    if (!segments.length || segments.some((n) => !Number.isInteger(n) || n < 0) || segments.length % 3 !== 1) {
      throw new EditError(`"${id}" is not a paragraph id.`);
    }
    let container = child(doc.documentElement, 'body')!;
    for (let i = 0; ; i += 3) {
      const element = children(container)[segments[i]];
      if (i === segments.length - 1) return { container, index: segments[i], element };
      const tc = element?.localName === 'tbl' ? children(children(element, 'tr')[segments[i + 1]], 'tc')[segments[i + 2]] : undefined;
      if (!tc) throw new EditError('That paragraph no longer exists.');
      container = tc;
    }
  }

  /** The paragraph at `id`; unless `anyContent`, it must be one that can be rewritten. */
  private paragraph(doc: Document, id: string, anyContent = false): Element {
    const p = this.slot(doc, id).element;
    if (p?.localName !== 'p') throw new EditError('That paragraph no longer exists.');
    if (!anyContent) {
      const { locked } = analyzeParagraph(p);
      if (locked) throw new EditError(locked);
    }
    return p;
  }

  private async replace(id: string, element: Element): Promise<void> {
    const slot = this.slot(await this.pkg.edit(this.mainPath), id);
    if (slot.element) slot.container.replaceChild(element, slot.element);
  }

  private deleteBlocker(slot: Slot, p: Element): string | undefined {
    const { locked } = analyzeParagraph(p);
    if (locked) return locked;
    if (children(p).some((el) => EDGE_MARKERS.has(el.localName ?? ''))) return 'Paragraphs with bookmarks cannot be deleted here yet.';
    if (path(p, 'pPr', 'sectPr')) return 'This paragraph ends a section, so it cannot be deleted here.';
    const paragraphs = children(slot.container, 'p');
    const blocks = children(slot.container).filter((el) => ['p', 'tbl', 'sdt', 'customXml'].includes(el.localName ?? ''));
    const at = blocks.indexOf(p);
    // Two tables with nothing between them become one table in Word.
    if (blocks[at - 1]?.localName === 'tbl' && blocks[at + 1]?.localName === 'tbl') {
      return 'This paragraph separates two tables, so it cannot be deleted (Word would merge them).';
    }
    // Word requires a paragraph at the end of the document and of every table cell.
    if (paragraphs.length === 1 || blocks[blocks.length - 1] === p) {
      return slot.container.localName === 'tc'
        ? 'The last paragraph of a table cell cannot be deleted.'
        : 'The last paragraph of the document cannot be deleted.';
    }
    return undefined;
  }

  // --- Writing -------------------------------------------------------------------------------

  /**
   * Replaces the runs of a paragraph with one run of plain text. The paragraph properties stay,
   * the first text run's formatting applies to the new text, and bookmarks at either end are kept.
   */
  private setText(doc: Document, p: Element, value: string): void {
    const items = children(p).filter((el) => el.localName !== 'pPr');
    const content = items.flatMap((el, i) => (isContentRun(el) ? [i] : []));
    const first = content.length ? content[0] : items.length;
    const last = content.length ? content[content.length - 1] : items.length;
    const rPr = templateRunProps(doc, p);
    for (const el of items) p.removeChild(el);
    const keep = (list: Element[]) => {
      for (const el of list) if (EDGE_MARKERS.has(el.localName ?? '')) p.appendChild(el);
    };
    keep(items.slice(0, first));
    if (value) p.appendChild(buildRun(doc, p, rPr, value));
    keep(items.slice(last + 1));
  }

  /** A new paragraph with the properties of `like` (without its section break) and the given text. */
  private newParagraph(doc: Document, like: Element, value: string): Element {
    const p = createLike(doc, like, 'p');
    const pPr = child(like, 'pPr');
    if (pPr) {
      const copy = pPr.cloneNode(true) as Element;
      for (const el of children(copy)) if (el.localName === 'sectPr' || el.localName === 'pPrChange') copy.removeChild(el);
      const markProps = child(copy, 'rPr');
      if (markProps) stripTracking(markProps);
      p.appendChild(copy);
    }
    if (value) p.appendChild(buildRun(doc, like, templateRunProps(doc, like), value));
    return p;
  }
}

/** The id of the element `delta` positions after `id` in the same container. */
function siblingId(id: string, delta: number): string {
  const parts = id.split('/');
  parts[parts.length - 1] = String(Number(parts[parts.length - 1]) + delta);
  return parts.join('/');
}

/** Plain text of a paragraph that can be rewritten, or why it cannot be. */
function analyzeParagraph(p: Element): Analysis {
  let text = '';
  let locked: string | undefined;
  const lock = (reason: string) => (locked ??= reason);
  const pPr = child(p, 'pPr');
  if (child(pPr, 'pPrChange')) lock(TRACKED);
  const markProps = child(pPr, 'rPr');
  if (children(markProps).some((el) => ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange'].includes(el.localName ?? ''))) lock(TRACKED);

  const items = children(p).filter((el) => el.localName !== 'pPr');
  const content = items.flatMap((el, i) => (isContentRun(el) ? [i] : []));
  items.forEach((el, i) => {
    const name = el.localName ?? '';
    if (name === 'r') {
      const run = analyzeRun(el);
      text += run.text;
      if (run.locked) lock(run.locked);
    } else if (EDGE_MARKERS.has(name)) {
      if (content.length && i > content[0] && i < content[content.length - 1]) lock(BOOKMARK);
    } else if (!DROPPABLE.has(name)) {
      lock(REASONS[name] ?? `Paragraphs containing "${name}" cannot be edited here yet.`);
    }
  });
  return { text, locked };
}

function analyzeRun(r: Element): Analysis {
  let text = '';
  let locked: string | undefined;
  for (const el of children(r)) {
    const name = el.localName ?? '';
    switch (name) {
      case 't':
        text += el.textContent ?? '';
        break;
      case 'tab':
        text += '\t';
        break;
      case 'br':
      case 'cr': {
        const type = attr(el, 'type');
        if ((type && type !== 'textWrapping') || attr(el, 'clear')) locked ??= 'Paragraphs with page or column breaks cannot be edited here yet.';
        text += '\n';
        break;
      }
      case 'rPr':
        if (child(el, 'rPrChange')) locked ??= TRACKED;
        break;
      default:
        if (!RUN_IGNORED.has(name)) locked ??= RUN_REASONS[name] ?? `Paragraphs containing "${name}" cannot be edited here yet.`;
    }
  }
  return { text, locked };
}

/** A run with something in it besides formatting. */
function isContentRun(el: Element): boolean {
  return el.localName === 'r' && children(el).some((c) => !RUN_IGNORED.has(c.localName ?? ''));
}

/** Formatting for new text: the first text run's, else the first run's, else the paragraph mark's. */
function templateRunProps(doc: Document, p: Element): Element | undefined {
  const runs = children(p, 'r');
  const textRun = runs.find((r) => children(r).some((c) => c.localName === 't')) ?? runs.find(isContentRun);
  const own = child(textRun, 'rPr');
  if (own) return own.cloneNode(true) as Element;
  if (textRun) return undefined;
  const mark = path(p, 'pPr', 'rPr');
  if (!mark) return undefined;
  const rPr = createLike(doc, mark, 'rPr');
  for (const el of children(mark)) rPr.appendChild(el.cloneNode(true));
  stripTracking(rPr);
  return rPr;
}

/** Removes revision marks from run properties (they describe the old text, not the new). */
function stripTracking(rPr: Element): void {
  for (const el of children(rPr)) {
    if (['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange'].includes(el.localName ?? '')) rPr.removeChild(el);
  }
}

/** `<w:r>` with the given properties and text; "\t" becomes `<w:tab/>` and line breaks `<w:br/>`. */
function buildRun(doc: Document, like: Element, rPr: Element | undefined, value: string): Element {
  const r = createLike(doc, like, 'r');
  if (rPr) r.appendChild(rPr);
  for (const part of value.replace(INVALID_XML_CHARS, '').split(/(\t|\r\n|\n|\r)/)) {
    if (!part) continue;
    if (part === '\t') r.appendChild(createLike(doc, like, 'tab'));
    else if (/^(\r\n|\n|\r)$/.test(part)) r.appendChild(createLike(doc, like, 'br'));
    else {
      const t = createLike(doc, like, 't');
      if (/^\s|\s$/.test(part)) t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(part));
      r.appendChild(t);
    }
  }
  return r;
}

function ensurePPr(doc: Document, p: Element): Element {
  let pPr = child(p, 'pPr');
  if (!pPr) {
    pPr = createLike(doc, p, 'pPr');
    p.insertBefore(pPr, p.firstChild);
  }
  return pPr;
}

/** Text of the paragraphs inside a block-level container (content controls and similar). */
function blockText(el: Element | undefined): string {
  if (!el) return '';
  const parts: string[] = [];
  for (const node of children(el)) {
    if (node.localName === 'p') parts.push(paragraphText(node));
    else parts.push(blockText(node));
  }
  return parts.filter(Boolean).join('\n');
}
