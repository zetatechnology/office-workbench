// PresentationML editor: text of shapes and table cells on a slide, and slide duplication,
// deletion and reordering. Text edits touch only the text body of one shape (undo puts back a
// copy of it); slide operations run with the package journal so undo restores every part.
import { EditError } from '../errors';
import { OoxmlPackage, relativeTarget, relsPathOf, relsSourceOf } from '../package';
import type { AppliedChange, EditableDocument } from '../session';
import { attr, child, children, createLike, descendants, path, relAttr, type Document, type Element } from '../xml';
import { NON_TEXT_PLACEHOLDERS, readPresentation, slideParts, type ParsedPresentation } from './presentation';
import type { PptxEditOp } from './types';

const CONTENT_TYPES = '[Content_Types].xml';
const REL_TYPES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const NOTES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
/** Characters XML 1.0 cannot represent; dropped from typed text. */
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;
const LINE_BREAK = /\r\n|[\n\v\r]/;
const CJK = /[　-ヿ㐀-鿿豈-﫿＀-￯]/;
/** Parts a duplicated slide may share with the original. */
const SHAREABLE = new Set(['slideLayout', 'image', 'media', 'video', 'audio', 'hdphoto', 'slide']);
const NOT_COPYABLE: Record<string, string> = {
  chart: 'charts',
  chartUsingStyle: 'charts',
  diagramData: 'SmartArt',
  oleObject: 'embedded objects',
  package: 'embedded objects',
  control: 'ActiveX controls',
  vmlDrawing: 'legacy drawings',
  tags: 'add-in data (tags)',
};

interface Located {
  doc: Document;
  /** The shape (`p:sp`) or table cell (`a:tc`) that holds the text body. */
  host: Element;
}

type TextTarget = { slide: number; slideId?: string; shape: string; cell?: { row: number; col: number } };

export class PresentationDocument implements EditableDocument<PptxEditOp> {
  private constructor(
    private readonly pkg: OoxmlPackage,
    private readonly presPath: string,
  ) {}

  static async open(data: Uint8Array): Promise<PresentationDocument> {
    const pkg = await OoxmlPackage.open(data);
    const presPath = await pkg.mainPart();
    const root = (await pkg.xml(presPath))?.documentElement;
    if (!root || root.localName !== 'presentation') throw new Error('This file is not a PowerPoint presentation.');
    return new PresentationDocument(pkg, presPath);
  }

  read(withMedia: boolean): Promise<ParsedPresentation> {
    return readPresentation(this.pkg, withMedia);
  }

  save(): Promise<Uint8Array> {
    return this.pkg.save();
  }

  async apply(op: PptxEditOp): Promise<AppliedChange> {
    switch (op.type) {
      case 'setShapeText':
        return this.setText(await this.target(op), op.paragraphs);
      case 'setTableCellText':
        return this.setText({ ...(await this.target(op)), cell: { row: op.row, col: op.col } }, op.paragraphs);
      case 'duplicateSlide':
        return this.structural('Duplicate Slide', () => this.duplicateSlide(op.slide));
      case 'deleteSlide':
        return this.structural('Delete Slide', () => this.deleteSlide(op.slide));
      case 'moveSlide':
        return this.structural('Move Slide', () => this.moveSlide(op.slide, op.to));
      default:
        throw new EditError(`Unknown edit "${(op as { type?: string }).type}".`);
    }
  }

  // --- Text ----------------------------------------------------------------------------------

  /** The slide of a text edit by position, or by id when the edit names one. */
  private async target(op: { slide: number; slideId?: string; shape: string }): Promise<TextTarget> {
    if (op.slideId === undefined) return { slide: op.slide, shape: op.shape };
    const slide = (await slideParts(this.pkg, this.presPath)).findIndex((s) => attr(s.sldId, 'id') === op.slideId);
    if (slide < 0) throw new EditError('The slide being edited was deleted.');
    return { slide, shape: op.shape };
  }

  private async setText(target: TextTarget, paragraphs: string[]): Promise<AppliedChange> {
    if (!Array.isArray(paragraphs) || paragraphs.some((p) => typeof p !== 'string')) throw new EditError('paragraphs must be a list of strings.');
    const { doc, host } = await this.locate(target);
    const before = child(host, 'txBody')?.cloneNode(true) as Element | undefined;
    writeParagraphs(doc, ensureTextBody(doc, host), paragraphs.length ? paragraphs : [''], target.slide + 1);
    return {
      label: target.cell ? 'Edit Table Cell' : 'Edit Text',
      activate: target.slide,
      revert: async () => {
        const { host: current } = await this.locate(target);
        const body = child(current, 'txBody');
        if (body && before) current.replaceChild(before, body);
        else if (body) current.removeChild(body);
        return { activate: target.slide };
      },
    };
  }

  private async locate({ slide, shape, cell }: TextTarget): Promise<Located> {
    const slides = await slideParts(this.pkg, this.presPath);
    const entry = slides[slide];
    if (!entry) throw new EditError(`Slide ${slide + 1} does not exist.`);
    const doc = await this.pkg.edit(entry.path);
    const cNvPr = descendants(path(doc.documentElement, 'cSld', 'spTree'), 'cNvPr').find((el) => attr(el, 'id') === shape);
    const frame = cNvPr?.parentNode?.parentNode as Element | undefined;
    if (!cNvPr || !frame || insideAlternateContent(frame)) throw new EditError('That shape no longer exists on the slide.');
    if (!cell) {
      if (frame.localName !== 'sp') throw new EditError('Only the text of shapes and text boxes can be edited.');
      if (NON_TEXT_PLACEHOLDERS.has(attr(path(frame, 'nvSpPr', 'nvPr', 'ph'), 'type') ?? '')) {
        throw new EditError('This placeholder holds a picture, chart or table, not text.');
      }
      return { doc, host: frame };
    }
    const tbl = path(frame, 'graphic', 'graphicData', 'tbl');
    if (frame.localName !== 'graphicFrame' || !tbl) throw new EditError('That shape is not a table.');
    const tc = children(children(tbl, 'tr')[cell.row], 'tc')[cell.col];
    if (!tc) throw new EditError('That table cell no longer exists.');
    if (['hMerge', 'vMerge'].some((name) => ['1', 'true'].includes(attr(tc, name) ?? ''))) {
      throw new EditError('This cell is covered by a merged cell; edit the merged cell instead.');
    }
    return { doc, host: tc };
  }

  // --- Slides --------------------------------------------------------------------------------

  /** Runs a slide operation with the package journal on: undo restores every touched part. */
  private async structural(label: string, run: () => Promise<number>): Promise<AppliedChange> {
    this.pkg.begin();
    let activate: number;
    try {
      activate = await run();
    } catch (e) {
      this.pkg.restore(this.pkg.commit());
      throw e;
    }
    const snapshot = this.pkg.commit();
    return {
      label,
      activate,
      revert: async () => {
        this.pkg.restore(snapshot);
        return {};
      },
    };
  }

  private async slide(index: number) {
    const slides = await slideParts(this.pkg, this.presPath);
    const entry = slides[index];
    if (!entry) throw new EditError(`Slide ${index + 1} does not exist.`);
    return { slides, entry };
  }

  private async duplicateSlide(index: number): Promise<number> {
    const { entry } = await this.slide(index);
    const rels = [...(await this.pkg.rels(entry.path)).values()];
    for (const rel of rels) {
      if (rel.external || SHAREABLE.has(rel.type) || rel.type === 'notesSlide' || rel.type === 'comments') continue;
      throw new EditError(`Slides with ${NOT_COPYABLE[rel.type] ?? `"${rel.type}" content`} cannot be duplicated yet.`);
    }

    const slidePath = this.freePart('ppt/slides/slide', '.xml');
    await this.pkg.setText(slidePath, (await this.pkg.text(entry.path))!);
    const slideDoc = await this.pkg.edit(slidePath);
    // PowerPoint identifies slides by this id (e.g. for co-authoring); give the copy its own.
    for (const el of descendants(slideDoc, 'creationId')) {
      if (el.parentNode && (el.parentNode as Element).localName === 'ext' && attr(el, 'val')) el.setAttribute('val', String(randomId()));
    }

    // The copy's relationships: same targets, except its own notes and no comments.
    const relsText = await this.pkg.text(relsPathOf(entry.path));
    const notesRel = rels.find((r) => r.type === 'notesSlide' && !r.external);
    if (relsText) {
      const newRels = relsPathOf(slidePath);
      await this.pkg.setText(newRels, relsText);
      const relsDoc = await this.pkg.edit(newRels);
      for (const el of children(relsDoc.documentElement, 'Relationship')) {
        const type = relType(el);
        if (type === 'comments') relsDoc.documentElement!.removeChild(el);
        else if (type === 'notesSlide' && notesRel) {
          const notesPath = await this.copyNotes(notesRel.target, slidePath);
          el.setAttribute('Target', relativeTarget(slidePath, notesPath));
        }
      }
    }
    await this.addContentType(slidePath, SLIDE_CONTENT_TYPE);

    // presentation.xml: a new relationship and slide id right after the original.
    const rId = await this.addRelationship(this.presPath, 'slide', slidePath);
    const presDoc = await this.pkg.edit(this.presPath);
    const id = nextSlideId(presDoc);
    const sldId = entry.sldId.cloneNode(false) as Element;
    sldId.setAttribute('id', String(id));
    setRelId(sldId, rId);
    entry.sldId.parentNode!.insertBefore(sldId, entry.sldId.nextSibling);
    const section = sectionEntry(presDoc, attr(entry.sldId, 'id'));
    if (section) {
      const copy = section.cloneNode(false) as Element;
      copy.setAttribute('id', String(id));
      section.parentNode!.insertBefore(copy, section.nextSibling);
    }
    return index + 1;
  }

  private async copyNotes(notesPath: string, slidePath: string): Promise<string> {
    const copyPath = this.freePart('ppt/notesSlides/notesSlide', '.xml');
    await this.pkg.setText(copyPath, (await this.pkg.text(notesPath))!);
    const relsText = await this.pkg.text(relsPathOf(notesPath));
    if (relsText) {
      await this.pkg.setText(relsPathOf(copyPath), relsText);
      const relsDoc = await this.pkg.edit(relsPathOf(copyPath));
      for (const el of children(relsDoc.documentElement, 'Relationship')) {
        if (relType(el) === 'slide') el.setAttribute('Target', relativeTarget(copyPath, slidePath));
      }
    }
    await this.addContentType(copyPath, NOTES_CONTENT_TYPE);
    return copyPath;
  }

  private async deleteSlide(index: number): Promise<number> {
    const { slides, entry } = await this.slide(index);
    if (slides.length === 1) throw new EditError('A presentation must keep at least one slide.');
    const presDoc = await this.pkg.edit(this.presPath);
    const shows = descendants(child(presDoc.documentElement, 'custShowLst'), 'sldId');
    if (shows.some((el) => relAttr(el, 'id') === entry.rId)) throw new EditError('This slide is part of a custom show, so it cannot be deleted here.');

    // Parts that belong to the slide alone go with it; everything else it uses stays.
    const own = [...(await this.pkg.rels(entry.path)).values()].filter((r) => !r.external && (r.type === 'notesSlide' || r.type === 'comments'));
    const removing = new Set([entry.path, ...own.map((r) => r.target)]);
    for (const relsPath of this.pkg.parts().filter((p) => p.endsWith('.rels'))) {
      const source = relsSourceOf(relsPath);
      if (source === undefined || removing.has(source) || source === this.presPath) continue;
      const links = [...(await this.pkg.rels(source)).values()].some((r) => !r.external && r.target === entry.path);
      if (links) throw new EditError('Another slide links to this slide, so it cannot be deleted here yet.');
    }

    entry.sldId.parentNode!.removeChild(entry.sldId);
    const section = sectionEntry(presDoc, attr(entry.sldId, 'id'));
    section?.parentNode?.removeChild(section);
    const presRels = await this.pkg.edit(relsPathOf(this.presPath));
    for (const el of children(presRels.documentElement, 'Relationship')) {
      if (attr(el, 'Id') === entry.rId) presRels.documentElement!.removeChild(el);
    }
    for (const part of removing) {
      await this.pkg.remove(part);
      await this.pkg.remove(relsPathOf(part));
      await this.removeContentType(part);
    }
    return Math.min(index, slides.length - 2);
  }

  private async moveSlide(index: number, to: number): Promise<number> {
    const { slides, entry } = await this.slide(index);
    if (!Number.isInteger(to) || to < 0 || to >= slides.length) throw new EditError('The new position is outside the presentation.');
    if (to === index) return index;
    const presDoc = await this.pkg.edit(this.presPath);
    const list = entry.sldId.parentNode as Element;
    list.removeChild(entry.sldId);
    const rest = slides.filter((s) => s !== entry);
    list.insertBefore(entry.sldId, rest[to]?.sldId ?? null);

    // Sections list their slides in presentation order: the slide joins the section of its new
    // predecessor (or of its successor when it becomes the first slide).
    const section = sectionEntry(presDoc, attr(entry.sldId, 'id'));
    if (section) {
      section.parentNode!.removeChild(section);
      const prev = to > 0 ? sectionEntry(presDoc, attr(rest[to - 1].sldId, 'id')) : undefined;
      const next = to === 0 ? sectionEntry(presDoc, attr(rest[0].sldId, 'id')) : undefined;
      if (prev) prev.parentNode!.insertBefore(section, prev.nextSibling);
      else if (next) next.parentNode!.insertBefore(section, next);
    }
    return to;
  }

  // --- Package bookkeeping -------------------------------------------------------------------

  /** `prefix` + the lowest number not used by an existing part + `suffix`. */
  private freePart(prefix: string, suffix: string): string {
    for (let n = 1; ; n++) {
      const candidate = `${prefix}${n}${suffix}`;
      if (!this.pkg.has(candidate)) return candidate;
    }
  }

  private async addContentType(partPath: string, contentType: string): Promise<void> {
    const doc = await this.pkg.edit(CONTENT_TYPES);
    const root = doc.documentElement!;
    const override = createLike(doc, root, 'Override');
    override.setAttribute('PartName', `/${partPath}`);
    override.setAttribute('ContentType', contentType);
    root.appendChild(override);
  }

  private async removeContentType(partPath: string): Promise<void> {
    const doc = await this.pkg.edit(CONTENT_TYPES);
    for (const el of children(doc.documentElement, 'Override')) {
      if ((attr(el, 'PartName') ?? '').replace(/^\//, '').toLowerCase() === partPath.toLowerCase()) doc.documentElement!.removeChild(el);
    }
  }

  private async addRelationship(source: string, type: string, target: string): Promise<string> {
    const relsPath = relsPathOf(source);
    const doc = await this.pkg.edit(relsPath);
    const root = doc.documentElement!;
    const ids = new Set(children(root, 'Relationship').map((el) => attr(el, 'Id')));
    let n = ids.size + 1;
    while (ids.has(`rId${n}`)) n++;
    const rel = createLike(doc, root, 'Relationship');
    rel.setAttribute('Id', `rId${n}`);
    rel.setAttribute('Type', `${REL_TYPES}/${type}`);
    rel.setAttribute('Target', relativeTarget(source, target));
    root.appendChild(rel);
    return `rId${n}`;
  }
}

// ---------------------------------------------------------------------------------------------
// Text bodies

/** The shape's or cell's text body, created (empty) when it has none. */
function ensureTextBody(doc: Document, host: Element): Element {
  const existing = child(host, 'txBody');
  if (existing) return existing;
  const drawing = host.localName === 'tc' ? host : undefined;
  // Shapes use p:txBody (placed before any extLst); table cells a:txBody (first child).
  const body = drawing ? createLike(doc, host, 'txBody') : createLike(doc, child(host, 'spPr') ?? host, 'txBody');
  const a = (name: string) => doc.createElementNS(DRAWING_NS, `a:${name}`);
  const bodyPr = a('bodyPr');
  const p = a('p');
  const placeholder = !!path(host, 'nvSpPr', 'nvPr', 'ph');
  const textBox = attr(path(host, 'nvSpPr', 'cNvSpPr'), 'txBox') === '1';
  if (!drawing && !placeholder && !textBox) {
    // New text in a drawn shape is centered, as PowerPoint does for shapes it creates.
    bodyPr.setAttribute('rtlCol', '0');
    bodyPr.setAttribute('anchor', 'ctr');
    const pPr = a('pPr');
    pPr.setAttribute('algn', 'ctr');
    p.appendChild(pPr);
  }
  body.appendChild(bodyPr);
  body.appendChild(a('lstStyle'));
  body.appendChild(p);
  if (drawing) host.insertBefore(body, host.firstChild);
  else host.insertBefore(body, child(host, 'extLst') ?? null);
  return body;
}

/**
 * Text of a DrawingML paragraph as the slide model shows it: line breaks as "\n", and slide number
 * fields as the slide's number (their stored text is only a placeholder such as "‹#›").
 */
function paragraphText(p: Element, slideNumber: number): string {
  return children(p)
    .map((el) => {
      if (el.localName === 'br') return '\n';
      if (el.localName === 'fld' && attr(el, 'type') === 'slidenum') return String(slideNumber);
      return el.localName === 'r' || el.localName === 'fld' ? (child(el, 't')?.textContent ?? '') : '';
    })
    .join('');
}

/**
 * Sets the paragraphs of a text body. Paragraphs whose text is unchanged are left as they are
 * (matched by a longest common subsequence, so inserting or deleting a line keeps the others).
 * Changed paragraphs keep their `a:pPr` and `a:endParaRPr`, and take the formatting of their first
 * run; new paragraphs copy the properties of the paragraph before them.
 */
function writeParagraphs(doc: Document, body: Element, values: string[], slideNumber: number): void {
  const old = children(body, 'p');
  const oldText = old.map((p) => paragraphText(p, slideNumber));
  const next = values.map((v) => v.replace(/\r\n|\r/g, '\n'));

  // Longest common subsequence of unchanged paragraphs.
  const lcs = Array.from({ length: old.length + 1 }, () => new Array<number>(next.length + 1).fill(0));
  for (let i = old.length - 1; i >= 0; i--) {
    for (let j = next.length - 1; j >= 0; j--) {
      lcs[i][j] = oldText[i] === next[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1]);
    }
  }
  // The plan: unchanged paragraphs are kept; between them, old paragraphs are rewritten in order
  // and the rest dropped, or new ones added.
  type Step = { keep: Element } | { rewrite: Element; value: string } | { add: string };
  const plan: Step[] = [];
  const gapOld: Element[] = [];
  const gapNew: string[] = [];
  const flush = () => {
    // Within a gap, the old and new paragraphs are paired in order; when their counts differ, the
    // unpaired ones (added or dropped) go where the pairs are most alike.
    const pairs = Math.min(gapOld.length, gapNew.length);
    let best = 0;
    let bestScore = -1;
    for (let offset = 0; offset <= Math.abs(gapOld.length - gapNew.length); offset++) {
      let score = 0;
      for (let k = 0; k < pairs; k++) {
        const [a, b] = gapOld.length > gapNew.length ? [k + offset, k] : [k, k + offset];
        score += similarity(oldText[old.indexOf(gapOld[a])], gapNew[b]);
      }
      if (score > bestScore) [best, bestScore] = [offset, score];
    }
    if (gapNew.length >= gapOld.length) {
      gapNew.forEach((value, k) => {
        const old = k >= best && k - best < pairs ? gapOld[k - best] : undefined;
        plan.push(old ? { rewrite: old, value } : { add: value });
      });
    } else {
      gapNew.forEach((value, k) => plan.push({ rewrite: gapOld[k + best], value }));
    }
    gapOld.length = 0;
    gapNew.length = 0;
  };
  for (let i = 0, j = 0; i < old.length || j < next.length; ) {
    if (i < old.length && j < next.length && oldText[i] === next[j] && lcs[i][j] === lcs[i + 1][j + 1] + 1) {
      flush();
      plan.push({ keep: old[i++] });
      j++;
    } else if (j < next.length && (i >= old.length || lcs[i][j + 1] >= lcs[i + 1][j])) {
      gapNew.push(next[j++]);
    } else {
      gapOld.push(old[i++]);
    }
  }
  flush();
  if (plan.some((step) => 'rewrite' in step && children(step.rewrite, 'fld').length)) {
    throw new EditError('A paragraph with a field (slide number, date) cannot be edited here yet.');
  }

  const result: Element[] = [];
  for (const step of plan) {
    if ('keep' in step) result.push(step.keep);
    else if ('rewrite' in step) {
      rewriteParagraph(doc, step.rewrite, step.value);
      result.push(step.rewrite);
    } else {
      result.push(newParagraph(doc, result[result.length - 1] ?? old[0], step.add));
    }
  }
  for (const p of old) body.removeChild(p);
  for (const p of result) body.appendChild(p);
}

/** Length of the common prefix plus the common suffix. */
function similarity(a: string, b: string): number {
  let prefix = 0;
  while (prefix < a.length && prefix < b.length && a[prefix] === b[prefix]) prefix++;
  let suffix = 0;
  while (suffix < a.length - prefix && suffix < b.length - prefix && a[a.length - 1 - suffix] === b[b.length - 1 - suffix]) suffix++;
  return prefix + suffix;
}

/** Replaces the runs of a paragraph, keeping its properties and end-of-paragraph properties. */
function rewriteParagraph(doc: Document, p: Element, value: string): void {
  const rPr = templateRunProps(doc, p);
  for (const el of children(p)) if (['r', 'br', 'fld'].includes(el.localName ?? '')) p.removeChild(el);
  let end = child(p, 'endParaRPr');
  if (!end && !value && rPr) {
    // Keeps an emptied paragraph at the size of the text it had.
    end = renamed(doc, rPr, 'endParaRPr');
    p.appendChild(end);
  }
  appendRuns(doc, p, rPr, value, end ?? null);
}

function newParagraph(doc: Document, like: Element | undefined, value: string): Element {
  const p = doc.createElementNS(DRAWING_NS, like?.prefix ? `${like.prefix}:p` : 'a:p');
  const pPr = child(like, 'pPr');
  if (pPr) p.appendChild(pPr.cloneNode(true));
  const rPr = like ? templateRunProps(doc, like) : undefined;
  const end = child(like, 'endParaRPr')?.cloneNode(true) as Element | undefined;
  if (end) p.appendChild(end);
  appendRuns(doc, p, rPr, value, end ?? null);
  return p;
}

/** Formatting for new text: the first run's, else the paragraph end's, else a language tag. */
function templateRunProps(doc: Document, p: Element): Element | undefined {
  const first = children(p).find((el) => ['r', 'fld', 'br'].includes(el.localName ?? '') && child(el, 'rPr'));
  const source = child(first, 'rPr') ?? child(p, 'endParaRPr');
  if (!source) return undefined;
  const rPr = renamed(doc, source, 'rPr');
  rPr.removeAttribute('err');
  return rPr;
}

function appendRuns(doc: Document, p: Element, rPr: Element | undefined, value: string, before: Element | null): void {
  const text = value.replace(INVALID_XML_CHARS, (ch) => (ch === '\v' ? ch : ''));
  const make = (name: string) => createLike(doc, p, name);
  let template = rPr;
  if (!template && text) {
    template = make('rPr');
    template.setAttribute('lang', CJK.test(text) ? 'ja-JP' : 'en-US');
  }
  text.split(LINE_BREAK).forEach((segment, k) => {
    if (k > 0) {
      const br = make('br');
      if (template) br.appendChild(template.cloneNode(true));
      p.insertBefore(br, before);
    }
    if (!segment) return;
    const r = make('r');
    if (template) r.appendChild(template.cloneNode(true));
    const t = make('t');
    t.appendChild(doc.createTextNode(segment));
    r.appendChild(t);
    p.insertBefore(r, before);
  });
}

/** A copy of `el` (attributes and children) under another local name in the same namespace. */
function renamed(doc: Document, el: Element, localName: string): Element {
  const copy = createLike(doc, el, localName);
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes.item(i)!;
    copy.setAttributeNS(a.namespaceURI, a.name, a.value);
  }
  for (const c of children(el)) copy.appendChild(c.cloneNode(true));
  return copy;
}

// ---------------------------------------------------------------------------------------------
// Helpers

function insideAlternateContent(el: Element): boolean {
  for (let node = el.parentNode as Element | null; node; node = node.parentNode as Element | null) {
    if (node.localName === 'AlternateContent') return true;
  }
  return false;
}

function relType(el: Element): string {
  const type = attr(el, 'Type') ?? '';
  return type.slice(type.lastIndexOf('/') + 1);
}

function setRelId(el: Element, rId: string): void {
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes.item(i)!;
    if (a.localName === 'id' && a.namespaceURI && /relationships$/.test(a.namespaceURI)) {
      a.value = rId;
      return;
    }
  }
  el.setAttributeNS(REL_TYPES, 'r:id', rId);
}

/** Slide ids must be unique and at least 256. */
function nextSlideId(presDoc: Document): number {
  const ids = children(child(presDoc.documentElement, 'sldIdLst'), 'sldId').map((el) => Number(attr(el, 'id')) || 0);
  return Math.max(255, ...ids) + 1;
}

/** The entry of a slide in the presentation's section list (PowerPoint 2010 extension), if any. */
function sectionEntry(presDoc: Document, slideId: string | undefined): Element | undefined {
  if (!slideId) return undefined;
  return descendants(presDoc.documentElement, 'section')
    .flatMap((section) => children(child(section, 'sldIdLst'), 'sldId'))
    .find((el) => attr(el, 'id') === slideId);
}

function randomId(): number {
  return Math.floor(Math.random() * 0xffff_ffff);
}
