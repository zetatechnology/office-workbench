// Word document → Markdown: headings, lists, tables and plain text runs, in reading order.
// Meant for reading and diffing, so layout and character formatting are dropped.
import { OoxmlPackage } from '../package';
import { attr, child, children, numAttr, path, type Element } from '../xml';
import { escapeCell, markdownTable } from './markdown';

interface StyleInfo {
  name: string;
  basedOn?: string;
  outlineLevel?: number;
  numId?: string;
  ilvl?: number;
}

/** Styles and numbering, for telling headings and list items apart. */
export interface DocxContext {
  styles: Map<string, StyleInfo>;
  /** numId → level → numFmt ("bullet", "decimal", ...). */
  numbering: Map<string, Map<number, string>>;
}

type Context = DocxContext;

/** Reads the styles and numbering of the main document part. */
export async function readContext(pkg: OoxmlPackage, mainPath: string): Promise<DocxContext> {
  const rels = [...(await pkg.rels(mainPath)).values()];
  const target = (type: string) => rels.find((r) => r.type === type && !r.external)?.target;
  return { styles: await readStyles(pkg, target('styles')), numbering: await readNumbering(pkg, target('numbering')) };
}

export async function docxToMarkdown(data: Uint8Array): Promise<string> {
  const pkg = await OoxmlPackage.open(data);
  const mainPath = await pkg.mainPart();
  const body = child((await pkg.xml(mainPath))?.documentElement, 'body');
  if (!body) throw new Error('This file is not a Word document.');
  const rels = [...(await pkg.rels(mainPath)).values()];
  const target = (type: string) => rels.filter((r) => r.type === type && !r.external).map((r) => r.target);

  const ctx = await readContext(pkg, mainPath);
  const blocks = renderBlocks(children(body), ctx);

  // Headers and footers, once each (sections often repeat the same part).
  const extras: string[] = [];
  for (const [type, label] of [['header', 'Header'], ['footer', 'Footer']] as const) {
    const seen = new Set<string>();
    for (const part of target(type)) {
      const text = renderBlocks(children((await pkg.xml(part))?.documentElement), ctx).join('\n\n').trim();
      if (text && !seen.has(text)) {
        seen.add(text);
        extras.push(`**${label}:** ${text.replace(/\n+/g, ' ')}`);
      }
    }
  }
  for (const [type, label] of [['footnotes', 'Footnotes'], ['endnotes', 'Endnotes']] as const) {
    const notes = children((await pkg.xml(target(type)[0] ?? ''))?.documentElement).filter((n) => !attr(n, 'type'));
    const lines = notes.flatMap((n) => {
      const text = renderBlocks(children(n), ctx).join(' ').trim();
      return text ? [`${attr(n, 'id')}. ${text}`] : [];
    });
    if (lines.length) extras.push(`**${label}:**\n\n${lines.join('\n')}`);
  }
  return [...blocks, ...(extras.length ? ['---', ...extras] : [])].join('\n\n') + '\n';
}

function renderBlocks(elements: Element[], ctx: Context): string[] {
  const out: string[] = [];
  let listRun: string[] = [];
  const flushList = () => {
    if (listRun.length) out.push(listRun.join('\n'));
    listRun = [];
  };
  for (const el of elements) {
    switch (el.localName) {
      case 'p': {
        const { text, list } = renderParagraph(el, ctx);
        if (list !== undefined) {
          listRun.push(list + text);
          continue;
        }
        flushList();
        if (text.trim()) out.push(text);
        break;
      }
      case 'tbl':
        flushList();
        out.push(renderTable(el, ctx));
        break;
      case 'sdt':
        flushList();
        out.push(...renderBlocks(children(child(el, 'sdtContent')), ctx));
        break;
      default:
        break;
    }
  }
  flushList();
  return out;
}

/** Text of a paragraph, prefixed for headings; `list` is the list marker when it is a list item. */
function renderParagraph(p: Element, ctx: Context): { text: string; list?: string } {
  const text = paragraphText(p).replace(/\s+$/, '');
  const info = paragraphInfo(p, ctx);
  if (info.heading && text.trim()) return { text: `${'#'.repeat(Math.min(info.heading, 6))} ${text.trim()}` };
  if (info.list && text.trim()) {
    return { text: text.trim(), list: `${'  '.repeat(info.list.level)}${info.list.ordered ? '1.' : '-'} ` };
  }
  return { text };
}

/** Style name, heading level (from the style name or outline level) and list membership of a paragraph. */
export function paragraphInfo(
  p: Element,
  ctx: DocxContext,
): { style?: string; heading?: number; list?: { ordered: boolean; level: number } } {
  const pPr = child(p, 'pPr');
  const styleId = attr(child(pPr, 'pStyle'), 'val');
  const style = resolveStyle(ctx.styles, styleId);

  const outline = numAttr(child(pPr, 'outlineLvl'), 'val') ?? style.outlineLevel;
  const heading = /^heading (\d)$/i.exec(style.name ?? '');
  const level = heading ? Number(heading[1]) : outline !== undefined && outline < 9 ? outline + 1 : /^title$/i.test(style.name ?? '') ? 1 : 0;
  if (level) return { style: style.name, heading: level };

  const numPr = child(pPr, 'numPr');
  const numId = attr(child(numPr, 'numId'), 'val') ?? style.numId;
  const ilvl = numAttr(child(numPr, 'ilvl'), 'val') ?? style.ilvl ?? 0;
  if (numId && numId !== '0') {
    const format = ctx.numbering.get(numId)?.get(ilvl) ?? 'bullet';
    return { style: style.name, list: { ordered: !(format === 'bullet' || format === 'none'), level: ilvl } };
  }
  return { style: style.name };
}

/** Concatenated run text, skipping deleted revisions, drawings and field instructions. */
export function paragraphText(p: Element): string {
  let out = '';
  const walk = (el: Element) => {
    for (const node of children(el)) {
      switch (node.localName) {
        case 't':
          out += node.textContent ?? '';
          break;
        case 'tab':
        case 'ptab':
          out += '\t';
          break;
        case 'br':
        case 'cr':
          out += attr(node, 'type') === 'page' ? '' : '\n';
          break;
        case 'noBreakHyphen':
          out += '-';
          break;
        case 'sym':
          out += String.fromCharCode(parseInt(attr(node, 'char') ?? '0', 16) || 0x20);
          break;
        case 'del':
        case 'drawing':
        case 'pict':
        case 'object':
        case 'instrText':
        case 'pPr':
        case 'rPr':
          break;
        default:
          walk(node);
      }
    }
  };
  walk(p);
  return out;
}

function renderTable(tbl: Element, ctx: Context): string {
  const rows = children(tbl, 'tr').map((tr) =>
    children(tr, 'tc').flatMap((tc) => {
      const text = escapeCell(renderBlocks(children(tc), ctx).join('\n'));
      const span = numAttr(path(tc, 'tcPr', 'gridSpan'), 'val') ?? 1;
      // Vertically merged continuation cells are left empty, like the visible table.
      const continued = path(tc, 'tcPr', 'vMerge') && attr(path(tc, 'tcPr', 'vMerge'), 'val') !== 'restart';
      return [continued ? '' : text, ...Array<string>(span - 1).fill('')];
    }),
  );
  return markdownTable(rows);
}

function resolveStyle(styles: Map<string, StyleInfo>, id: string | undefined): Partial<StyleInfo> {
  const result: Partial<StyleInfo> = {};
  for (let depth = 0; id && depth < 20; depth++) {
    const style = styles.get(id);
    if (!style) break;
    result.name ??= style.name;
    result.outlineLevel ??= style.outlineLevel;
    result.numId ??= style.numId;
    result.ilvl ??= style.ilvl;
    id = style.basedOn;
  }
  return result;
}

async function readStyles(pkg: OoxmlPackage, partPath: string | undefined): Promise<Map<string, StyleInfo>> {
  const styles = new Map<string, StyleInfo>();
  if (!partPath) return styles;
  for (const el of children((await pkg.xml(partPath))?.documentElement, 'style')) {
    const pPr = child(el, 'pPr');
    styles.set(attr(el, 'styleId') ?? '', {
      name: attr(child(el, 'name'), 'val') ?? '',
      basedOn: attr(child(el, 'basedOn'), 'val'),
      outlineLevel: numAttr(child(pPr, 'outlineLvl'), 'val'),
      numId: attr(path(pPr, 'numPr', 'numId'), 'val'),
      ilvl: numAttr(path(pPr, 'numPr', 'ilvl'), 'val'),
    });
  }
  return styles;
}

async function readNumbering(pkg: OoxmlPackage, partPath: string | undefined): Promise<Context['numbering']> {
  const result: Context['numbering'] = new Map();
  if (!partPath) return result;
  const root = (await pkg.xml(partPath))?.documentElement;
  const abstract = new Map<string, Map<number, string>>();
  for (const el of children(root, 'abstractNum')) {
    const levels = new Map<number, string>();
    for (const lvl of children(el, 'lvl')) levels.set(numAttr(lvl, 'ilvl') ?? 0, attr(child(lvl, 'numFmt'), 'val') ?? 'decimal');
    abstract.set(attr(el, 'abstractNumId') ?? '', levels);
  }
  for (const num of children(root, 'num')) {
    const levels = abstract.get(attr(child(num, 'abstractNumId'), 'val') ?? '');
    if (levels) result.set(attr(num, 'numId') ?? '', levels);
  }
  return result;
}
