// SpreadsheetML reader: workbook structure, shared strings, styles and lazily parsed sheets.
import { applyTint, hexToRgb, INDEXED_COLORS, parseTheme, rgbToCss, type Theme } from '../color';
import { OoxmlPackage } from '../package';
import { attr, boolAttr, child, children, descendants, numAttr, path, relAttr, text, toggle, type Element } from '../xml';
import { BUILTIN_FORMATS, formatValue, isDateFormat } from './numfmt';
import { columnIndex, parseRange } from './address';
import type { CellData, CellKind, CellStyle, MergeRange, SheetData, SheetInfo, WorkbookInfo } from './types';

interface SheetEntry extends SheetInfo {
  partPath: string;
}

interface Xf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment?: Element;
}

interface Styles {
  numFmts: Map<number, string>;
  fonts: Element[];
  fills: Element[];
  borders: Element[];
  cellXfs: Xf[];
}

const MAX_COLUMNS = 16_384;
/** Extra empty columns of width information sent beyond the used range. */
const COLUMN_PADDING = 100;

export class Workbook {
  private readonly styleCache = new Map<number, CellStyle>();

  private constructor(
    private readonly pkg: OoxmlPackage,
    private readonly sheets: SheetEntry[],
    private readonly sharedStrings: string[],
    private readonly styles: Styles,
    private readonly theme: Theme,
    private readonly date1904: boolean,
  ) {}

  static async open(data: Uint8Array): Promise<Workbook> {
    const pkg = await OoxmlPackage.open(data);
    const workbookPath = await pkg.mainPart();
    const doc = await pkg.xml(workbookPath);
    const root = doc?.documentElement;
    if (!root || root.localName !== 'workbook') throw new Error('This file is not an Excel workbook.');
    const rels = await pkg.rels(workbookPath);

    const sheets: SheetEntry[] = [];
    for (const el of children(child(root, 'sheets'), 'sheet')) {
      const rel = rels.get(relAttr(el, 'id') ?? '');
      // Chart sheets and dialog sheets are listed too; only worksheets are rendered.
      if (!rel || rel.type !== 'worksheet') continue;
      const state = attr(el, 'state');
      sheets.push({
        index: sheets.length,
        name: attr(el, 'name') ?? `Sheet${sheets.length + 1}`,
        state: state === 'hidden' || state === 'veryHidden' ? state : 'visible',
        partPath: rel.target,
      });
    }

    const relList = [...rels.values()];
    const sstPath = relList.find((r) => r.type === 'sharedStrings')?.target;
    const stylesPath = relList.find((r) => r.type === 'styles')?.target;
    const themePath = relList.find((r) => r.type === 'theme')?.target;
    const [sharedStrings, styles, theme] = await Promise.all([
      sstPath ? parseSharedStrings(pkg, sstPath) : Promise.resolve([]),
      parseStyles(pkg, stylesPath),
      themePath ? pkg.xml(themePath).then(parseTheme) : Promise.resolve(parseTheme(undefined)),
    ]);
    const date1904 = boolAttr(child(root, 'workbookPr'), 'date1904');
    return new Workbook(pkg, sheets, sharedStrings, styles, theme, date1904);
  }

  info(): WorkbookInfo {
    return { sheets: this.sheets.map(({ index, name, state }) => ({ index, name, state })) };
  }

  async sheet(index: number): Promise<SheetData> {
    const entry = this.sheets[index];
    if (!entry) throw new Error(`Sheet ${index} does not exist.`);
    const doc = await this.pkg.xml(entry.partPath);
    const root = doc?.documentElement;
    if (!root) throw new Error(`Sheet "${entry.name}" could not be read.`);

    const formatPr = child(root, 'sheetFormatPr');
    const dcw = numAttr(formatPr, 'defaultColWidth');
    const defaultColWidth = dcw ? Math.round(dcw * 7) : 64;
    const defaultRowHeight = Math.round((numAttr(formatPr, 'defaultRowHeight') ?? 15) * (4 / 3));

    const styleRemap = new Map<number, number>();
    const usedStyles: CellStyle[] = [];
    const styleIndex = (xf: number): number => {
      let i = styleRemap.get(xf);
      if (i === undefined) {
        i = usedStyles.length;
        usedStyles.push(this.cellStyle(xf));
        styleRemap.set(xf, i);
      }
      return i;
    };
    styleIndex(0);

    const cells: CellData[] = [];
    const rowHeights: Record<number, number> = {};
    let maxRow = -1;
    let maxCol = -1;
    let rowCursor = -1;
    for (const rowEl of children(child(root, 'sheetData'), 'row')) {
      const r = (numAttr(rowEl, 'r') ?? rowCursor + 2) - 1;
      rowCursor = r;
      if (boolAttr(rowEl, 'hidden')) rowHeights[r] = 0;
      else {
        const ht = numAttr(rowEl, 'ht');
        if (ht !== undefined && boolAttr(rowEl, 'customHeight', true)) rowHeights[r] = Math.round(ht * (4 / 3));
      }
      let colCursor = -1;
      for (let node = rowEl.firstChild; node; node = node.nextSibling) {
        if (node.nodeType !== 1 || (node as Element).localName !== 'c') continue;
        const cellEl = node as Element;
        const ref = attr(cellEl, 'r');
        const c = ref ? columnIndex(ref) : colCursor + 1;
        colCursor = c;
        const xf = numAttr(cellEl, 's') ?? 0;
        const cell = this.readCell(cellEl, r, c, xf);
        if (!cell) continue;
        cell.s = styleIndex(xf);
        cells.push(cell);
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      }
    }

    const merges: MergeRange[] = children(child(root, 'mergeCells'), 'mergeCell').flatMap((el) => {
      const range = parseRange(attr(el, 'ref') ?? '');
      if (!range) return [];
      maxRow = Math.max(maxRow, range.r2);
      maxCol = Math.max(maxCol, range.c2);
      return [range];
    });

    const colWidths: Record<number, number> = {};
    const colLimit = Math.min(MAX_COLUMNS, Math.max(maxCol + 1, 26) + COLUMN_PADDING);
    for (const col of children(child(root, 'cols'), 'col')) {
      const min = (numAttr(col, 'min') ?? 1) - 1;
      const max = Math.min((numAttr(col, 'max') ?? min + 1) - 1, colLimit);
      const width = boolAttr(col, 'hidden') ? 0 : Math.round((numAttr(col, 'width') ?? 0) * 7) || defaultColWidth;
      for (let c = min; c <= max; c++) colWidths[c] = width;
    }

    const view = path(root, 'sheetViews', 'sheetView');
    const pane = child(view, 'pane');
    const frozen =
      pane && (attr(pane, 'state') === 'frozen' || attr(pane, 'state') === 'frozenSplit')
        ? { rows: numAttr(pane, 'ySplit') ?? 0, cols: numAttr(pane, 'xSplit') ?? 0 }
        : undefined;

    return {
      index,
      name: entry.name,
      rowCount: maxRow + 1,
      colCount: maxCol + 1,
      defaultColWidth,
      defaultRowHeight,
      colWidths,
      rowHeights,
      cells,
      styles: usedStyles,
      merges,
      frozen,
      showGridLines: boolAttr(view, 'showGridLines', true),
      unsupported: await this.unsupportedContent(entry.partPath, root),
    };
  }

  private readCell(el: Element, r: number, c: number, xf: number): CellData | undefined {
    const type = attr(el, 't') ?? 'n';
    const vEl = child(el, 'v');
    const fEl = child(el, 'f');
    const f = fEl && text(fEl) ? text(fEl) : undefined;
    let v: string;
    let t: CellKind;
    switch (type) {
      case 's':
        v = this.sharedStrings[Number(text(vEl))] ?? '';
        t = 's';
        break;
      case 'str':
        v = decodeEscapes(text(vEl));
        t = 's';
        break;
      case 'inlineStr':
        v = richText(child(el, 'is'));
        t = 's';
        break;
      case 'b':
        v = text(vEl) === '1' ? 'TRUE' : 'FALSE';
        t = 'b';
        break;
      case 'e':
        v = text(vEl);
        t = 'e';
        break;
      case 'd':
        v = text(vEl);
        t = 'd';
        break;
      default: {
        const raw = text(vEl);
        if (raw === '') {
          v = '';
          t = 'n';
          break;
        }
        const num = Number(raw);
        const format = this.numberFormat(xf);
        v = Number.isFinite(num) ? formatValue(num, format, this.date1904) : raw;
        t = isDateFormat(format) ? 'd' : 'n';
      }
    }
    if (v === '' && xf === 0 && !f) return undefined;
    return f ? { r, c, v, s: 0, t, f } : { r, c, v, s: 0, t };
  }

  private numberFormat(xf: number): string {
    const id = this.styles.cellXfs[xf]?.numFmtId ?? 0;
    return this.styles.numFmts.get(id) ?? BUILTIN_FORMATS[id] ?? 'General';
  }

  private cellStyle(xfIndex: number): CellStyle {
    const cached = this.styleCache.get(xfIndex);
    if (cached) return cached;
    const xf = this.styles.cellXfs[xfIndex];
    const style: CellStyle = {};
    if (xf) {
      const font = this.styles.fonts[xf.fontId];
      if (font) {
        if (toggle(child(font, 'b'))) style.bold = true;
        if (toggle(child(font, 'i'))) style.italic = true;
        const u = child(font, 'u');
        if (u && attr(u, 'val') !== 'none') style.underline = true;
        if (toggle(child(font, 'strike'))) style.strike = true;
        const sz = numAttr(child(font, 'sz'), 'val');
        if (sz) style.fontSize = sz;
        const name = attr(child(font, 'name'), 'val');
        if (name) style.fontFamily = name;
        const color = this.color(child(font, 'color'));
        if (color) style.color = color;
      }
      const background = this.fillColor(this.styles.fills[xf.fillId]);
      if (background) style.background = background;
      const border = this.styles.borders[xf.borderId];
      if (border) {
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
          const css = this.borderCss(child(border, side) ?? child(border, side === 'left' ? 'start' : side === 'right' ? 'end' : side));
          if (css) style[`border${side[0].toUpperCase()}${side.slice(1)}` as 'borderTop'] = css;
        }
      }
      const al = xf.alignment;
      if (al) {
        const h = attr(al, 'horizontal');
        if (h && h !== 'general') style.hAlign = h as CellStyle['hAlign'];
        const v = attr(al, 'vertical');
        if (v) style.vAlign = v as CellStyle['vAlign'];
        if (boolAttr(al, 'wrapText')) style.wrap = true;
        const indent = numAttr(al, 'indent');
        if (indent) style.indent = indent;
        const rotation = numAttr(al, 'textRotation');
        if (rotation) style.rotation = rotation;
      }
    }
    this.styleCache.set(xfIndex, style);
    return style;
  }

  private color(el: Element | undefined): string | undefined {
    if (!el || boolAttr(el, 'auto')) return undefined;
    let hex: string | undefined;
    const rgb = attr(el, 'rgb');
    const theme = numAttr(el, 'theme');
    const indexed = numAttr(el, 'indexed');
    if (rgb) hex = rgb.length === 8 ? rgb.slice(2) : rgb;
    else if (theme !== undefined) hex = this.theme.colors[THEME_ORDER[theme]];
    else if (indexed !== undefined) hex = INDEXED_COLORS[indexed];
    if (!hex) return undefined;
    const tint = numAttr(el, 'tint');
    const color = hexToRgb(hex);
    return rgbToCss(tint ? applyTint(color, tint) : color);
  }

  private fillColor(fill: Element | undefined): string | undefined {
    const pattern = child(fill, 'patternFill');
    if (pattern) {
      const type = attr(pattern, 'patternType');
      if (!type || type === 'none') return undefined;
      return this.color(child(pattern, 'fgColor')) ?? (type === 'solid' ? undefined : this.color(child(pattern, 'bgColor')));
    }
    const stop = path(fill, 'gradientFill', 'stop');
    return stop ? this.color(child(stop, 'color')) : undefined;
  }

  private borderCss(el: Element | undefined): string | undefined {
    const style = attr(el, 'style');
    if (!el || !style || style === 'none') return undefined;
    const color = this.color(child(el, 'color')) ?? '#000000';
    const map: Record<string, string> = {
      thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed',
      dashDotDot: '1px dashed', double: '3px double', medium: '2px solid', mediumDashed: '2px dashed',
      mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed', thick: '3px solid',
    };
    return `${map[style] ?? '1px solid'} ${color}`;
  }

  private async unsupportedContent(partPath: string, root: Element): Promise<string[]> {
    const notes: string[] = [];
    const rels = await this.pkg.rels(partPath);
    for (const rel of rels.values()) {
      if (rel.type !== 'drawing') continue;
      const drawing = await this.pkg.xml(rel.target);
      const anchors = ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].reduce(
        (n, name) => n + descendants(drawing?.documentElement, name).length,
        0,
      );
      if (anchors) notes.push(`${anchors} drawing object(s) (charts, images or shapes)`);
    }
    const conditional = children(root, 'conditionalFormatting').length;
    if (conditional) notes.push('conditional formatting');
    if (children(root, 'legacyDrawing').length) notes.push('comments');
    return notes;
  }
}

/** SpreadsheetML theme color indices swap the first two pairs relative to the theme XML order. */
const THEME_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

async function parseSharedStrings(pkg: OoxmlPackage, partPath: string): Promise<string[]> {
  const doc = await pkg.xml(partPath);
  return children(doc?.documentElement, 'si').map(richText);
}

/** Concatenates the text runs of a string item, skipping phonetic (furigana) runs. */
function richText(el: Element | undefined): string {
  if (!el) return '';
  const t = child(el, 't');
  if (t) return decodeEscapes(text(t));
  return decodeEscapes(children(el, 'r').map((r) => text(child(r, 't'))).join(''));
}

function decodeEscapes(value: string): string {
  return value.includes('_x') ? value.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16))) : value;
}

async function parseStyles(pkg: OoxmlPackage, partPath: string | undefined): Promise<Styles> {
  const root = partPath ? (await pkg.xml(partPath))?.documentElement : undefined;
  const numFmts = new Map<number, string>();
  for (const el of children(child(root, 'numFmts'), 'numFmt')) {
    numFmts.set(numAttr(el, 'numFmtId') ?? -1, attr(el, 'formatCode') ?? 'General');
  }
  return {
    numFmts,
    fonts: children(child(root, 'fonts'), 'font'),
    fills: children(child(root, 'fills'), 'fill'),
    borders: children(child(root, 'borders'), 'border'),
    cellXfs: children(child(root, 'cellXfs'), 'xf').map((xf) => ({
      numFmtId: numAttr(xf, 'numFmtId') ?? 0,
      fontId: numAttr(xf, 'fontId') ?? 0,
      fillId: numAttr(xf, 'fillId') ?? 0,
      borderId: numAttr(xf, 'borderId') ?? 0,
      alignment: child(xf, 'alignment'),
    })),
  };
}
