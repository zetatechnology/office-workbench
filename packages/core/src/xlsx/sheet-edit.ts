// In-place edits of a worksheet DOM. Elements are changed individually so that everything the
// editor does not understand (extLst, x14 extensions, unknown attributes) is preserved.
import type { Document, Element } from '../xml';
import { attr, child, children, createLike, descendants, text } from '../xml';
import { columnIndex, columnName, parseRange, rowIndex } from './address';
import type { ParsedInput } from './input';
import { MAX_COLS, MAX_ROWS, shiftFormula, shiftInterval, shiftSqref, translateFormula, type Shift } from './refs';
import type { MergeRange } from './types';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';

export function cellRef(r: number, c: number): string {
  return `${columnName(c)}${r + 1}`;
}

export { createLike };

/** Inserts `node` as a child of `parent` before the first child named in `before`, or at the end. */
export function insertBefore(parent: Element, node: Element, before: string[]): void {
  const next = children(parent).find((el) => before.includes(el.localName ?? ''));
  parent.insertBefore(node, next ?? null);
}

/** Row and cell lookup over a worksheet's sheetData, with elements created on demand. */
export class SheetCursor {
  private readonly rows = new Map<number, Element>();
  private readonly sheetData: Element;

  constructor(private readonly doc: Document) {
    const root = doc.documentElement!;
    let sheetData = child(root, 'sheetData');
    if (!sheetData) {
      sheetData = createLike(doc, root, 'sheetData');
      insertBefore(root, sheetData, AFTER_SHEET_DATA);
    }
    this.sheetData = sheetData;
    // Excel always writes r attributes, but they are optional; fill them in so edits can address rows.
    let rowCursor = -1;
    for (const row of children(sheetData, 'row')) {
      const r = attr(row, 'r') ? Number(attr(row, 'r')) - 1 : rowCursor + 1;
      rowCursor = r;
      if (!attr(row, 'r')) row.setAttribute('r', String(r + 1));
      this.rows.set(r, row);
      let colCursor = -1;
      for (const cell of children(row, 'c')) {
        const ref = attr(cell, 'r');
        const c = ref ? columnIndex(ref) : colCursor + 1;
        colCursor = c;
        if (!ref) cell.setAttribute('r', cellRef(r, c));
      }
    }
  }

  get rowElements(): Element[] {
    return children(this.sheetData, 'row');
  }

  row(r: number, create: boolean): Element | undefined {
    let row = this.rows.get(r);
    if (row || !create) return row;
    row = createLike(this.doc, this.sheetData, 'row');
    row.setAttribute('r', String(r + 1));
    const next = children(this.sheetData, 'row').find((el) => Number(attr(el, 'r')) - 1 > r);
    this.sheetData.insertBefore(row, next ?? null);
    this.rows.set(r, row);
    return row;
  }

  cell(r: number, c: number, create: boolean): Element | undefined {
    const row = this.row(r, create);
    if (!row) return undefined;
    let next: Element | undefined;
    for (const el of children(row, 'c')) {
      const col = columnIndex(attr(el, 'r') ?? '');
      if (col === c) return el;
      if (col > c) {
        next = el;
        break;
      }
    }
    if (!create) return undefined;
    const cell = createLike(this.doc, row, 'c');
    cell.setAttribute('r', cellRef(r, c));
    // Cells inherit the row's format when the row has one.
    if (attr(row, 'customFormat') === '1' && attr(row, 's')) cell.setAttribute('s', attr(row, 's')!);
    row.insertBefore(cell, next ?? null);
    // `spans` is only an optimization hint; drop it rather than keep a stale one.
    row.removeAttribute('spans');
    return cell;
  }

  /** Cell elements inside a range, in document order. */
  cellsIn(range: MergeRange): Element[] {
    const out: Element[] = [];
    for (const [r, row] of this.rows) {
      if (r < range.r1 || r > range.r2) continue;
      for (const cell of children(row, 'c')) {
        const c = columnIndex(attr(cell, 'r') ?? '');
        if (c >= range.c1 && c <= range.c2) out.push(cell);
      }
    }
    return out;
  }

  /** Replaces (or removes, when `replacement` is null) the cell at an address. */
  put(r: number, c: number, replacement: Element | null): void {
    const current = this.cell(r, c, replacement !== null);
    if (!current) return;
    if (replacement) current.parentNode!.replaceChild(this.doc.importNode(replacement, true), current);
    else current.parentNode!.removeChild(current);
  }
}

/** Worksheet children that follow sheetData, in schema order. */
const AFTER_SHEET_DATA = [
  'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate',
  'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks',
  'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties',
  'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF',
  'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
];

/** Replaces the value of a cell element, keeping its style and any extensions. */
export function writeCell(doc: Document, cell: Element, value: ParsedInput): void {
  for (const el of children(cell)) {
    if (el.localName === 'f' || el.localName === 'v' || el.localName === 'is') cell.removeChild(el);
  }
  cell.removeAttribute('t');
  const add = (name: string, content: string) => {
    const el = createLike(doc, cell, name);
    el.appendChild(doc.createTextNode(content));
    insertBefore(cell, el, ['extLst']);
    return el;
  };
  switch (value.kind) {
    case 'empty':
      break;
    case 'formula':
      add('f', value.formula);
      break;
    case 'number':
      add('v', String(value.value));
      break;
    case 'bool':
      cell.setAttribute('t', 'b');
      add('v', value.value ? '1' : '0');
      break;
    case 'error':
      cell.setAttribute('t', 'e');
      add('v', value.code);
      break;
    case 'string': {
      // Inline strings avoid rewriting the shared string table; Excel converts them on its next save.
      cell.setAttribute('t', 'inlineStr');
      const is = createLike(doc, cell, 'is');
      const t = createLike(doc, cell, 't');
      if (/^\s|\s$|\n/.test(value.text)) t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(value.text));
      is.appendChild(t);
      insertBefore(cell, is, ['extLst']);
      break;
    }
  }
}

/** True when the element holds nothing worth keeping (no value, formula or format). */
export function isBlankCell(cell: Element): boolean {
  return !children(cell).length && !attr(cell, 's') && !attr(cell, 'cm') && !attr(cell, 'vm');
}

/**
 * Turns shared formulas into ordinary ones, either all groups or only the given `si` ids.
 * `before` is called with each cell element just before it changes.
 */
export function expandSharedFormulas(root: Element, only?: Set<string>, before?: (cell: Element) => void): void {
  const formulas = descendants(child(root, 'sheetData'), 'f').filter(
    (f) => attr(f, 't') === 'shared' && (!only || only.has(attr(f, 'si') ?? '')),
  );
  const masters = new Map<string, { formula: string; r: number; c: number }>();
  for (const f of formulas) {
    const ref = attr(f, 'ref');
    const formula = text(f);
    const cellAddr = attr(f.parentNode as Element, 'r') ?? '';
    if (ref && formula) masters.set(attr(f, 'si') ?? '', { formula, r: rowIndex(cellAddr), c: columnIndex(cellAddr) });
  }
  for (const f of formulas) {
    const master = masters.get(attr(f, 'si') ?? '');
    const cell = f.parentNode as Element;
    before?.(cell);
    if (master) {
      const addr = attr(cell, 'r') ?? '';
      const formula = translateFormula(master.formula, rowIndex(addr) - master.r, columnIndex(addr) - master.c);
      while (f.firstChild) f.removeChild(f.firstChild);
      f.appendChild(f.ownerDocument!.createTextNode(formula));
    }
    f.removeAttribute('t');
    f.removeAttribute('si');
    f.removeAttribute('ref');
  }
}

/** The `si` group of a shared formula cell, if any. */
export function sharedGroup(cell: Element): string | undefined {
  const f = child(cell, 'f');
  return f && attr(f, 't') === 'shared' ? attr(f, 'si') : undefined;
}

/** Ranges covered by multi-cell array formulas. */
export function arrayRanges(root: Element): MergeRange[] {
  return descendants(child(root, 'sheetData'), 'f').flatMap((f) => {
    const range = attr(f, 't') === 'array' ? parseRange(attr(f, 'ref') ?? '') : undefined;
    return range && (range.r1 !== range.r2 || range.c1 !== range.c2) ? [range] : [];
  });
}

/** Moves the cells and cell-anchored features of the sheet whose rows/columns change. */
export function shiftSheet(doc: Document, shift: Shift): void {
  const root = doc.documentElement!;
  const cursor = new SheetCursor(doc);
  const max = shift.axis === 'row' ? MAX_ROWS : MAX_COLS;
  const move = (i: number) => shiftInterval(i, i, shift.at, shift.count, max)?.[0];

  for (const row of cursor.rowElements) {
    const r = Number(attr(row, 'r')) - 1;
    if (shift.axis === 'row') {
      const nr = move(r);
      if (nr === undefined) {
        row.parentNode!.removeChild(row);
        continue;
      }
      if (nr !== r) {
        row.setAttribute('r', String(nr + 1));
        for (const cell of children(row, 'c')) cell.setAttribute('r', cellRef(nr, columnIndex(attr(cell, 'r') ?? '')));
      }
    } else {
      row.removeAttribute('spans');
      for (const cell of children(row, 'c')) {
        const c = columnIndex(attr(cell, 'r') ?? '');
        const nc = move(c);
        if (nc === undefined) row.removeChild(cell);
        else if (nc !== c) cell.setAttribute('r', cellRef(r, nc));
      }
    }
  }

  // Array formula ranges (shared formulas are expected to be expanded beforehand).
  for (const f of descendants(child(root, 'sheetData'), 'f')) {
    const ref = attr(f, 'ref');
    const next = ref && shiftSqref(ref, shift);
    if (next) f.setAttribute('ref', next);
  }

  // Column widths and formats.
  if (shift.axis === 'col') {
    for (const col of children(child(root, 'cols'), 'col')) {
      const moved = shiftInterval(Number(attr(col, 'min')) - 1, Number(attr(col, 'max')) - 1, shift.at, shift.count, MAX_COLS);
      if (!moved) col.parentNode!.removeChild(col);
      else {
        col.setAttribute('min', String(moved[0] + 1));
        col.setAttribute('max', String(moved[1] + 1));
      }
    }
    const cols = child(root, 'cols');
    if (cols && !children(cols, 'col').length) root.removeChild(cols);
  }

  // Features anchored to ranges: removed when their whole range is deleted.
  const sqref = (el: Element, name: string, removeIfGone: boolean): boolean => {
    const value = attr(el, name);
    if (!value) return true;
    const next = shiftSqref(value, shift);
    if (next) el.setAttribute(name, next);
    else if (removeIfGone) el.parentNode?.removeChild(el);
    else el.removeAttribute(name);
    return !!next;
  };
  const list = (container: string, item: string, attrName: string) => {
    const parent = child(root, container);
    if (!parent) return;
    for (const el of children(parent, item)) {
      if (sqref(el, attrName, true)) shiftChildFormulas(el, shift);
    }
    const remaining = children(parent, item).length;
    if (!remaining) root.removeChild(parent);
    else if (parent.hasAttribute('count')) parent.setAttribute('count', String(remaining));
  };
  list('mergeCells', 'mergeCell', 'ref');
  list('dataValidations', 'dataValidation', 'sqref');
  list('hyperlinks', 'hyperlink', 'ref');
  list('protectedRanges', 'protectedRange', 'sqref');
  list('ignoredErrors', 'ignoredError', 'sqref');
  for (const cf of children(root, 'conditionalFormatting')) {
    if (sqref(cf, 'sqref', true)) shiftChildFormulas(cf, shift);
  }
  const autoFilter = child(root, 'autoFilter');
  if (autoFilter) sqref(autoFilter, 'ref', true);
  for (const view of children(child(root, 'sheetViews'), 'sheetView')) {
    const pane = child(view, 'pane');
    if (pane) sqref(pane, 'topLeftCell', false);
    if (pane && attr(pane, 'state')?.startsWith('frozen')) shiftFrozenPane(view, pane, shift);
    for (const selection of children(view, 'selection')) {
      if (!sqref(selection, 'activeCell', false) || !sqref(selection, 'sqref', false)) {
        selection.setAttribute('activeCell', 'A1');
        selection.setAttribute('sqref', 'A1');
      }
    }
  }
}

/**
 * Frozen rows/columns grow or shrink with the shift, and the scrolling pane must start right after
 * them; Excel refuses files where topLeftCell lies inside the frozen area.
 */
function shiftFrozenPane(view: Element, pane: Element, shift: Shift): void {
  const splitAttr = shift.axis === 'row' ? 'ySplit' : 'xSplit';
  const split = Number(attr(pane, splitAttr) ?? 0);
  if (split > 0) {
    const moved = shiftInterval(0, split - 1, shift.at, shift.count, shift.axis === 'row' ? MAX_ROWS : MAX_COLS);
    if (moved) pane.setAttribute(splitAttr, String(moved[1] + 1));
    else pane.removeAttribute(splitAttr);
  }
  const ySplit = Number(attr(pane, 'ySplit') ?? 0), xSplit = Number(attr(pane, 'xSplit') ?? 0);
  if (!ySplit && !xSplit) {
    view.removeChild(pane);
    children(view, 'selection').forEach((s, i) => (i ? view.removeChild(s) : s.removeAttribute('pane')));
    return;
  }
  const top = attr(pane, 'topLeftCell') ?? '';
  pane.setAttribute('topLeftCell', cellRef(Math.max(rowIndex(top), ySplit), Math.max(columnIndex(top), xSplit)));
  const active = ySplit && xSplit ? 'bottomRight' : ySplit ? 'bottomLeft' : 'topRight';
  if (attr(pane, 'activePane') !== active) {
    pane.setAttribute('activePane', active);
    children(view, 'selection').forEach((s, i) => (i ? view.removeChild(s) : s.setAttribute('pane', active)));
  }
}

/** Shifts `<formula>`, `<formula1>`, `<formula2>` children (conditional formats, validations). */
function shiftChildFormulas(el: Element, shift: Shift): void {
  for (const name of ['formula', 'formula1', 'formula2']) {
    for (const f of descendants(el, name)) {
      const next = shiftFormula(text(f), '', true, shift);
      if (next !== text(f)) setText(f, next);
    }
  }
}

export function setText(el: Element, value: string): void {
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(el.ownerDocument!.createTextNode(value));
}

/** Recomputes `<dimension ref>` from the cells present. */
export function updateDimension(doc: Document): void {
  const root = doc.documentElement!;
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  for (const row of children(child(root, 'sheetData'), 'row')) {
    for (const cell of children(row, 'c')) {
      const ref = attr(cell, 'r') ?? '';
      const r = rowIndex(ref), c = columnIndex(ref);
      if (r < r1) r1 = r;
      if (r > r2) r2 = r;
      if (c < c1) c1 = c;
      if (c > c2) c2 = c;
    }
  }
  const ref = r2 < 0 ? 'A1' : r1 === r2 && c1 === c2 ? cellRef(r1, c1) : `${cellRef(r1, c1)}:${cellRef(r2, c2)}`;
  let dimension = child(root, 'dimension');
  if (!dimension) {
    dimension = createLike(doc, root, 'dimension');
    const sheetPr = child(root, 'sheetPr');
    root.insertBefore(dimension, sheetPr ? sheetPr.nextSibling : root.firstChild);
  }
  dimension.setAttribute('ref', ref);
}
