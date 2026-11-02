// SpreadsheetML reader and editor: workbook structure, shared strings, styles, lazily parsed
// sheets, and edits applied as minimal changes to the parts they touch.
import { EditError } from '../errors';
import { applyTint, hexToRgb, INDEXED_COLORS, parseTheme, rgbToCss, type Theme } from '../color';
import { dirname, OoxmlPackage, relsPathOf } from '../package';
import { attr, boolAttr, child, children, descendants, numAttr, path, relAttr, text, toggle, type Element } from '../xml';
import { BUILTIN_FORMATS, formatValue, isDateFormat } from './numfmt';
import { columnIndex, parseRange, rowIndex } from './address';
import { numberEditText, parseInput, stringEditText } from './input';
import { MAX_COLS, MAX_ROWS, renameSheetInFormula, shiftFormula, shiftSqref, translateFormula, type Shift } from './refs';
import {
  arrayRanges, createLike, expandSharedFormulas, insertBefore, isBlankCell, setText, sharedGroup, SheetCursor, shiftSheet,
  updateDimension, writeCell,
} from './sheet-edit';
import type {
  CellAddress, CellData, CellKind, CellStyle, EditOp, MergeRange, SheetData, SheetInfo, WorkbookInfo,
} from './types';

export { EditError };

/** What changed, so views can refresh: individual cells, or the workbook structure. */
export type EditChange = { kind: 'cells'; sheet: number; cells: CellAddress[] } | { kind: 'structure'; activate?: number };

export interface AppliedEdit {
  label: string;
  change: EditChange;
  /** Restores the state before the edit; valid only while later edits have been reverted first. */
  revert(): Promise<EditChange>;
}

export interface CellPatch {
  cells: CellData[];
  /** Addresses that no longer have a cell. */
  removed: CellAddress[];
  styles: Record<number, CellStyle>;
}

type EditCell = CellAddress & { input: string };

const CONTENT_TYPES = '[Content_Types].xml';
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_TYPES = RELATIONSHIPS_NS;
const WORKSHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const PUSHED_OFF = 'Cells with data would be pushed off the end of the sheet.';
/** Workbook children that follow calcPr, in schema order. */
const AFTER_CALC_PR = [
  'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr',
  'webPublishObjects', 'extLst',
];

function newSheetXml(defaultRowHeight: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${RELATIONSHIPS_NS}">` +
    '<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    `<sheetFormatPr defaultRowHeight="${defaultRowHeight}"/><sheetData/>` +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
  );
}

/** The prefix a <sheet> element uses for its relationship id attribute (usually "r"). */
function relIdPrefix(el: Element): string | undefined {
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes.item(i);
    if (a?.localName === 'id' && a.namespaceURI === RELATIONSHIPS_NS) return a.prefix ?? undefined;
  }
  return undefined;
}

interface SheetEntry extends SheetInfo {
  partPath: string;
  /** The <sheet> element in workbook.xml. */
  element: Element;
}

type SharedFormulas = Map<string, { formula: string; r: number; c: number }>;

interface Xf {
  element: Element;
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment?: Element;
}

interface Styles {
  /** styles.xml, when the workbook has one. */
  path?: string;
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
  /** Whether anything was edited since the workbook was opened. */
  private modified = false;
  /** Worksheet parts whose cells changed, so their dimension is recomputed on save. */
  private readonly editedSheets = new Set<string>();

  private constructor(
    private readonly pkg: OoxmlPackage,
    private readonly workbookPath: string,
    private sheets: SheetEntry[],
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
    const sheets = await readSheets(pkg, workbookPath);

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
    return new Workbook(pkg, workbookPath, sheets, sharedStrings, styles, theme, date1904);
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

    const usedStyles: Record<number, CellStyle> = { 0: this.cellStyle(0) };

    const cells: CellData[] = [];
    const shared: SharedFormulas = new Map();
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
        const cell = this.readCell(cellEl, r, c, xf, shared);
        if (!cell) continue;
        usedStyles[xf] ??= this.cellStyle(xf);
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

  /** Current values of the given cells (for updating a view after an edit). */
  async cells(index: number, addresses: CellAddress[]): Promise<CellPatch> {
    const entry = this.entry(index);
    const doc = await this.pkg.xml(entry.partPath);
    const root = doc?.documentElement;
    if (!doc || !root) throw new EditError(`Sheet "${entry.name}" could not be read.`);
    const cursor = new SheetCursor(doc);
    // Shared formula masters, so the formula text of their dependents can be shown.
    const shared: SharedFormulas = new Map();
    for (const f of descendants(child(root, 'sheetData'), 'f')) {
      const ref = attr(f.parentNode as Element, 'r') ?? '';
      if (attr(f, 't') === 'shared' && attr(f, 'ref') && text(f)) {
        shared.set(attr(f, 'si') ?? '', { formula: text(f), r: rowIndex(ref), c: columnIndex(ref) });
      }
    }
    const patch: CellPatch = { cells: [], removed: [], styles: {} };
    for (const { r, c } of addresses) {
      const el = cursor.cell(r, c, false);
      const xf = el ? (numAttr(el, 's') ?? 0) : 0;
      const cell = el && this.readCell(el, r, c, xf, new Map(shared));
      if (cell) {
        patch.cells.push(cell);
        patch.styles[xf] ??= this.cellStyle(xf);
      } else {
        patch.removed.push({ r, c });
      }
    }
    return patch;
  }

  // --- Editing -------------------------------------------------------------------------------

  /** Applies one edit. The result can revert it; re-applying the same op redoes it. */
  async apply(op: EditOp): Promise<AppliedEdit> {
    switch (op.type) {
      case 'setCells':
        return this.setCells(op.sheet, op.cells, op.cells.length === 1 ? 'Edit cell' : 'Edit cells');
      case 'clearRange': {
        const doc = await this.pkg.xml(this.entry(op.sheet).partPath);
        const targets = new SheetCursor(doc!).cellsIn(op.range).flatMap((el) => {
          if (!children(el).length) return [];
          const ref = attr(el, 'r') ?? '';
          return [{ r: rowIndex(ref), c: columnIndex(ref), input: '' }];
        });
        return this.setCells(op.sheet, targets, 'Clear cells', op.range);
      }
      case 'insertRows':
      case 'deleteRows':
      case 'insertCols':
      case 'deleteCols': {
        const axis = op.type.endsWith('Rows') ? 'row' : 'col';
        const count = Math.max(1, Math.floor(op.count)) * (op.type.startsWith('insert') ? 1 : -1);
        const label = `${op.type.startsWith('insert') ? 'Insert' : 'Delete'} ${axis === 'row' ? 'rows' : 'columns'}`;
        return this.structural(label, () => this.shift(op.sheet, { axis, at: op.at, count }));
      }
      case 'addSheet':
        return this.structural('Add sheet', () => this.addSheet(op.name));
      case 'renameSheet':
        return this.structural('Rename sheet', () => this.renameSheet(op.sheet, op.name));
      case 'deleteSheet':
        return this.structural('Delete sheet', () => this.deleteSheet(op.sheet));
    }
  }

  /** Serializes the workbook, including the bookkeeping Excel expects after edits. */
  async save(): Promise<Uint8Array> {
    if (this.modified) {
      for (const partPath of this.editedSheets) {
        if (this.pkg.has(partPath)) updateDimension(await this.pkg.edit(partPath));
      }
      await this.dropCalcChain();
      // Cached formula results may be stale; ask Excel to recalculate everything on open.
      const root = (await this.pkg.edit(this.workbookPath)).documentElement!;
      let calcPr = child(root, 'calcPr');
      if (!calcPr) {
        calcPr = createLike(root.ownerDocument!, root, 'calcPr');
        insertBefore(root, calcPr, AFTER_CALC_PR);
      }
      calcPr.setAttribute('fullCalcOnLoad', '1');
    }
    return this.pkg.save();
  }

  private entry(index: number): SheetEntry {
    const entry = this.sheets[index];
    if (!entry) throw new EditError(`Sheet ${index} does not exist.`);
    return entry;
  }

  private async setCells(index: number, targets: EditCell[], label: string, clearing?: MergeRange): Promise<AppliedEdit> {
    const entry = this.entry(index);
    const doc = await this.pkg.edit(entry.partPath);
    const root = doc.documentElement!;
    const cursor = new SheetCursor(doc);

    for (const range of arrayRanges(root)) {
      const touched = targets.some(({ r, c }) => r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2);
      const whole = clearing && clearing.r1 <= range.r1 && clearing.c1 <= range.c1 && clearing.r2 >= range.r2 && clearing.c2 >= range.c2;
      if (touched && !whole) throw new EditError('Part of an array formula cannot be changed.');
    }

    const saved = new Map<number, { r: number; c: number; prev: Element | null }>();
    const record = (r: number, c: number, el: Element | undefined) => {
      const key = r * MAX_COLS + c;
      if (!saved.has(key)) saved.set(key, { r, c, prev: el ? (el.cloneNode(true) as Element) : null });
    };
    // Editing one cell of a shared formula group would orphan the others; make them ordinary formulas.
    const groups = new Set(
      targets.flatMap(({ r, c }) => {
        const el = cursor.cell(r, c, false);
        return (el && sharedGroup(el)) ?? [];
      }),
    );
    if (groups.size) {
      expandSharedFormulas(root, groups, (el) => {
        const ref = attr(el, 'r') ?? '';
        record(rowIndex(ref), columnIndex(ref), el);
      });
    }

    for (const { r, c, input } of targets) {
      const existing = cursor.cell(r, c, false);
      record(r, c, existing);
      if (!existing && input === '') continue;
      const cell = existing ?? cursor.cell(r, c, true)!;
      const xf = numAttr(cell, 's') ?? 0;
      const value = parseInput(input, this.numberFormat(xf), this.date1904);
      writeCell(doc, cell, value);
      if (value.kind === 'number' && value.numFmtId !== undefined) {
        const next = await this.withNumberFormat(xf, value.numFmtId);
        if (next !== xf) cell.setAttribute('s', String(next));
      }
      if (value.kind === 'empty' && isBlankCell(cell)) cell.parentNode!.removeChild(cell);
    }
    this.markEdited(entry.partPath);

    const changes = [...saved.values()];
    const change: EditChange = { kind: 'cells', sheet: index, cells: changes.map(({ r, c }) => ({ r, c })) };
    return {
      label,
      change,
      revert: async () => {
        const current = new SheetCursor(await this.pkg.edit(entry.partPath));
        for (const { r, c, prev } of [...changes].reverse()) current.put(r, c, prev);
        return change;
      },
    };
  }

  /** The cell format equal to `xf` but with another number format, added to styles.xml if needed. */
  private async withNumberFormat(xf: number, numFmtId: number): Promise<number> {
    const styles = this.styles;
    const base = styles.cellXfs[xf];
    if (!styles.path || !base || base.numFmtId === numFmtId) return xf;
    const signature = (x: Xf) =>
      ['fontId', 'fillId', 'borderId', 'xfId'].map((a) => attr(x.element, a) ?? '0').join() + children(x.element).map(String).join('');
    const wanted = signature(base);
    const found = styles.cellXfs.findIndex((x) => x.numFmtId === numFmtId && signature(x) === wanted);
    if (found >= 0) return found;

    await this.pkg.edit(styles.path);
    const element = base.element.cloneNode(true) as Element;
    element.setAttribute('numFmtId', String(numFmtId));
    element.setAttribute('applyNumberFormat', '1');
    const list = base.element.parentNode as Element;
    list.appendChild(element);
    styles.cellXfs.push({ ...base, element, numFmtId });
    if (list.hasAttribute('count')) list.setAttribute('count', String(styles.cellXfs.length));
    return styles.cellXfs.length - 1;
  }

  /**
   * Runs a structural edit with the package journal on: undo restores every touched part.
   * `run` returns the sheet to show afterwards, if it should change.
   */
  private async structural(label: string, run: () => Promise<number | undefined>): Promise<AppliedEdit> {
    this.pkg.begin();
    let activate: number | undefined;
    try {
      activate = await run();
    } catch (e) {
      this.pkg.restore(this.pkg.commit());
      this.sheets = await readSheets(this.pkg, this.workbookPath);
      throw e;
    }
    const snapshot = this.pkg.commit();
    this.sheets = await readSheets(this.pkg, this.workbookPath);
    this.modified = true;
    return {
      label,
      change: { kind: 'structure', activate },
      revert: async () => {
        this.pkg.restore(snapshot);
        this.sheets = await readSheets(this.pkg, this.workbookPath);
        return { kind: 'structure' };
      },
    };
  }

  private async shift(index: number, shift: Shift): Promise<undefined> {
    const entry = this.entry(index);
    const max = shift.axis === 'row' ? MAX_ROWS : MAX_COLS;
    if (shift.at < 0 || shift.at >= max) throw new EditError('The position is outside the sheet.');
    if (shift.count < 0) shift.count = Math.max(shift.count, shift.at - max);
    const doc = await this.pkg.edit(entry.partPath);
    const root = doc.documentElement!;

    if (shift.count > 0) {
      const limit = max - shift.count;
      for (const row of children(child(root, 'sheetData'), 'row')) {
        if (shift.axis === 'row' && Number(attr(row, 'r')) - 1 >= limit && children(row, 'c').length) throw new EditError(PUSHED_OFF);
        if (shift.axis === 'col' && children(row, 'c').some((c) => columnIndex(attr(c, 'r') ?? '') >= limit)) throw new EditError(PUSHED_OFF);
      }
    }

    // Tables (ListObjects) must keep their header row and columns intact.
    const tables: { part: string; ref: string }[] = [];
    for (const rel of await this.pkg.relsByType(entry.partPath, 'table')) {
      const ref = attr((await this.pkg.xml(rel.target))?.documentElement, 'ref');
      const range = ref && parseRange(ref);
      if (!range) continue;
      const [lo, hi] = shift.axis === 'row' ? [range.r1, range.r2] : [range.c1, range.c2];
      const d2 = shift.at - shift.count - 1;
      const blocked =
        shift.count > 0
          ? shift.axis === 'col' && shift.at > lo && shift.at <= hi
          : shift.axis === 'col'
            ? shift.at <= hi && d2 >= lo
            : shift.at <= lo && d2 >= lo;
      if (blocked) throw new EditError('Rows or columns that cut through a table cannot be inserted or deleted.');
      tables.push({ part: rel.target, ref });
    }

    expandSharedFormulas(root);
    shiftSheet(doc, shift);
    await this.rewriteFormulas((formula, home) => shiftFormula(formula, entry.name, home, shift), entry.partPath);
    for (const table of tables) {
      const next = shiftSqref(table.ref, shift);
      if (!next || next === table.ref) continue;
      const tableRoot = (await this.pkg.edit(table.part)).documentElement!;
      tableRoot.setAttribute('ref', next);
      const filter = child(tableRoot, 'autoFilter');
      const filterRef = attr(filter, 'ref');
      const nextFilter = filterRef && shiftSqref(filterRef, shift);
      if (filter && nextFilter) filter.setAttribute('ref', nextFilter);
    }
    this.markEdited(entry.partPath);
    return undefined;
  }

  private async addSheet(requested: string | undefined): Promise<number> {
    const wbDoc = await this.pkg.edit(this.workbookPath);
    const root = wbDoc.documentElement!;
    const sheetsEl = child(root, 'sheets')!;
    const allSheets = children(sheetsEl, 'sheet');
    let name = requested?.trim();
    if (!name) {
      let n = this.sheets.length + 1;
      while (allSheets.some((s) => attr(s, 'name')?.toLowerCase() === `sheet${n}`)) n++;
      name = `Sheet${n}`;
    }
    this.validateName(name);

    const dir = dirname(this.workbookPath);
    let k = 1;
    while (this.pkg.has(`${dir}/worksheets/sheet${k}.xml`)) k++;
    const partPath = `${dir}/worksheets/sheet${k}.xml`;

    const relsDoc = await this.pkg.edit(relsPathOf(this.workbookPath));
    const relsRoot = relsDoc.documentElement!;
    const ids = children(relsRoot, 'Relationship').map((el) => Number(/^rId(\d+)$/.exec(attr(el, 'Id') ?? '')?.[1] ?? 0));
    const rId = `rId${Math.max(0, ...ids) + 1}`;
    const rel = createLike(relsDoc, relsRoot, 'Relationship');
    rel.setAttribute('Id', rId);
    rel.setAttribute('Type', `${REL_TYPES}/worksheet`);
    rel.setAttribute('Target', `worksheets/sheet${k}.xml`);
    relsRoot.appendChild(rel);

    await this.addContentType(`/${partPath}`, WORKSHEET_CONTENT_TYPE);
    const firstSheet = this.sheets[0] && (await this.pkg.xml(this.sheets[0].partPath))?.documentElement;
    const rowHeight = attr(child(firstSheet, 'sheetFormatPr'), 'defaultRowHeight') ?? '15';
    await this.pkg.setText(partPath, newSheetXml(rowHeight));

    const sheetEl = createLike(wbDoc, sheetsEl, 'sheet');
    sheetEl.setAttribute('name', name);
    sheetEl.setAttribute('sheetId', String(Math.max(0, ...allSheets.map((s) => numAttr(s, 'sheetId') ?? 0)) + 1));
    const relPrefix = allSheets.map((s) => relIdPrefix(s)).find(Boolean) ?? 'r';
    sheetEl.setAttributeNS(RELATIONSHIPS_NS, `${relPrefix}:id`, rId);
    sheetsEl.appendChild(sheetEl);
    return this.sheets.length;
  }

  private async renameSheet(index: number, name: string): Promise<undefined> {
    const entry = this.entry(index);
    name = name.trim();
    if (name === entry.name) return undefined;
    this.validateName(name, entry.element);
    await this.pkg.edit(this.workbookPath);
    entry.element.setAttribute('name', name);
    const from = entry.name;
    await this.rewriteFormulas((formula) => renameSheetInFormula(formula, from, name));
    for (const part of this.pkg.parts().filter((p) => /pivotCacheDefinition[^/]*\.xml$/.test(p))) {
      const sources = descendants((await this.pkg.xml(part))?.documentElement, 'worksheetSource').filter(
        (el) => attr(el, 'sheet')?.toLowerCase() === from.toLowerCase(),
      );
      if (!sources.length) continue;
      await this.pkg.edit(part);
      for (const el of sources) el.setAttribute('sheet', name);
    }
    return undefined;
  }

  private async deleteSheet(index: number): Promise<number> {
    const entry = this.entry(index);
    if (!this.sheets.some((s) => s !== entry && s.state === 'visible')) {
      throw new EditError('A workbook must contain at least one visible worksheet.');
    }
    const root = (await this.pkg.edit(this.workbookPath)).documentElement!;
    const sheetsEl = child(root, 'sheets')!;
    const position = children(sheetsEl, 'sheet').indexOf(entry.element);
    const rId = relAttr(entry.element, 'id');
    sheetsEl.removeChild(entry.element);

    const relsRoot = (await this.pkg.edit(relsPathOf(this.workbookPath))).documentElement!;
    for (const rel of children(relsRoot, 'Relationship')) if (attr(rel, 'Id') === rId) relsRoot.removeChild(rel);
    await this.removeContentType(`/${entry.partPath}`);
    await this.pkg.remove(entry.partPath);
    await this.pkg.remove(relsPathOf(entry.partPath));

    // Names scoped to the sheet go away; later sheets' scopes move down by one.
    const names = child(root, 'definedNames');
    for (const dn of children(names, 'definedName')) {
      const local = numAttr(dn, 'localSheetId');
      if (local === position) names!.removeChild(dn);
      else if (local !== undefined && local > position) dn.setAttribute('localSheetId', String(local - 1));
    }
    if (names && !children(names, 'definedName').length) root.removeChild(names);
    const remaining = children(sheetsEl, 'sheet').length;
    for (const view of children(child(root, 'bookViews'), 'workbookView')) {
      for (const a of ['activeTab', 'firstSheet']) {
        const value = numAttr(view, a);
        if (value !== undefined && (value > position || value >= remaining)) view.setAttribute(a, String(Math.max(0, value - 1)));
      }
    }
    await this.rewriteFormulas((formula) => renameSheetInFormula(formula, entry.name, undefined));
    return Math.min(index, this.sheets.length - 2);
  }

  /**
   * Rewrites formulas on every worksheet, in defined names and in chart series. `home` is true for
   * formulas on the sheet at `homePath`, where unqualified references point.
   */
  private async rewriteFormulas(fn: (formula: string, home: boolean) => string, homePath?: string): Promise<void> {
    const rewrite = async (partPath: string, elements: Element[], home: boolean) => {
      const changes = elements.flatMap((el) => {
        const before = text(el);
        const after = before ? fn(before, home) : before;
        return after !== before ? [[el, after] as const] : [];
      });
      if (!changes.length) return false;
      await this.pkg.edit(partPath);
      for (const [el, value] of changes) setText(el, value);
      return true;
    };
    for (const sheet of this.sheets) {
      const root = (await this.pkg.xml(sheet.partPath))?.documentElement;
      if (!root) continue;
      if (await rewrite(sheet.partPath, descendants(child(root, 'sheetData'), 'f'), sheet.partPath === homePath)) {
        this.markEdited(sheet.partPath);
      }
    }
    const wbRoot = (await this.pkg.xml(this.workbookPath))?.documentElement;
    await rewrite(this.workbookPath, children(child(wbRoot, 'definedNames'), 'definedName'), false);
    for (const part of this.pkg.parts().filter((p) => /^xl\/charts\/chart[^/]*\.xml$/.test(p))) {
      await rewrite(part, descendants((await this.pkg.xml(part))?.documentElement, 'f'), false);
    }
  }

  private validateName(name: string, except?: Element): void {
    if (!name || name.length > 31) throw new EditError('Sheet names must be 1 to 31 characters long.');
    if (/[:\\/?*[\]]/.test(name) || name.startsWith("'") || name.endsWith("'")) {
      throw new EditError("Sheet names cannot contain : \\ / ? * [ ] or begin or end with '.");
    }
    if (name.toLowerCase() === 'history') throw new EditError('"History" is a reserved sheet name.');
    const sheetsEl = this.sheets[0]?.element.parentNode as Element | undefined;
    const taken = children(sheetsEl, 'sheet').some((s) => s !== except && attr(s, 'name')?.toLowerCase() === name.toLowerCase());
    if (taken) throw new EditError(`A sheet named "${name}" already exists.`);
  }

  private markEdited(partPath: string): void {
    this.modified = true;
    this.editedSheets.add(partPath);
  }

  private async addContentType(partName: string, contentType: string): Promise<void> {
    const root = (await this.pkg.edit(CONTENT_TYPES)).documentElement!;
    const override = createLike(root.ownerDocument!, root, 'Override');
    override.setAttribute('PartName', partName);
    override.setAttribute('ContentType', contentType);
    root.appendChild(override);
  }

  private async removeContentType(partName: string): Promise<void> {
    const root = (await this.pkg.xml(CONTENT_TYPES))?.documentElement;
    const matches = children(root, 'Override').filter((el) => attr(el, 'PartName')?.toLowerCase() === partName.toLowerCase());
    if (!matches.length) return;
    await this.pkg.edit(CONTENT_TYPES);
    for (const el of matches) root!.removeChild(el);
  }

  /** The calculation chain lists formula cells; a stale one makes Excel repair the file. */
  private async dropCalcChain(): Promise<void> {
    const rel = (await this.pkg.relsByType(this.workbookPath, 'calcChain'))[0];
    if (!rel) return;
    const relsRoot = (await this.pkg.edit(relsPathOf(this.workbookPath))).documentElement!;
    for (const el of children(relsRoot, 'Relationship')) if (attr(el, 'Id') === rel.id) relsRoot.removeChild(el);
    await this.removeContentType(`/${rel.target}`);
    await this.pkg.remove(rel.target);
  }

  private readCell(el: Element, r: number, c: number, xf: number, shared?: SharedFormulas): CellData | undefined {
    const type = attr(el, 't') ?? 'n';
    const vEl = child(el, 'v');
    const fEl = child(el, 'f');
    let f = fEl && text(fEl) ? text(fEl) : undefined;
    if (fEl && shared && attr(fEl, 't') === 'shared') {
      const si = attr(fEl, 'si') ?? '';
      const master = shared.get(si);
      if (f && attr(fEl, 'ref')) shared.set(si, { formula: f, r, c });
      else if (master) f = translateFormula(master.formula, r - master.r, c - master.c);
    }
    const format = this.numberFormat(xf);
    let v: string;
    let t: CellKind;
    let input: string | undefined;
    switch (type) {
      case 's':
        v = this.sharedStrings[Number(text(vEl))] ?? '';
        t = 's';
        input = stringEditText(v, format);
        break;
      case 'str':
        v = decodeEscapes(text(vEl));
        t = 's';
        break;
      case 'inlineStr':
        v = richText(child(el, 'is'));
        t = 's';
        input = stringEditText(v, format);
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
        v = Number.isFinite(num) ? formatValue(num, format, this.date1904) : raw;
        t = isDateFormat(format) ? 'd' : 'n';
        if (Number.isFinite(num)) input = numberEditText(num, format, this.date1904);
      }
    }
    if (v === '' && xf === 0 && !f) return undefined;
    const cell: CellData = { r, c, v, s: xf, t };
    if (f) {
      cell.f = f;
      cell.input = `=${f}`;
    } else if (input !== undefined && input !== v) {
      cell.input = input;
    }
    return cell;
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

async function readSheets(pkg: OoxmlPackage, workbookPath: string): Promise<SheetEntry[]> {
  const root = (await pkg.xml(workbookPath))?.documentElement;
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
      element: el,
    });
  }
  return sheets;
}

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
    path: root ? partPath : undefined,
    numFmts,
    fonts: children(child(root, 'fonts'), 'font'),
    fills: children(child(root, 'fills'), 'fill'),
    borders: children(child(root, 'borders'), 'border'),
    cellXfs: children(child(root, 'cellXfs'), 'xf').map((xf) => ({
      element: xf,
      numFmtId: numAttr(xf, 'numFmtId') ?? 0,
      fontId: numAttr(xf, 'fontId') ?? 0,
      fillId: numAttr(xf, 'fillId') ?? 0,
      borderId: numAttr(xf, 'borderId') ?? 0,
      alignment: child(xf, 'alignment'),
    })),
  };
}
