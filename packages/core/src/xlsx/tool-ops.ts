// Spreadsheet edits as language models express them (A1 addresses, sheet names, 1-based rows)
// translated into the editor's EditOps, with validation messages a model can act on.
import { columnIndex, parseRange } from './address';
import { MAX_COLS, MAX_ROWS } from './refs';
import type { EditOp, SheetInfo } from './types';
import { EditError } from './workbook';

export type ToolEdit =
  /** Map of A1 address → typed input ("42", "2026/1/15", "=SUM(A1:A3)", "" to clear). */
  | { type: 'setCells'; sheet?: string; cells: Record<string, string | number | boolean | null> }
  | { type: 'clearRange'; sheet?: string; range: string }
  | { type: 'insertRows' | 'deleteRows'; sheet?: string; row: number; count?: number }
  | { type: 'insertColumns' | 'deleteColumns'; sheet?: string; column: string; count?: number }
  | { type: 'addSheet'; name?: string }
  | { type: 'renameSheet'; sheet: string; name: string }
  | { type: 'deleteSheet'; sheet: string };

/** Converts one tool edit using the workbook's current sheets (indices shift as sheets change). */
export function toEditOp(edit: ToolEdit, sheets: SheetInfo[]): EditOp {
  const sheetIndex = (name: string | undefined): number => {
    if (name === undefined || name === '') {
      const first = sheets.find((s) => s.state === 'visible') ?? sheets[0];
      if (!first) throw new EditError('The workbook has no sheets.');
      return first.index;
    }
    const found = sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
    if (!found) throw new EditError(`No sheet named "${name}". Sheets: ${sheets.map((s) => s.name).join(', ')}`);
    return found.index;
  };
  const count = (n: number | undefined) => {
    const value = n ?? 1;
    if (!Number.isInteger(value) || value < 1) throw new EditError(`count must be a positive integer (got ${n}).`);
    return value;
  };

  switch (edit.type) {
    case 'setCells': {
      const entries = Object.entries(edit.cells ?? {});
      if (!entries.length) throw new EditError('setCells needs at least one cell.');
      return {
        type: 'setCells',
        sheet: sheetIndex(edit.sheet),
        cells: entries.map(([address, value]) => {
          const range = parseRange(address.replace(/\$/g, ''));
          if (!range || range.r1 !== range.r2 || range.c1 !== range.c2 || !/^\$?[A-Za-z]{1,3}\$?\d+$/.test(address)) {
            throw new EditError(`"${address}" is not a single cell address like "B3".`);
          }
          if (range.r1 >= MAX_ROWS || range.c1 >= MAX_COLS) throw new EditError(`"${address}" is outside the sheet.`);
          return { r: range.r1, c: range.c1, input: value === null || value === undefined ? '' : String(value) };
        }),
      };
    }
    case 'clearRange': {
      const range = parseRange((edit.range ?? '').replace(/\$/g, ''));
      if (!range) throw new EditError(`"${edit.range}" is not an A1 range like "A1:C10".`);
      return { type: 'clearRange', sheet: sheetIndex(edit.sheet), range };
    }
    case 'insertRows':
    case 'deleteRows': {
      if (!Number.isInteger(edit.row) || edit.row < 1 || edit.row > MAX_ROWS) throw new EditError(`row must be a 1-based row number (got ${edit.row}).`);
      return { type: edit.type, sheet: sheetIndex(edit.sheet), at: edit.row - 1, count: count(edit.count) };
    }
    case 'insertColumns':
    case 'deleteColumns': {
      const at = /^[A-Za-z]{1,3}$/.test(edit.column ?? '') ? columnIndex(edit.column) : -1;
      if (at < 0 || at >= MAX_COLS) throw new EditError(`column must be a column letter like "C" (got ${edit.column}).`);
      return { type: edit.type === 'insertColumns' ? 'insertCols' : 'deleteCols', sheet: sheetIndex(edit.sheet), at, count: count(edit.count) };
    }
    case 'addSheet':
      return { type: 'addSheet', name: edit.name };
    case 'renameSheet':
      return { type: 'renameSheet', sheet: sheetIndex(edit.sheet), name: edit.name };
    case 'deleteSheet':
      return { type: 'deleteSheet', sheet: sheetIndex(edit.sheet) };
    default:
      throw new EditError(`Unknown edit type "${(edit as { type?: string }).type}".`);
  }
}

/** A short human-readable description, for confirmations and results. */
export function describeEdit(edit: ToolEdit): string {
  const on = 'sheet' in edit && edit.sheet ? ` on "${edit.sheet}"` : '';
  switch (edit.type) {
    case 'setCells': {
      const cells = Object.entries(edit.cells ?? {});
      const shown = cells.slice(0, 5).map(([a, v]) => `${a} = ${JSON.stringify(v)}`).join(', ');
      return `Set ${cells.length} cell(s)${on}: ${shown}${cells.length > 5 ? ', …' : ''}`;
    }
    case 'clearRange':
      return `Clear ${edit.range}${on}`;
    case 'insertRows':
    case 'deleteRows':
      return `${edit.type === 'insertRows' ? 'Insert' : 'Delete'} ${edit.count ?? 1} row(s) at row ${edit.row}${on}`;
    case 'insertColumns':
    case 'deleteColumns':
      return `${edit.type === 'insertColumns' ? 'Insert' : 'Delete'} ${edit.count ?? 1} column(s) at column ${edit.column}${on}`;
    case 'addSheet':
      return `Add sheet${edit.name ? ` "${edit.name}"` : ''}`;
    case 'renameSheet':
      return `Rename sheet "${edit.sheet}" to "${edit.name}"`;
    case 'deleteSheet':
      return `Delete sheet "${edit.sheet}"`;
    default:
      return 'Unknown edit';
  }
}
