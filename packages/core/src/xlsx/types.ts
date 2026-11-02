export interface WorkbookInfo {
  sheets: SheetInfo[];
}

export interface SheetInfo {
  index: number;
  name: string;
  state: 'visible' | 'hidden' | 'veryHidden';
}

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** CSS color. */
  color?: string;
  /** Points. */
  fontSize?: number;
  fontFamily?: string;
  background?: string;
  hAlign?: 'left' | 'center' | 'right' | 'justify' | 'fill' | 'centerContinuous' | 'distributed';
  vAlign?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrap?: boolean;
  indent?: number;
  rotation?: number;
  /** CSS border shorthands. */
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
}

/** n: number, s: string, b: boolean, e: error, d: date-formatted number. */
export type CellKind = 'n' | 's' | 'b' | 'e' | 'd';

export interface CellData {
  /** 0-based row. */
  r: number;
  /** 0-based column. */
  c: number;
  /** Display text after number formatting. */
  v: string;
  /** Cell format (xf) index; key into SheetData.styles. */
  s: number;
  t: CellKind;
  /** Formula text without the leading "=". */
  f?: string;
  /** Text for the cell editor when it differs from `v` (raw number, y/m/d date, "=formula"). */
  input?: string;
}

export interface MergeRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface SheetData {
  index: number;
  name: string;
  /** Number of rows/columns that contain data, styles or merges. */
  rowCount: number;
  colCount: number;
  /** Pixels. */
  defaultColWidth: number;
  defaultRowHeight: number;
  /** 0-based index → pixels (0 when hidden). */
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  cells: CellData[];
  /** Resolved styles of the cell formats used on the sheet, keyed by xf index. */
  styles: Record<number, CellStyle>;
  merges: MergeRange[];
  frozen?: { rows: number; cols: number };
  showGridLines: boolean;
  /** Human-readable list of content this viewer does not render (charts, images, ...). */
  unsupported: string[];
}

export interface CellAddress {
  r: number;
  c: number;
}

/** Edits applied by the extension host; each one is a single undo step. */
export type EditOp =
  /** `input` is what the user typed: a value, "=formula", or "" to clear the cell. */
  | { type: 'setCells'; sheet: number; cells: (CellAddress & { input: string })[] }
  | { type: 'clearRange'; sheet: number; range: MergeRange }
  | { type: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols'; sheet: number; at: number; count: number }
  | { type: 'addSheet'; name?: string }
  | { type: 'renameSheet'; sheet: number; name: string }
  | { type: 'deleteSheet'; sheet: number };
