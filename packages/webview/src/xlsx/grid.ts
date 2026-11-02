// Virtualized sheet grid: only the cells inside the viewport (plus a small overscan) exist in the DOM.
// Keyboard focus lives in a hidden textarea that doubles as the cell editor, so typing (including
// IME composition) starts editing the active cell directly.
import type { CellData, CellStyle, MergeRange, SheetData } from '@office-workbench/core';
import { columnName } from '@office-workbench/core/address';
import { el, escapeHtml } from '../common/host';

const HEADER_HEIGHT = 22;
const OVERSCAN = 4;
/** Empty rows/columns shown beyond the used range so the sheet does not end abruptly. */
const EXTRA_ROWS = 50;
const EXTRA_COLS = 10;
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export interface Selection {
  r: number;
  c: number;
}

export type Range = MergeRange;

export interface CellInput extends Selection {
  input: string;
}

export interface CellsPatch {
  cells: CellData[];
  removed: Selection[];
  styles: Record<number, CellStyle>;
}

export interface GridOptions {
  editable: boolean;
  /** `single` is true when the selection is one cell (or one merged area). */
  onSelect(cell: CellData | undefined, active: Selection, range: Range, single: boolean): void;
  /** Live text of the cell editor, or undefined when editing ends. */
  onEditorInput?(value: string | undefined): void;
  onEdit?(cells: CellInput[]): void;
  onClear?(range: Range): void;
  onContextMenu?(event: MouseEvent, kind: 'cell' | 'row' | 'col', range: Range): void;
}

type EditMode = 'enter' | 'edit';

export class SheetGrid {
  readonly element: HTMLElement;
  private readonly corner = el('div', 'grid-corner');
  private readonly colHeader = el('div', 'grid-col-header');
  private readonly rowHeader = el('div', 'grid-row-header');
  private readonly body = el('div', 'grid-body');
  private readonly canvas = el('div', 'grid-canvas');
  private readonly cellLayer = el('div', 'grid-cells');
  private readonly rangeBox = el('div', 'grid-range');
  private readonly selectionBox = el('div', 'grid-selection');
  private readonly editor = el('textarea', 'cell-editor');
  private readonly colHeaderLayer = el('div', 'grid-header-layer');
  private readonly rowHeaderLayer = el('div', 'grid-header-layer');
  private readonly styleSheet = el('style');

  private readonly rows: number;
  private readonly cols: number;
  private readonly rowOffsets: Float64Array;
  private readonly colOffsets: Float64Array;
  private readonly cells = new Map<number, CellData>();
  /** Covered (non-anchor) cells of merged ranges → their range. */
  private readonly covered = new Map<number, MergeRange>();
  private readonly anchors = new Map<number, MergeRange>();
  private readonly rowHeaderWidth: number;
  private frame = 0;
  /** The active cell; the selected range spans from it to `extent`. */
  private active: Selection = { r: 0, c: 0 };
  private extent: Selection = { r: 0, c: 0 };
  private editing?: { mode: EditMode; original: string };
  private dragging?: 'cell' | 'row' | 'col';

  constructor(private readonly sheet: SheetData, private readonly options: GridOptions) {
    this.rows = Math.min(MAX_ROWS, Math.max(sheet.rowCount + EXTRA_ROWS, 100));
    this.cols = Math.min(MAX_COLS, Math.max(sheet.colCount + EXTRA_COLS, 26));
    this.rowOffsets = offsets(this.rows, (r) => sheet.rowHeights[r] ?? sheet.defaultRowHeight);
    this.colOffsets = offsets(this.cols, (c) => sheet.colWidths[c] ?? sheet.defaultColWidth);
    for (const cell of sheet.cells) this.cells.set(key(cell.r, cell.c), cell);
    for (const m of sheet.merges) {
      this.anchors.set(key(m.r1, m.c1), m);
      for (let r = m.r1; r <= m.r2; r++) {
        for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) this.covered.set(key(r, c), m);
      }
    }
    this.rowHeaderWidth = Math.max(40, String(this.rows).length * 8 + 16);

    this.element = el('div', 'grid');
    this.element.classList.toggle('no-gridlines', !sheet.showGridLines);
    this.element.style.setProperty('--row-header-width', `${this.rowHeaderWidth}px`);
    this.element.style.setProperty('--header-height', `${HEADER_HEIGHT}px`);
    this.styleSheet.textContent = Object.entries(sheet.styles)
      .map(([i, s]) => `.s${i}{${styleCss(s)}}`)
      .join('\n');
    this.canvas.style.width = `${this.colOffsets[this.cols]}px`;
    this.canvas.style.height = `${this.rowOffsets[this.rows]}px`;
    this.editor.spellcheck = false;
    this.editor.setAttribute('aria-label', 'Cell editor');
    this.editor.readOnly = !options.editable;
    this.canvas.append(this.cellLayer, this.rangeBox, this.selectionBox, this.editor);
    this.body.append(this.canvas);
    this.colHeader.append(this.colHeaderLayer);
    this.rowHeader.append(this.rowHeaderLayer);
    this.element.append(this.styleSheet, this.corner, this.colHeader, this.rowHeader, this.body);

    this.body.addEventListener('scroll', () => this.schedule());
    new ResizeObserver(() => this.schedule()).observe(this.body);
    this.listen();
  }

  get scroll(): { top: number; left: number } {
    return { top: this.body.scrollTop, left: this.body.scrollLeft };
  }

  focus(): void {
    this.editor.focus({ preventScroll: true });
  }

  restore(scroll: { top: number; left: number } | undefined, selection: Selection | undefined): void {
    if (scroll) {
      this.body.scrollTop = scroll.top;
      this.body.scrollLeft = scroll.left;
    }
    this.select(selection?.r ?? 0, selection?.c ?? 0, { reveal: false });
    this.schedule();
  }

  /** Applies updated cells from the host. */
  applyPatch(patch: CellsPatch): void {
    const rules: string[] = [];
    for (const [i, style] of Object.entries(patch.styles)) {
      if (this.sheet.styles[Number(i)]) continue;
      this.sheet.styles[Number(i)] = style;
      rules.push(`.s${i}{${styleCss(style)}}`);
    }
    if (rules.length) this.styleSheet.textContent += `\n${rules.join('\n')}`;
    for (const { r, c } of patch.removed) this.cells.delete(key(r, c));
    for (const cell of patch.cells) this.cells.set(key(cell.r, cell.c), cell);
    if (!this.editing) this.notifySelect();
    this.schedule();
  }

  /** Starts editing the active cell, e.g. from the formula bar's state or F2. */
  beginEdit(mode: EditMode, text?: string): void {
    if (!this.options.editable) return;
    const cell = this.cells.get(key(this.active.r, this.active.c));
    const original = cell?.input ?? cell?.v ?? '';
    this.editing = { mode, original };
    this.editor.value = text ?? original;
    this.editor.className = `cell-editor active${cell ? ` s${cell.s}` : ''}`;
    this.placeEditor();
    this.reveal(this.active.r, this.active.c);
    this.focus();
    const end = this.editor.value.length;
    this.editor.setSelectionRange(end, end);
    this.options.onEditorInput?.(this.editor.value);
  }

  /** Ends editing; returns whether a value was committed. */
  endEdit(commit: boolean, toRange = false): boolean {
    if (!this.editing) return false;
    const value = this.editor.value;
    const changed = value !== this.editing.original;
    this.editing = undefined;
    this.editor.className = 'cell-editor';
    this.editor.value = '';
    this.placeEditor();
    this.options.onEditorInput?.(undefined);
    if (!commit || (!changed && !toRange)) return false;
    const { r1, c1, r2, c2 } = this.normalized();
    const targets: CellInput[] = [];
    if (toRange) {
      for (let r = r1; r <= Math.min(r2, this.rows - 1); r++) {
        for (let c = c1; c <= Math.min(c2, this.cols - 1); c++) if (!this.covered.has(key(r, c))) targets.push({ r, c, input: value });
      }
    } else {
      targets.push({ ...this.active, input: value });
    }
    this.options.onEdit?.(targets);
    return true;
  }

  /** Tab-separated text of the selection, as Excel puts it on the clipboard. */
  selectionText(): string {
    const { r1, c1 } = this.normalized();
    const r2 = Math.min(this.normalized().r2, Math.max(this.sheet.rowCount - 1, r1));
    const c2 = Math.min(this.normalized().c2, Math.max(this.sheet.colCount - 1, c1));
    const lines: string[] = [];
    for (let r = r1; r <= r2; r++) {
      if (this.rowOffsets[r + 1] === this.rowOffsets[r] && r1 !== r2) continue;
      const fields: string[] = [];
      for (let c = c1; c <= c2; c++) {
        const v = this.cells.get(key(r, c))?.v ?? '';
        fields.push(/[\t\n"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v);
      }
      lines.push(fields.join('\t'));
    }
    return lines.join('\r\n') + (lines.length > 1 ? '\r\n' : '');
  }

  /** Pastes tab-separated text at the active cell (repeating a single value over the selection). */
  paste(text: string): void {
    if (!this.options.editable) return;
    const table = parseTsv(text);
    if (!table.length) return;
    const range = this.normalized();
    const targets: CellInput[] = [];
    if (table.length === 1 && table[0].length === 1 && (range.r1 !== range.r2 || range.c1 !== range.c2)) {
      for (let r = range.r1; r <= Math.min(range.r2, this.rows - 1); r++) {
        for (let c = range.c1; c <= Math.min(range.c2, this.cols - 1); c++) targets.push({ r, c, input: table[0][0] });
      }
    } else {
      table.forEach((row, dr) => row.forEach((input, dc) => targets.push({ r: range.r1 + dr, c: range.c1 + dc, input })));
      const last = targets[targets.length - 1];
      this.select(range.r1, range.c1, { reveal: false });
      this.select(last.r, last.c, { reveal: false, extend: true });
    }
    const inside = targets.filter(({ r, c }) => r < MAX_ROWS && c < MAX_COLS);
    if (inside.length) this.options.onEdit?.(inside);
  }

  // --- Input -----------------------------------------------------------------------------

  private listen(): void {
    const point = (e: MouseEvent) => {
      const rect = this.canvas.getBoundingClientRect();
      return { r: search(this.rowOffsets, e.clientY - rect.top), c: search(this.colOffsets, e.clientX - rect.left) };
    };
    const headerPoint = (e: MouseEvent, axis: 'row' | 'col') => {
      const rect = this.body.getBoundingClientRect();
      return axis === 'col'
        ? search(this.colOffsets, e.clientX - rect.left + this.body.scrollLeft)
        : search(this.rowOffsets, e.clientY - rect.top + this.body.scrollTop);
    };
    const contextMenu = (e: MouseEvent, kind: 'cell' | 'row' | 'col', r: number, c: number) => {
      if (!this.options.onContextMenu) return;
      e.preventDefault();
      const range = this.normalized();
      const inside = r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2;
      const fullRows = range.c1 === 0 && range.c2 === MAX_COLS - 1;
      const fullCols = range.r1 === 0 && range.r2 === MAX_ROWS - 1;
      if (!inside || (kind === 'row' && !fullRows) || (kind === 'col' && !fullCols)) {
        if (kind === 'row') this.selectLines('row', r, false);
        else if (kind === 'col') this.selectLines('col', c, false);
        else this.select(r, c, { reveal: false });
      }
      this.focus();
      this.options.onContextMenu(e, kind, this.normalized());
    };

    this.canvas.addEventListener('mousedown', (e) => {
      if (e.button !== 0 || e.target === this.editor) return;
      e.preventDefault();
      if (this.editing) this.endEdit(true);
      const { r, c } = point(e);
      this.select(r, c, { reveal: false, extend: e.shiftKey });
      this.dragging = 'cell';
      this.focus();
    });
    this.canvas.addEventListener('dblclick', (e) => {
      if (e.target === this.editor) return;
      this.beginEdit('edit');
    });
    this.canvas.addEventListener('contextmenu', (e) => {
      if (this.editing) return;
      const { r, c } = point(e);
      contextMenu(e, 'cell', r, c);
    });
    for (const [header, axis] of [[this.colHeader, 'col'], [this.rowHeader, 'row']] as const) {
      header.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        e.preventDefault();
        if (this.editing) this.endEdit(true);
        this.selectLines(axis, headerPoint(e, axis), e.shiftKey);
        this.dragging = axis;
        this.focus();
      });
      header.addEventListener('contextmenu', (e) => {
        const i = headerPoint(e, axis);
        contextMenu(e, axis, axis === 'row' ? i : 0, axis === 'col' ? i : 0);
      });
    }
    this.corner.addEventListener('mousedown', (e) => {
      e.preventDefault();
      this.selectAll();
      this.focus();
    });
    window.addEventListener('mousemove', (e) => {
      if (!this.dragging || !(e.buttons & 1)) {
        this.dragging = undefined;
        return;
      }
      if (this.dragging === 'cell') {
        const { r, c } = point(e);
        this.select(r, c, { reveal: false, extend: true });
      } else {
        this.selectLines(this.dragging, headerPoint(e, this.dragging), true);
      }
    });
    window.addEventListener('mouseup', () => (this.dragging = undefined));

    this.editor.addEventListener('keydown', (e) => this.onKey(e));
    this.editor.addEventListener('compositionstart', () => {
      if (!this.editing) this.beginEdit('enter', '');
    });
    this.editor.addEventListener('input', () => {
      if (!this.editing) {
        // A typed character while navigating replaces the cell's content.
        const typed = this.editor.value;
        if (!this.options.editable) {
          this.editor.value = '';
          return;
        }
        this.beginEdit('enter', typed);
        return;
      }
      this.placeEditor();
      this.options.onEditorInput?.(this.editor.value);
    });
    this.editor.addEventListener('copy', (e) => {
      if (this.editing) return;
      e.preventDefault();
      e.clipboardData?.setData('text/plain', this.selectionText());
    });
    this.editor.addEventListener('cut', (e) => {
      if (this.editing) return;
      e.preventDefault();
      e.clipboardData?.setData('text/plain', this.selectionText());
      if (this.options.editable) this.options.onClear?.(this.normalized());
    });
    this.editor.addEventListener('paste', (e) => {
      if (this.editing) return;
      e.preventDefault();
      this.paste(e.clipboardData?.getData('text/plain') ?? '');
    });
    this.editor.addEventListener('blur', () => {
      // Clicking elsewhere in the webview (formula bar, tabs) commits like Excel does.
      if (this.editing) this.endEdit(true);
    });
  }

  private onKey(e: KeyboardEvent): void {
    if (e.isComposing || e.keyCode === 229) return;
    if (this.editing) {
      const { mode } = this.editing;
      const move = (dr: number, dc: number, toRange = false) => {
        e.preventDefault();
        this.endEdit(true, toRange);
        if (!toRange) this.moveBy(dr, dc);
      };
      if (e.key === 'Enter' && e.altKey) {
        e.preventDefault();
        this.editor.setRangeText('\n', this.editor.selectionStart, this.editor.selectionEnd, 'end');
        this.placeEditor();
        this.options.onEditorInput?.(this.editor.value);
      } else if (e.key === 'Enter') move(e.shiftKey ? -1 : 1, 0, e.ctrlKey);
      else if (e.key === 'Tab') move(0, e.shiftKey ? -1 : 1);
      else if (e.key === 'Escape') {
        e.preventDefault();
        this.endEdit(false);
      } else if (mode === 'enter' && ARROWS[e.key]) move(...ARROWS[e.key]);
      else if (e.key === 'F2') {
        e.preventDefault();
        this.editing.mode = mode === 'enter' ? 'edit' : 'enter';
      }
      return;
    }

    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'F2' || (e.key === 'Backspace' && !mod)) {
      e.preventDefault();
      this.beginEdit(e.key === 'F2' ? 'edit' : 'enter', e.key === 'F2' ? undefined : '');
      return;
    }
    if (e.key === 'Delete') {
      e.preventDefault();
      if (this.options.editable) this.options.onClear?.(this.normalized());
      return;
    }
    if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      this.selectAll();
      return;
    }
    // Shift extends the selection from its far corner; plain keys move the active cell.
    const { r, c } = e.shiftKey ? this.extent : this.active;
    const merge = e.shiftKey ? undefined : this.anchors.get(key(r, c));
    const page = Math.max(1, Math.floor(this.body.clientHeight / this.sheet.defaultRowHeight) - 1);
    const lastRow = Math.max(this.sheet.rowCount - 1, 0), lastCol = Math.max(this.sheet.colCount - 1, 0);
    const moves: Record<string, [number, number] | undefined> = {
      ArrowUp: [mod ? this.jump(r, c, -1, 0) : r - 1, c],
      ArrowDown: [mod ? this.jump(r, c, 1, 0) : (merge?.r2 ?? r) + 1, c],
      ArrowLeft: [r, mod ? this.jump(r, c, 0, -1) : c - 1],
      ArrowRight: [r, mod ? this.jump(r, c, 0, 1) : (merge?.c2 ?? c) + 1],
      PageUp: [r - page, c],
      PageDown: [r + page, c],
      Home: [mod ? 0 : r, 0],
      End: mod ? [lastRow, lastCol] : undefined,
    };
    if (e.key === 'Enter' || e.key === 'Tab') {
      e.preventDefault();
      const down = e.key === 'Enter';
      this.moveBy(down ? (e.shiftKey ? -1 : 1) : 0, down ? 0 : e.shiftKey ? -1 : 1);
      return;
    }
    const target = moves[e.key];
    if (!target) return;
    e.preventDefault();
    let [nr, nc] = target;
    // Skip hidden rows/columns in the direction of travel.
    const dr = Math.sign(nr - r), dc = Math.sign(nc - c);
    while (dr && nr > 0 && nr < this.rows - 1 && this.rowOffsets[nr + 1] === this.rowOffsets[nr]) nr += dr;
    while (dc && nc > 0 && nc < this.cols - 1 && this.colOffsets[nc + 1] === this.colOffsets[nc]) nc += dc;
    this.select(nr, nc, { reveal: true, extend: e.shiftKey });
  }

  /** Ctrl+arrow: to the edge of the current block of values, or the next value. */
  private jump(r: number, c: number, dr: number, dc: number): number {
    const filled = (rr: number, cc: number) => !!this.cells.get(key(rr, cc))?.v;
    const limit = dr ? this.rows - 1 : this.cols - 1;
    let pos = dr ? r : c;
    const at = (p: number) => (dr ? filled(p, c) : filled(r, p));
    const step = dr || dc;
    const next = pos + step;
    if (next < 0 || next > limit) return pos;
    if (at(pos) && at(next)) {
      while (pos + step >= 0 && pos + step <= limit && at(pos + step)) pos += step;
      return pos;
    }
    pos = next;
    while (pos + step >= 0 && pos + step <= limit && !at(pos)) pos += step;
    return pos;
  }

  /** Moves the active cell, staying inside a multi-cell selection like Excel's Enter/Tab. */
  private moveBy(dr: number, dc: number): void {
    const range = this.normalized();
    const multi = range.r1 !== range.r2 || range.c1 !== range.c2;
    if (!multi || range.r2 >= this.rows || range.c2 >= this.cols) {
      const merge = this.anchors.get(key(this.active.r, this.active.c));
      const r = dr > 0 ? (merge?.r2 ?? this.active.r) + dr : this.active.r + dr;
      const c = dc > 0 ? (merge?.c2 ?? this.active.c) + dc : this.active.c + dc;
      this.select(r, c, { reveal: true });
      return;
    }
    let { r, c } = this.active;
    r += dr;
    c += dc;
    if (c > range.c2) [c, r] = [range.c1, r + 1];
    if (c < range.c1) [c, r] = [range.c2, r - 1];
    if (r > range.r2) [r, c] = [range.r1, dc ? range.c1 : c + 1 > range.c2 ? range.c1 : c + 1];
    if (r < range.r1) [r, c] = [range.r2, dc ? range.c2 : c - 1 < range.c1 ? range.c2 : c - 1];
    this.active = { r, c };
    this.paintSelection();
    this.reveal(r, c);
    this.notifySelect();
  }

  // --- Selection -------------------------------------------------------------------------

  private select(r: number, c: number, { reveal, extend = false }: { reveal: boolean; extend?: boolean }): void {
    r = Math.min(Math.max(r, 0), this.rows - 1);
    c = Math.min(Math.max(c, 0), this.cols - 1);
    if (extend) {
      this.extent = { r, c };
    } else {
      const merge = this.covered.get(key(r, c)) ?? this.anchors.get(key(r, c));
      this.active = { r: merge?.r1 ?? r, c: merge?.c1 ?? c };
      this.extent = { r: merge?.r2 ?? r, c: merge?.c2 ?? c };
    }
    this.paintSelection();
    if (reveal) this.reveal(extend ? r : this.active.r, extend ? c : this.active.c);
    this.notifySelect();
    this.schedule();
  }

  /** Selects whole rows or columns from the anchor line (the active cell's) to `index`. */
  private selectLines(axis: 'row' | 'col', index: number, extend: boolean): void {
    if (axis === 'row') {
      const anchor = extend ? this.active.r : index;
      this.active = { r: anchor, c: 0 };
      this.extent = { r: index, c: MAX_COLS - 1 };
    } else {
      const anchor = extend ? this.active.c : index;
      this.active = { r: 0, c: anchor };
      this.extent = { r: MAX_ROWS - 1, c: index };
    }
    this.paintSelection();
    this.notifySelect();
    this.schedule();
  }

  private selectAll(): void {
    this.active = { r: 0, c: 0 };
    this.extent = { r: MAX_ROWS - 1, c: MAX_COLS - 1 };
    this.paintSelection();
    this.notifySelect();
    this.schedule();
  }

  private normalized(): Range {
    const range = {
      r1: Math.min(this.active.r, this.extent.r),
      c1: Math.min(this.active.c, this.extent.c),
      r2: Math.max(this.active.r, this.extent.r),
      c2: Math.max(this.active.c, this.extent.c),
    };
    // Grow the range until it fully contains every merged area it touches.
    for (let changed = true; changed; ) {
      changed = false;
      for (const m of this.sheet.merges) {
        if (m.r2 < range.r1 || m.r1 > range.r2 || m.c2 < range.c1 || m.c1 > range.c2) continue;
        if (m.r1 < range.r1 || m.c1 < range.c1 || m.r2 > range.r2 || m.c2 > range.c2) {
          range.r1 = Math.min(range.r1, m.r1);
          range.c1 = Math.min(range.c1, m.c1);
          range.r2 = Math.max(range.r2, m.r2);
          range.c2 = Math.max(range.c2, m.c2);
          changed = true;
        }
      }
    }
    return range;
  }

  private paintSelection(): void {
    const range = this.normalized();
    const box = (target: HTMLElement, r1: number, c1: number, r2: number, c2: number) => {
      r2 = Math.min(r2, this.rows - 1);
      c2 = Math.min(c2, this.cols - 1);
      const left = this.colOffsets[c1], top = this.rowOffsets[r1];
      Object.assign(target.style, {
        left: `${left}px`,
        top: `${top}px`,
        width: `${this.colOffsets[c2 + 1] - left}px`,
        height: `${this.rowOffsets[r2 + 1] - top}px`,
      });
    };
    const merge = this.anchors.get(key(this.active.r, this.active.c));
    box(this.selectionBox, this.active.r, this.active.c, merge?.r2 ?? this.active.r, merge?.c2 ?? this.active.c);
    const single = this.isSingle(range);
    this.rangeBox.hidden = single;
    if (!single) box(this.rangeBox, range.r1, range.c1, range.r2, range.c2);
    this.placeEditor();
  }

  private placeEditor(): void {
    const { r, c } = this.active;
    const merge = this.anchors.get(key(r, c));
    const left = this.colOffsets[c], top = this.rowOffsets[r];
    const width = this.colOffsets[(merge?.c2 ?? c) + 1] - left;
    const height = this.rowOffsets[(merge?.r2 ?? r) + 1] - top;
    const style = this.editor.style;
    style.left = `${left}px`;
    style.top = `${top}px`;
    if (!this.editing) {
      style.width = `${Math.max(width, 4)}px`;
      style.height = `${Math.max(height, 4)}px`;
      return;
    }
    // Grow with the content like Excel's in-cell editor.
    style.width = `${width}px`;
    style.height = `${height}px`;
    const maxWidth = Math.max(width, this.body.clientWidth - (left - this.body.scrollLeft) - 8);
    style.width = `${Math.min(Math.max(width, this.editor.scrollWidth + 4), maxWidth)}px`;
    style.height = `${Math.max(height, this.editor.scrollHeight)}px`;
  }

  private reveal(r: number, c: number): void {
    const b = this.body;
    const top = this.rowOffsets[r], height = this.rowOffsets[r + 1] - top;
    const left = this.colOffsets[c], width = this.colOffsets[c + 1] - left;
    if (top < b.scrollTop) b.scrollTop = top;
    else if (top + height > b.scrollTop + b.clientHeight) b.scrollTop = top + height - b.clientHeight;
    if (left < b.scrollLeft) b.scrollLeft = left;
    else if (left + width > b.scrollLeft + b.clientWidth) b.scrollLeft = left + width - b.clientWidth;
  }

  private isSingle(range: Range): boolean {
    const merge = this.anchors.get(key(this.active.r, this.active.c));
    return range.r1 === this.active.r && range.c1 === this.active.c && range.r2 === (merge?.r2 ?? this.active.r) && range.c2 === (merge?.c2 ?? this.active.c);
  }

  private notifySelect(): void {
    const range = this.normalized();
    this.options.onSelect(this.cells.get(key(this.active.r, this.active.c)), { ...this.active }, range, this.isSingle(range));
  }

  // --- Rendering -------------------------------------------------------------------------

  private schedule(): void {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.render();
    });
  }

  private render(): void {
    const { scrollTop, scrollLeft, clientWidth, clientHeight } = this.body;
    const r1 = Math.max(0, search(this.rowOffsets, scrollTop) - OVERSCAN);
    const r2 = Math.min(this.rows - 1, search(this.rowOffsets, scrollTop + clientHeight) + OVERSCAN);
    const c1 = Math.max(0, search(this.colOffsets, scrollLeft) - OVERSCAN);
    const c2 = Math.min(this.cols - 1, search(this.colOffsets, scrollLeft + clientWidth) + OVERSCAN);

    const html: string[] = [];
    const drawnMerges = new Set<MergeRange>();
    // Merged ranges that start outside the viewport still need their anchor drawn.
    for (const m of this.sheet.merges) {
      if (m.r2 < r1 || m.r1 > r2 || m.c2 < c1 || m.c1 > c2) continue;
      if (m.r1 < r1 || m.c1 < c1) {
        html.push(this.cellHtml(m.r1, m.c1, m));
        drawnMerges.add(m);
      }
    }
    for (let r = r1; r <= r2; r++) {
      if (this.rowOffsets[r + 1] === this.rowOffsets[r]) continue;
      for (let c = c1; c <= c2; c++) {
        const k = key(r, c);
        if (this.covered.has(k)) continue;
        const merge = this.anchors.get(k);
        if (merge && drawnMerges.has(merge)) continue;
        html.push(this.cellHtml(r, c, merge));
      }
    }
    this.cellLayer.innerHTML = html.join('');

    const range = this.normalized();
    const colHtml: string[] = [];
    for (let c = c1; c <= c2; c++) {
      const w = this.colOffsets[c + 1] - this.colOffsets[c];
      if (!w) continue;
      const active = c >= range.c1 && c <= range.c2 ? ' active' : '';
      colHtml.push(`<div class="hcell${active}" style="left:${this.colOffsets[c]}px;width:${w}px">${columnName(c)}</div>`);
    }
    this.colHeaderLayer.innerHTML = colHtml.join('');
    this.colHeaderLayer.style.transform = `translateX(${-scrollLeft}px)`;

    const rowHtml: string[] = [];
    for (let r = r1; r <= r2; r++) {
      const h = this.rowOffsets[r + 1] - this.rowOffsets[r];
      if (!h) continue;
      const active = r >= range.r1 && r <= range.r2 ? ' active' : '';
      rowHtml.push(`<div class="hcell${active}" style="top:${this.rowOffsets[r]}px;height:${h}px">${r + 1}</div>`);
    }
    this.rowHeaderLayer.innerHTML = rowHtml.join('');
    this.rowHeaderLayer.style.transform = `translateY(${-scrollTop}px)`;
  }

  private cellHtml(r: number, c: number, merge?: MergeRange): string {
    const left = this.colOffsets[c];
    const top = this.rowOffsets[r];
    const width = this.colOffsets[merge ? merge.c2 + 1 : c + 1] - left;
    const height = this.rowOffsets[merge ? merge.r2 + 1 : r + 1] - top;
    const cell = this.cells.get(key(r, c));
    const pos = `left:${left}px;top:${top}px;width:${width}px;height:${height}px`;
    if (!cell) return `<div class="cell empty" style="${pos}"></div>`;
    const style = this.sheet.styles[cell.s] ?? {};
    const classes = ['cell', `s${cell.s}`, `t-${cell.t}`];
    if (merge) classes.push('merged');
    // Formulas entered here have no cached result until Excel recalculates them.
    const pending = !!cell.f && cell.v === '';
    const text = pending ? `=${cell.f}` : cell.v;
    if (pending) classes.push('pending');
    // Unwrapped, left-aligned text may spill into empty neighbours like in Excel.
    if ((cell.t === 's' || pending) && !style.wrap && !merge && (!style.hAlign || style.hAlign === 'left') && this.isEmpty(r, c + 1)) classes.push('spill');
    if (!text) classes.push('blank');
    const title = pending ? ' title="Calculated when the file is opened in Excel"' : '';
    return `<div class="${classes.join(' ')}" style="${pos}"${title}><span>${escapeHtml(text)}</span></div>`;
  }

  private isEmpty(r: number, c: number): boolean {
    const k = key(r, c);
    return !this.covered.has(k) && !this.anchors.has(k) && !this.cells.get(k)?.v;
  }
}

const ARROWS: Record<string, [number, number]> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

function key(r: number, c: number): number {
  return r * MAX_COLS + c;
}

function offsets(count: number, size: (i: number) => number): Float64Array {
  const out = new Float64Array(count + 1);
  for (let i = 0; i < count; i++) out[i + 1] = out[i] + size(i);
  return out;
}

/** Index i such that offsets[i] <= value < offsets[i + 1]. */
function search(offsets: Float64Array, value: number): number {
  let lo = 0, hi = offsets.length - 2;
  if (value <= 0) return 0;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (offsets[mid] <= value) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** Splits clipboard text into rows of fields, honouring Excel's quoting of tabs, newlines and quotes. */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let i = 0;
  const source = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  if (!source) return [];
  while (i < source.length) {
    const ch = source[i];
    if (ch === '"' && field === '') {
      let j = i + 1;
      let value = '';
      while (j < source.length) {
        if (source[j] === '"' && source[j + 1] === '"') {
          value += '"';
          j += 2;
        } else if (source[j] === '"') break;
        else value += source[j++];
      }
      if (j < source.length && (j + 1 === source.length || source[j + 1] === '\t' || source[j + 1] === '\n')) {
        field = value;
        i = j + 1;
        continue;
      }
    }
    if (ch === '\t') {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += ch;
    }
    i++;
  }
  row.push(field);
  rows.push(row);
  return rows;
}

function styleCss(s: CellStyle): string {
  const css: string[] = [];
  if (s.bold) css.push('font-weight:bold');
  if (s.italic) css.push('font-style:italic');
  const deco = [s.underline && 'underline', s.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) css.push(`text-decoration:${deco}`);
  if (s.color) css.push(`color:${s.color}`);
  if (s.fontSize) css.push(`font-size:${(s.fontSize * 4) / 3}px`);
  if (s.fontFamily) css.push(`font-family:"${s.fontFamily.replace(/"/g, '')}",var(--sheet-font)`);
  if (s.background) css.push(`background:${s.background}`);
  const h = s.hAlign === 'center' || s.hAlign === 'centerContinuous' ? 'center' : s.hAlign === 'right' ? 'flex-end' : s.hAlign === 'left' || s.hAlign === 'fill' ? 'flex-start' : s.hAlign === 'justify' || s.hAlign === 'distributed' ? 'stretch' : undefined;
  if (h) css.push(`justify-content:${h}`, `text-align:${h === 'flex-end' ? 'right' : h === 'center' ? 'center' : 'left'}`);
  const v = s.vAlign === 'top' ? 'flex-start' : s.vAlign === 'center' ? 'center' : s.vAlign === 'justify' || s.vAlign === 'distributed' ? 'stretch' : undefined;
  if (v) css.push(`align-items:${v}`);
  if (s.wrap) css.push('white-space:pre-wrap', 'overflow-wrap:anywhere');
  if (s.indent) css.push(`padding-left:${2 + s.indent * 9}px`);
  if (s.borderTop) css.push(`border-top:${s.borderTop}`);
  if (s.borderRight) css.push(`border-right:${s.borderRight}`);
  if (s.borderBottom) css.push(`border-bottom:${s.borderBottom}`);
  if (s.borderLeft) css.push(`border-left:${s.borderLeft}`);
  return css.join(';');
}
