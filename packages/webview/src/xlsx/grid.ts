// Virtualized sheet grid: only the cells inside the viewport (plus a small overscan) exist in the DOM.
import type { CellData, CellStyle, MergeRange, SheetData } from '@office-workbench/core';
import { columnName } from '@office-workbench/core/address';
import { el, escapeHtml } from '../common/host';

const HEADER_HEIGHT = 22;
const OVERSCAN = 4;
/** Empty rows/columns shown beyond the used range so the sheet does not end abruptly. */
const EXTRA_ROWS = 50;
const EXTRA_COLS = 10;
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;

export interface Selection {
  r: number;
  c: number;
}

export interface GridOptions {
  onSelect(cell: CellData | undefined, selection: Selection): void;
}

export class SheetGrid {
  readonly element: HTMLElement;
  private readonly corner = el('div', 'grid-corner');
  private readonly colHeader = el('div', 'grid-col-header');
  private readonly rowHeader = el('div', 'grid-row-header');
  private readonly body = el('div', 'grid-body');
  private readonly canvas = el('div', 'grid-canvas');
  private readonly cellLayer = el('div', 'grid-cells');
  private readonly selectionBox = el('div', 'grid-selection');
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
  private selection: Selection = { r: 0, c: 0 };

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
    this.element.tabIndex = 0;
    this.styleSheet.textContent = sheet.styles.map((s, i) => `.s${i}{${styleCss(s)}}`).join('\n');
    this.canvas.style.width = `${this.colOffsets[this.cols]}px`;
    this.canvas.style.height = `${this.rowOffsets[this.rows]}px`;
    this.canvas.append(this.cellLayer, this.selectionBox);
    this.body.append(this.canvas);
    this.colHeader.append(this.colHeaderLayer);
    this.rowHeader.append(this.rowHeaderLayer);
    this.element.append(this.styleSheet, this.corner, this.colHeader, this.rowHeader, this.body);

    this.body.addEventListener('scroll', () => this.schedule());
    new ResizeObserver(() => this.schedule()).observe(this.body);
    this.canvas.addEventListener('mousedown', (e) => {
      const rect = this.canvas.getBoundingClientRect();
      this.select(search(this.rowOffsets, e.clientY - rect.top), search(this.colOffsets, e.clientX - rect.left), false);
      this.element.focus();
    });
    this.element.addEventListener('keydown', (e) => this.onKey(e));
  }

  get scroll(): { top: number; left: number } {
    return { top: this.body.scrollTop, left: this.body.scrollLeft };
  }

  restore(scroll: { top: number; left: number } | undefined, selection: Selection | undefined): void {
    if (scroll) {
      this.body.scrollTop = scroll.top;
      this.body.scrollLeft = scroll.left;
    }
    this.select(selection?.r ?? 0, selection?.c ?? 0, false);
    this.schedule();
  }

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

    const colHtml: string[] = [];
    for (let c = c1; c <= c2; c++) {
      const w = this.colOffsets[c + 1] - this.colOffsets[c];
      if (!w) continue;
      const active = c === this.selection.c ? ' active' : '';
      colHtml.push(`<div class="hcell${active}" style="left:${this.colOffsets[c]}px;width:${w}px">${columnName(c)}</div>`);
    }
    this.colHeaderLayer.innerHTML = colHtml.join('');
    this.colHeaderLayer.style.transform = `translateX(${-scrollLeft}px)`;

    const rowHtml: string[] = [];
    for (let r = r1; r <= r2; r++) {
      const h = this.rowOffsets[r + 1] - this.rowOffsets[r];
      if (!h) continue;
      const active = r === this.selection.r ? ' active' : '';
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
    // Unwrapped, left-aligned text may spill into empty neighbours like in Excel.
    if (cell.t === 's' && !style.wrap && !merge && (!style.hAlign || style.hAlign === 'left') && this.isEmpty(r, c + 1)) classes.push('spill');
    if (!cell.v) classes.push('blank');
    return `<div class="${classes.join(' ')}" style="${pos}"><span>${escapeHtml(cell.v)}</span></div>`;
  }

  private isEmpty(r: number, c: number): boolean {
    const k = key(r, c);
    return !this.covered.has(k) && !this.anchors.has(k) && !this.cells.get(k)?.v;
  }

  private select(r: number, c: number, reveal: boolean): void {
    r = Math.min(Math.max(r, 0), this.rows - 1);
    c = Math.min(Math.max(c, 0), this.cols - 1);
    const merge = this.covered.get(key(r, c)) ?? this.anchors.get(key(r, c));
    const r0 = merge?.r1 ?? r, c0 = merge?.c1 ?? c;
    this.selection = { r: r0, c: c0 };
    const left = this.colOffsets[c0], top = this.rowOffsets[r0];
    const width = this.colOffsets[(merge?.c2 ?? c0) + 1] - left;
    const height = this.rowOffsets[(merge?.r2 ?? r0) + 1] - top;
    Object.assign(this.selectionBox.style, { left: `${left}px`, top: `${top}px`, width: `${width}px`, height: `${height}px` });
    if (reveal) {
      const b = this.body;
      if (top < b.scrollTop) b.scrollTop = top;
      else if (top + height > b.scrollTop + b.clientHeight) b.scrollTop = top + height - b.clientHeight;
      if (left < b.scrollLeft) b.scrollLeft = left;
      else if (left + width > b.scrollLeft + b.clientWidth) b.scrollLeft = left + width - b.clientWidth;
    }
    this.options.onSelect(this.cells.get(key(r0, c0)), this.selection);
    this.schedule();
  }

  private onKey(e: KeyboardEvent): void {
    const { r, c } = this.selection;
    const merge = this.anchors.get(key(r, c));
    const page = Math.max(1, Math.floor(this.body.clientHeight / this.sheet.defaultRowHeight) - 1);
    const moves: Record<string, [number, number]> = {
      ArrowUp: [r - 1, c],
      ArrowDown: [(merge?.r2 ?? r) + 1, c],
      ArrowLeft: [r, c - 1],
      ArrowRight: [r, (merge?.c2 ?? c) + 1],
      PageUp: [r - page, c],
      PageDown: [r + page, c],
      Home: [e.ctrlKey ? 0 : r, 0],
      End: [e.ctrlKey ? this.sheet.rowCount - 1 : r, Math.max(this.sheet.colCount - 1, 0)],
      Enter: [(merge?.r2 ?? r) + (e.shiftKey ? -1 : 1), c],
      Tab: [r, c + (e.shiftKey ? -1 : 1)],
    };
    const target = moves[e.key];
    if (!target) return;
    e.preventDefault();
    // Skip hidden rows/columns in the direction of travel.
    let [nr, nc] = target;
    const dr = Math.sign(nr - r), dc = Math.sign(nc - c);
    while (dr && nr > 0 && nr < this.rows - 1 && this.rowOffsets[nr + 1] === this.rowOffsets[nr]) nr += dr;
    while (dc && nc > 0 && nc < this.cols - 1 && this.colOffsets[nc + 1] === this.colOffsets[nc]) nc += dc;
    this.select(nr, nc, true);
  }
}

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
