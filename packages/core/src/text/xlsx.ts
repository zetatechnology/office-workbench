// Workbook → Markdown: one table per sheet with displayed values, plus each sheet's formulas,
// so diffs show both what changed on screen and what changed in the logic.
import { columnName, parseRange } from '../xlsx/address';
import { translateFormula } from '../xlsx/refs';
import { Workbook } from '../xlsx/workbook';
import { escapeCell, markdownTable } from './markdown';

export interface SheetSelection {
  /** Only this sheet (by name, case-insensitive). */
  sheet?: string;
  /** Only this A1 range of the selected sheet (or of every sheet), e.g. "A1:F20". */
  range?: string;
}

export async function xlsxToMarkdown(data: Uint8Array, selection: SheetSelection = {}): Promise<string> {
  const workbook = await Workbook.open(data);
  const all = workbook.info().sheets;
  const sheets = selection.sheet ? all.filter((s) => s.name.toLowerCase() === selection.sheet!.toLowerCase()) : all;
  if (!sheets.length) throw new Error(`No sheet named "${selection.sheet}". Sheets: ${all.map((s) => s.name).join(', ')}`);
  const range = selection.range ? parseRange(selection.range.replace(/\$/g, '')) : undefined;
  if (selection.range && !range) throw new Error(`"${selection.range}" is not an A1 range.`);
  const inRange = (r: number, c: number) => !range || (r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2);
  const firstCol = range?.c1 ?? 0;
  const out: string[] = [];
  for (const info of sheets) {
    const sheet = await workbook.sheet(info.index);
    out.push(`## ${info.name}${info.state === 'visible' ? '' : ` (${info.state === 'hidden' ? 'hidden' : 'very hidden'})`}`);
    const byRow = new Map<number, Map<number, string>>();
    let lastCol = -1;
    const formulas: { r: number; c: number; f: string }[] = [];
    for (const cell of sheet.cells) {
      if (!inRange(cell.r, cell.c)) continue;
      if (cell.v !== '') {
        let row = byRow.get(cell.r);
        if (!row) byRow.set(cell.r, (row = new Map()));
        row.set(cell.c, cell.v);
        lastCol = Math.max(lastCol, cell.c);
      }
      if (cell.f) formulas.push({ r: cell.r, c: cell.c, f: cell.f });
    }
    if (!byRow.size) {
      out.push('_(empty)_');
    } else {
      // Only rows with values are listed; the first column keeps their row numbers.
      const cols = Array.from({ length: lastCol - firstCol + 1 }, (_, i) => firstCol + i);
      const header = ['', ...cols.map(columnName)];
      const rows = [...byRow.keys()]
        .sort((a, b) => a - b)
        .map((r) => [String(r + 1), ...cols.map((c) => escapeCell(byRow.get(r)!.get(c) ?? ''))]);
      out.push(markdownTable([header, ...rows], true));
    }
    if (formulas.length) out.push(`### Formulas (${info.name})\n\n${formulaLines(formulas).join('\n')}`);
    const merges = sheet.merges.filter((m) => inRange(m.r1, m.c1));
    if (merges.length) {
      out.push(`Merged: ${merges.map((m) => `${columnName(m.c1)}${m.r1 + 1}:${columnName(m.c2)}${m.r2 + 1}`).join(', ')}`);
    }
  }
  return out.join('\n\n') + '\n';
}

/**
 * One line per formula, except that a formula filled down a column (each row's formula equals the
 * first one moved down) is listed once with its range.
 */
function formulaLines(formulas: { r: number; c: number; f: string }[]): string[] {
  const sorted = [...formulas].sort((a, b) => a.c - b.c || a.r - b.r);
  const runs: { r1: number; r2: number; c: number; f: string }[] = [];
  for (const { r, c, f } of sorted) {
    const last = runs[runs.length - 1];
    if (last && last.c === c && last.r2 === r - 1 && translateFormula(last.f, r - last.r1, 0) === f) last.r2 = r;
    else runs.push({ r1: r, r2: r, c, f });
  }
  return runs
    .sort((a, b) => a.r1 - b.r1 || a.c - b.c)
    .map(({ r1, r2, c, f }) => `- ${columnName(c)}${r1 + 1}${r2 > r1 ? `:${columnName(c)}${r2 + 1}` : ''}: \`=${f}\``);
}
