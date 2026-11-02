// Workbook → Markdown: one table per sheet with displayed values, plus each sheet's formulas,
// so diffs show both what changed on screen and what changed in the logic.
import { columnName } from '../xlsx/address';
import { translateFormula } from '../xlsx/refs';
import { Workbook } from '../xlsx/workbook';
import { escapeCell, markdownTable } from './markdown';

export async function xlsxToMarkdown(data: Uint8Array): Promise<string> {
  const workbook = await Workbook.open(data);
  const out: string[] = [];
  for (const info of workbook.info().sheets) {
    const sheet = await workbook.sheet(info.index);
    out.push(`## ${info.name}${info.state === 'visible' ? '' : ` (${info.state === 'hidden' ? 'hidden' : 'very hidden'})`}`);
    const byRow = new Map<number, Map<number, string>>();
    let lastCol = -1;
    const formulas: { r: number; c: number; f: string }[] = [];
    for (const cell of sheet.cells) {
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
      const header = ['', ...Array.from({ length: lastCol + 1 }, (_, c) => columnName(c))];
      const rows = [...byRow.keys()]
        .sort((a, b) => a - b)
        .map((r) => [String(r + 1), ...Array.from({ length: lastCol + 1 }, (_, c) => escapeCell(byRow.get(r)!.get(c) ?? ''))]);
      out.push(markdownTable([header, ...rows], true));
    }
    if (formulas.length) out.push(`### Formulas (${info.name})\n\n${formulaLines(formulas).join('\n')}`);
    if (sheet.merges.length) {
      out.push(`Merged: ${sheet.merges.map((m) => `${columnName(m.c1)}${m.r1 + 1}:${columnName(m.c2)}${m.r2 + 1}`).join(', ')}`);
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
