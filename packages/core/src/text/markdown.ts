// Small Markdown helpers shared by the document → text converters.

/** Makes text safe inside a table cell: pipes escaped, line breaks as <br>. */
export function escapeCell(text: string): string {
  return text.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>').trim();
}

/**
 * A GitHub-flavoured table; the first row is the header. Body rows are padded to the header's width
 * unless `trim` is set, in which case trailing empty cells are dropped (still valid GFM) so that a
 * new column only changes the rows that use it.
 */
export function markdownTable(rows: string[][], trim = false): string {
  if (!rows.length) return '';
  const width = Math.max(1, ...rows.map((r) => r.length));
  const line = (cells: string[], pad: boolean) => {
    let n = pad ? width : cells.length;
    while (!pad && n > 1 && !cells[n - 1]) n--;
    return `| ${Array.from({ length: n }, (_, i) => cells[i] ?? '').join(' | ')} |`;
  };
  return [line(rows[0], true), `|${' --- |'.repeat(width)}`, ...rows.slice(1).map((r) => line(r, !trim))].join('\n');
}
