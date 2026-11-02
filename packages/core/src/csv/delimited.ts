// RFC 4180-style delimited text (CSV / TSV). Each record keeps its original text so that saving
// rewrites only the rows that were edited.

export interface DelimitedText {
  rows: string[][];
  /** Original text of each record (without the line break); undefined once a row is edited. */
  raws: (string | undefined)[];
  delimiter: string;
  /** Line break used when writing: the first one found, or CRLF for single-line files. */
  eol: '\r\n' | '\n' | '\r';
  /** Whether the text ended with a line break. */
  trailingNewline: boolean;
}

/** Delimiter for a file name; ".csv" files using semicolons (common in Europe) are detected. */
export function delimiterFor(fileName: string, text: string): string {
  const ext = fileName.slice(fileName.lastIndexOf('.') + 1).toLowerCase();
  if (ext === 'tsv' || ext === 'tab') return '\t';
  const firstLine = text.slice(0, Math.max(0, text.search(/\r|\n/)) || text.length);
  if (!firstLine.includes(',') && firstLine.includes(';')) return ';';
  if (!firstLine.includes(',') && firstLine.includes('\t')) return '\t';
  return ',';
}

export function parseDelimited(text: string, delimiter: string): DelimitedText {
  const rows: string[][] = [];
  const raws: string[] = [];
  const eolMatch = /\r\n|\n|\r/.exec(text);
  const eol = (eolMatch?.[0] ?? '\r\n') as DelimitedText['eol'];
  let row: string[] = [];
  let field = '';
  let start = 0;
  let i = 0;
  const n = text.length;
  const endRecord = (end: number) => {
    row.push(field);
    rows.push(row);
    raws.push(text.slice(start, end));
    row = [];
    field = '';
  };
  while (i < n) {
    const ch = text[i];
    if (ch === '"' && field === '') {
      // Quoted field: runs to the closing quote; "" is an escaped quote.
      let j = i + 1;
      let value = '';
      for (;;) {
        const q = text.indexOf('"', j);
        if (q < 0) {
          value += text.slice(j);
          j = n;
          break;
        }
        value += text.slice(j, q);
        if (text[q + 1] === '"') {
          value += '"';
          j = q + 2;
        } else {
          j = q + 1;
          break;
        }
      }
      // Anything between the closing quote and the next delimiter is kept, as spreadsheets do.
      let k = j;
      while (k < n && text[k] !== delimiter && text[k] !== '\n' && text[k] !== '\r') k++;
      field = value + text.slice(j, k);
      i = k;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = '';
      i++;
    } else if (ch === '\n' || ch === '\r') {
      endRecord(i);
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      start = i;
    } else {
      // Plain run up to the next special character.
      let k = i + 1;
      while (k < n && text[k] !== delimiter && text[k] !== '\n' && text[k] !== '\r') k++;
      field += text.slice(i, k);
      i = k;
    }
  }
  const trailingNewline = n > 0 && start === n;
  if (!trailingNewline && n > 0) endRecord(n);
  return { rows, raws, delimiter, eol, trailingNewline };
}

export function formatRow(fields: string[], delimiter: string): string {
  return fields
    .map((f) => (f.includes(delimiter) || /["\r\n]/.test(f) || /^\s|\s$/.test(f) ? `"${f.replace(/"/g, '""')}"` : f))
    .join(delimiter);
}

export function serializeDelimited(doc: DelimitedText): string {
  const lines = doc.rows.map((row, i) => doc.raws[i] ?? formatRow(row, doc.delimiter));
  return lines.join(doc.eol) + (doc.trailingNewline && lines.length ? doc.eol : '');
}
