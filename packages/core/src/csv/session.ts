// A CSV/TSV file presented as a single-sheet workbook, so the Excel grid can show and edit it.
// Values are plain text: what is typed is stored verbatim (no formulas or number formats).
import type { HostMessage, WebviewMessage } from '../protocol';
import { EditError } from '../xlsx/workbook';
import type { SessionEdit, SheetSession } from '../xlsx/session';
import type { CellAddress, CellData, EditOp, SheetData } from '../xlsx/types';
import { decodeText, encodeText, ENCODING_LABELS, type TextEncodingName } from './encoding';
import { delimiterFor, parseDelimited, serializeDelimited, type DelimitedText } from './delimited';

const NUMBER = /^[+-]?(?:\d[\d,]*(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$/;
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;

interface Snapshot {
  rows: string[][];
  raws: (string | undefined)[];
}

export class CsvSession implements SheetSession {
  private doc!: DelimitedText;
  private encoding!: TextEncodingName;
  private generation = 0;
  private queue: Promise<unknown> = Promise.resolve();

  constructor(
    private readonly fileName: string,
    data: Uint8Array,
  ) {
    this.load(data);
  }

  reset(data: Uint8Array): void {
    this.generation++;
    this.load(data);
  }

  handle(message: WebviewMessage): Promise<HostMessage[]> {
    return this.run(async () => {
      switch (message.type) {
        case 'ready':
          return [this.workbookMessage()];
        case 'loadSheet':
          return message.index === 0 ? [{ type: 'sheet', sheet: this.sheet() }] : [{ type: 'error', message: `Sheet ${message.index} does not exist.` }];
        default:
          return [];
      }
    });
  }

  edit(op: EditOp): Promise<SessionEdit> {
    return this.run(async () => {
      const generation = this.generation;
      let revert = this.apply(op);
      const messages = () => (revert.cells ? [this.cellsMessage(revert.cells)] : [this.workbookMessage()]);
      const step = (fn: () => HostMessage[]) => () => this.run(async () => (generation === this.generation ? fn() : []));
      return {
        label: LABELS[op.type],
        messages: messages(),
        undo: step(() => {
          revert.restore();
          return messages();
        }),
        redo: step(() => {
          revert = this.apply(op);
          return messages();
        }),
      };
    });
  }

  save(): Promise<Uint8Array> {
    return this.run(async () => encodeText(serializeDelimited(this.doc), this.encoding));
  }

  private load(data: Uint8Array): void {
    const { text, encoding } = decodeText(data);
    this.encoding = encoding;
    this.doc = parseDelimited(text, delimiterFor(this.fileName, text));
  }

  /** Applies an edit; returns how to undo it and which cells changed (none for structural edits). */
  private apply(op: EditOp): { restore(): void; cells?: CellAddress[] } {
    if ('sheet' in op && op.sheet !== 0) throw new EditError(`Sheet ${op.sheet} does not exist.`);
    const { rows, raws } = this.doc;
    switch (op.type) {
      case 'setCells':
      case 'clearRange': {
        const targets =
          op.type === 'setCells'
            ? op.cells
            : rows.slice(op.range.r1, op.range.r2 + 1).flatMap((row, i) =>
                row.flatMap((value, c) => (value && c >= op.range.c1 && c <= op.range.c2 ? [{ r: op.range.r1 + i, c, input: '' }] : [])),
              );
        if (targets.some(({ r, c }) => r >= MAX_ROWS || c >= MAX_COLS)) throw new EditError('The cell is outside the sheet.');
        const rowCount = rows.length;
        const saved = new Map<number, { fields: string[]; raw: string | undefined }>();
        for (const { r, c, input } of targets) {
          if (r < rowCount && !saved.has(r)) saved.set(r, { fields: rows[r].slice(), raw: raws[r] });
          if (input === '' && (r >= rows.length || c >= rows[r].length)) continue;
          while (rows.length <= r) {
            rows.push([]);
            raws.push(undefined);
          }
          const row = rows[r];
          while (row.length <= c) row.push('');
          row[c] = input;
          raws[r] = undefined;
        }
        return {
          cells: targets.map(({ r, c }) => ({ r, c })),
          restore: () => {
            rows.length = Math.min(rows.length, rowCount);
            raws.length = rows.length;
            for (const [r, { fields, raw }] of saved) {
              rows[r] = fields;
              raws[r] = raw;
            }
          },
        };
      }
      case 'insertRows':
      case 'deleteRows':
      case 'insertCols':
      case 'deleteCols': {
        const snapshot: Snapshot = { rows: rows.map((r) => r.slice()), raws: raws.slice() };
        const count = Math.max(1, Math.floor(op.count));
        if (op.type === 'insertRows') {
          if (op.at < rows.length) {
            rows.splice(op.at, 0, ...Array.from({ length: count }, () => [] as string[]));
            raws.splice(op.at, 0, ...Array.from({ length: count }, () => undefined));
          }
        } else if (op.type === 'deleteRows') {
          rows.splice(op.at, count);
          raws.splice(op.at, count);
        } else {
          rows.forEach((row, r) => {
            if (row.length <= op.at) return;
            if (op.type === 'insertCols') row.splice(op.at, 0, ...Array.from({ length: count }, () => ''));
            else row.splice(op.at, count);
            raws[r] = undefined;
          });
        }
        return {
          restore: () => {
            this.doc.rows = snapshot.rows;
            this.doc.raws = snapshot.raws;
          },
        };
      }
      default:
        throw new EditError('CSV files have a single sheet.');
    }
  }

  private sheet(): SheetData {
    const { rows } = this.doc;
    const cells: CellData[] = [];
    let colCount = 0;
    // Column widths from the longest value in the first rows, like a quick "autofit".
    const widths: number[] = [];
    rows.forEach((row, r) => {
      if (row.length > colCount) colCount = row.length;
      row.forEach((value, c) => {
        if (value) cells.push(this.cell(r, c, value));
        if (r < 1000) widths[c] = Math.max(widths[c] ?? 0, displayWidth(value));
      });
    });
    const colWidths: Record<number, number> = {};
    widths.forEach((w, c) => {
      if (w) colWidths[c] = Math.min(400, Math.max(64, w * 8 + 12));
    });
    return {
      index: 0,
      name: this.sheetName(),
      rowCount: rows.length,
      colCount,
      defaultColWidth: 64,
      defaultRowHeight: 22,
      colWidths,
      rowHeights: {},
      cells,
      styles: { 0: {} },
      merges: [],
      showGridLines: true,
      unsupported: [],
    };
  }

  private cell(r: number, c: number, value: string): CellData {
    return { r, c, v: value, s: 0, t: NUMBER.test(value.trim()) ? 'n' : 's' };
  }

  private cellsMessage(addresses: CellAddress[]): HostMessage {
    const cells: CellData[] = [];
    const removed: CellAddress[] = [];
    for (const { r, c } of addresses) {
      const value = this.doc.rows[r]?.[c];
      if (value) cells.push(this.cell(r, c, value));
      else removed.push({ r, c });
    }
    return { type: 'cells', sheet: 0, cells, removed, styles: {} };
  }

  private workbookMessage(): HostMessage {
    const { delimiter, eol } = this.doc;
    const separator = delimiter === '\t' ? 'Tab' : delimiter === ',' ? 'Comma' : `"${delimiter}"`;
    const lineBreak = eol === '\r\n' ? 'CRLF' : eol === '\n' ? 'LF' : 'CR';
    return {
      type: 'workbook',
      fileName: this.fileName,
      info: { sheets: [{ index: 0, name: this.sheetName(), state: 'visible' }] },
      editable: true,
      format: 'csv',
      details: `${separator} separated · ${ENCODING_LABELS[this.encoding]} · ${lineBreak}`,
    };
  }

  private sheetName(): string {
    const dot = this.fileName.lastIndexOf('.');
    return dot > 0 ? this.fileName.slice(0, dot) : this.fileName;
  }

  private run<T>(fn: () => Promise<T>): Promise<T> {
    const result = this.queue.then(fn);
    this.queue = result.catch(() => undefined);
    return result;
  }
}

const LABELS: Record<EditOp['type'], string> = {
  setCells: 'Edit cells',
  clearRange: 'Clear cells',
  insertRows: 'Insert rows',
  deleteRows: 'Delete rows',
  insertCols: 'Insert columns',
  deleteCols: 'Delete columns',
  addSheet: 'Add sheet',
  renameSheet: 'Rename sheet',
  deleteSheet: 'Delete sheet',
};

/** Approximate width in half-width characters (East Asian wide characters count double). */
function displayWidth(value: string): number {
  let w = 0;
  for (const ch of value) w += /[ᄀ-ᅟ⺀-꓏가-힣豈-﫿︰-﹏＀-｠￠-￦]/.test(ch) ? 2 : 1;
  return w;
}
