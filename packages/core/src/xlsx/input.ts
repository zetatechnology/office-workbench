// What a user typed into a cell → the value to store, and the reverse (text for the cell editor).
import { isDateFormat, serialToDate } from './numfmt';

export type ParsedInput =
  | { kind: 'empty' }
  | { kind: 'formula'; formula: string }
  | { kind: 'string'; text: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'error'; code: string }
  /** `numFmtId` is a built-in format to apply when the cell is still General (e.g. 14 for dates). */
  | { kind: 'number'; value: number; numFmtId?: number };

const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
const NUMBER = /^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?(?:[eE][+-]?\d+)?$/;
const DATE = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const TIME = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/;

/** `format` is the cell's current number format; text-formatted ("@") cells keep input as text. */
export function parseInput(input: string, format: string, date1904 = false): ParsedInput {
  if (input === '') return { kind: 'empty' };
  if (format === '@') return { kind: 'string', text: input };
  if (input.startsWith('=') && input.length > 1) return { kind: 'formula', formula: input.slice(1) };
  if (input.startsWith("'")) return { kind: 'string', text: input.slice(1) };
  const trimmed = input.trim();
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') return { kind: 'bool', value: upper === 'TRUE' };
  if (ERRORS.includes(upper)) return { kind: 'error', code: upper };

  const general = format === 'General';
  const percent = trimmed.endsWith('%');
  const numeric = percent ? trimmed.slice(0, -1).trim() : trimmed;
  if (/\d/.test(numeric) && NUMBER.test(numeric)) {
    const value = Number(numeric.replace(/,/g, ''));
    if (Number.isFinite(value)) {
      const decimals = numeric.includes('.');
      if (percent) return { kind: 'number', value: value / 100, numFmtId: general ? (decimals ? 10 : 9) : undefined };
      if (numeric.includes(',') && general) return { kind: 'number', value, numFmtId: decimals ? 4 : 3 };
      return { kind: 'number', value };
    }
  }

  const date = DATE.exec(trimmed);
  if (date) {
    const [, y, mo, d, h, mi, s] = date.map(Number);
    const serial = dateToSerial(y, mo, d, date1904);
    if (serial !== undefined && (date[4] === undefined || (h < 24 && mi < 60 && (s || 0) < 60))) {
      const time = date[4] === undefined ? 0 : (h * 3600 + mi * 60 + (s || 0)) / 86_400;
      const keep = isDateFormat(format);
      return { kind: 'number', value: serial + time, numFmtId: keep ? undefined : date[4] === undefined ? 14 : 22 };
    }
  }
  const time = TIME.exec(trimmed);
  if (time) {
    const [, h, mi, s] = time.map(Number);
    if (h < 24 && mi < 60 && (s || 0) < 60) {
      const keep = isDateFormat(format);
      return { kind: 'number', value: (h * 3600 + mi * 60 + (s || 0)) / 86_400, numFmtId: keep ? undefined : time[3] === undefined ? 20 : 21 };
    }
  }
  return { kind: 'string', text: input };
}

/** Calendar date → Excel serial, or undefined for invalid dates. */
export function dateToSerial(y: number, mo: number, d: number, date1904 = false): number | undefined {
  const ms = Date.UTC(y, mo - 1, d);
  const dt = new Date(ms);
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return undefined;
  if (date1904) {
    const serial = (ms - Date.UTC(1904, 0, 1)) / 86_400_000;
    return serial >= 0 ? serial : undefined;
  }
  let serial = (ms - Date.UTC(1899, 11, 30)) / 86_400_000;
  // Before 1900-03-01 the 1900 date system is off by one (it counts a fictitious 1900-02-29).
  if (serial < 61) serial -= 1;
  return serial >= 1 ? serial : undefined;
}

/** The text shown when editing a stored number, chosen so that re-entering it keeps the value. */
export function numberEditText(value: number, format: string, date1904 = false): string {
  if (isDateFormat(format) && value >= 0) {
    const p = serialToDate(value, date1904);
    const time = value % 1 !== 0;
    const clock = `${p.h}:${String(p.mi).padStart(2, '0')}:${String(p.s).padStart(2, '0')}`;
    if (value < 1) return clock;
    return `${p.y}/${p.mo}/${p.d}${time ? ` ${clock}` : ''}`;
  }
  if (/%/.test(format.replace(/"[^"]*"|\\./g, ''))) return `${plainNumber(value * 100)}%`;
  return plainNumber(value);
}

/** Up to 15 significant digits, like Excel's formula bar. */
function plainNumber(value: number): string {
  return String(Number(value.toPrecision(15)));
}

/** Edit text for a string cell; prefixed with ' when it would otherwise be read as another type. */
export function stringEditText(text: string, format: string): string {
  if (format === '@' || text === '') return text;
  return parseInput(text, format).kind === 'string' && !text.startsWith("'") ? text : `'${text}`;
}
