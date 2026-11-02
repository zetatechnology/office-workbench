// A1 references inside formula text: locating them (outside string literals and structured
// references) and rewriting them when rows, columns or sheets change.
import { columnIndex, columnName } from './address';

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export interface RefToken {
  /** Sheet name without quotes, when the reference is sheet-qualified. */
  sheet?: string;
  /** True for references into another workbook ([1]Sheet!A1) and 3-D references (S1:S3!A1). */
  foreign: boolean;
  /** cell: A1, range: A1:B2, cols: A:C, rows: 1:3. Rows are -1 for cols, columns -1 for rows. */
  kind: 'cell' | 'range' | 'cols' | 'rows';
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  /** `$` markers: [r1, c1, r2, c2]. */
  abs: [boolean, boolean, boolean, boolean];
}

const NAME = String.raw`[\p{L}\p{N}_.\\]+`;
const PREFIX = String.raw`(?:'((?:[^']|'')+)'|(${NAME}(?::${NAME})?))!`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})`;
const REF = new RegExp(
  String.raw`(?:${PREFIX})?(?:${CELL}(?::${CELL})?|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d{1,7}):(\$?)(\d{1,7}))(?![\p{L}\p{N}_.(!\[])`,
  'uy',
);
const IDENT = /[\p{L}\p{N}_.$\\]/u;

/**
 * Rewrites every reference in a formula. `fn` returns the replacement text for a token
 * (use `formatRef`), `'#REF!'`, or undefined to keep the original text.
 */
export function transformRefs(formula: string, fn: (ref: RefToken) => string | undefined): string {
  let out = '';
  let i = 0;
  let changed = false;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < formula.length && !(formula[j] === '"' && formula[j + 1] !== '"')) j += formula[j] === '"' ? 2 : 1;
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === '[') {
      let depth = 0;
      let j = i;
      for (; j < formula.length; j++) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']' && --depth === 0) break;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (i === 0 || !IDENT.test(formula[i - 1])) {
      REF.lastIndex = i;
      const m = REF.exec(formula);
      const token = m && toToken(m, formula[i - 1] === ']');
      if (m && token) {
        const replacement = fn(token);
        if (replacement !== undefined && replacement !== m[0]) changed = true;
        out += replacement ?? m[0];
        i += m[0].length;
        continue;
      }
    }
    out += ch;
    i++;
  }
  return changed ? out : formula;
}

function toToken(m: RegExpExecArray, afterBracket: boolean): RefToken | undefined {
  const quoted = m[1]?.replace(/''/g, "'");
  const sheet = quoted ?? m[2];
  const foreign = afterBracket || !!sheet?.startsWith('[') || (!quoted && !!m[2]?.includes(':'));
  let token: RefToken;
  if (m[4]) {
    const c1 = columnIndex(m[4]), r1 = Number(m[6]) - 1;
    const hasEnd = !!m[8];
    const c2 = hasEnd ? columnIndex(m[8]) : c1, r2 = hasEnd ? Number(m[10]) - 1 : r1;
    token = { kind: hasEnd ? 'range' : 'cell', r1, c1, r2, c2, abs: [!!m[5], !!m[3], hasEnd ? !!m[9] : !!m[5], hasEnd ? !!m[7] : !!m[3]], sheet, foreign };
  } else if (m[12]) {
    token = { kind: 'cols', r1: -1, r2: -1, c1: columnIndex(m[12]), c2: columnIndex(m[14]), abs: [false, !!m[11], false, !!m[13]], sheet, foreign };
  } else {
    token = { kind: 'rows', r1: Number(m[16]) - 1, r2: Number(m[18]) - 1, c1: -1, c2: -1, abs: [!!m[15], false, !!m[17], false], sheet, foreign };
  }
  const { r1, r2, c1, c2 } = token;
  if (r1 >= MAX_ROWS || r2 >= MAX_ROWS || c1 >= MAX_COLS || c2 >= MAX_COLS) return undefined;
  if ((token.kind !== 'cols' && (r1 < 0 || r2 < 0)) || (token.kind !== 'rows' && (c1 < 0 || c2 < 0))) return undefined;
  return token;
}

export function formatRef(ref: RefToken): string {
  const [ar1, ac1, ar2, ac2] = ref.abs.map((a) => (a ? '$' : ''));
  const prefix = ref.sheet !== undefined ? `${quoteSheetName(ref.sheet)}!` : '';
  const col = (c: number, a: string) => a + columnName(c);
  const row = (r: number, a: string) => a + (r + 1);
  switch (ref.kind) {
    case 'cell':
      return prefix + col(ref.c1, ac1) + row(ref.r1, ar1);
    case 'range':
      return `${prefix}${col(ref.c1, ac1)}${row(ref.r1, ar1)}:${col(ref.c2, ac2)}${row(ref.r2, ar2)}`;
    case 'cols':
      return `${prefix}${col(ref.c1, ac1)}:${col(ref.c2, ac2)}`;
    case 'rows':
      return `${prefix}${row(ref.r1, ar1)}:${row(ref.r2, ar2)}`;
  }
}

/** Sheet names need quotes unless they are plain identifiers that cannot be read as a reference. */
export function quoteSheetName(name: string): string {
  const plain = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^[Rr]\d*[Cc]\d*$/.test(name);
  return plain ? name : `'${name.replace(/'/g, "''")}'`;
}

export function sameSheet(a: string | undefined, b: string): boolean {
  return a !== undefined && a.toLowerCase() === b.toLowerCase();
}

/**
 * Moves a 0-based interval [a, b] for an insertion (count > 0) or deletion (count < 0) at `at`.
 * Returns undefined when the whole interval is deleted or pushed off the sheet.
 */
export function shiftInterval(a: number, b: number, at: number, count: number, max: number): [number, number] | undefined {
  if (count > 0) {
    if (a >= at) {
      a += count;
      b += count;
    } else if (b >= at) {
      b += count;
    }
    if (a >= max) return undefined;
    return [a, Math.min(b, max - 1)];
  }
  const d1 = at, d2 = at - count - 1;
  if (a >= d1 && b <= d2) return undefined;
  const na = a < d1 ? a : a > d2 ? a + count : d1;
  const nb = b < d1 ? b : b > d2 ? b + count : d1 - 1;
  return [na, nb];
}

export interface Shift {
  axis: 'row' | 'col';
  at: number;
  /** Positive to insert, negative to delete. */
  count: number;
}

/** Applies a row/column shift to a reference; undefined when the reference was deleted. */
export function shiftRef(ref: RefToken, shift: Shift): RefToken | undefined {
  if (shift.axis === 'row') {
    if (ref.kind === 'cols') return ref;
    const moved = shiftInterval(ref.r1, ref.r2, shift.at, shift.count, MAX_ROWS);
    return moved && { ...ref, r1: moved[0], r2: moved[1] };
  }
  if (ref.kind === 'rows') return ref;
  const moved = shiftInterval(ref.c1, ref.c2, shift.at, shift.count, MAX_COLS);
  return moved && { ...ref, c1: moved[0], c2: moved[1] };
}

/**
 * Rewrites a formula for a row/column shift on `sheet`. `home` tells whether the formula lives
 * on that sheet, in which case unqualified references point at it too.
 */
export function shiftFormula(formula: string, sheet: string, home: boolean, shift: Shift): string {
  return transformRefs(formula, (ref) => {
    if (ref.foreign) return undefined;
    if (ref.sheet === undefined ? !home : !sameSheet(ref.sheet, sheet)) return undefined;
    const moved = shiftRef(ref, shift);
    if (!moved) return ref.sheet !== undefined ? `${quoteSheetName(ref.sheet)}!#REF!` : '#REF!';
    return formatRef(moved);
  });
}

/** Moves the relative parts of every reference by (dr, dc), as when copying a formula. */
export function translateFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  return transformRefs(formula, (ref) => {
    if (ref.foreign) return undefined;
    const [ar1, ac1, ar2, ac2] = ref.abs;
    const next = { ...ref };
    if (ref.kind !== 'cols') {
      if (!ar1) next.r1 += dr;
      if (!ar2) next.r2 += dr;
    }
    if (ref.kind !== 'rows') {
      if (!ac1) next.c1 += dc;
      if (!ac2) next.c2 += dc;
    }
    const valid =
      (ref.kind === 'cols' || (next.r1 >= 0 && next.r2 >= 0 && next.r1 < MAX_ROWS && next.r2 < MAX_ROWS)) &&
      (ref.kind === 'rows' || (next.c1 >= 0 && next.c2 >= 0 && next.c1 < MAX_COLS && next.c2 < MAX_COLS));
    return valid ? formatRef(next) : '#REF!';
  });
}

/** Points references at a renamed sheet, or at #REF! when `to` is undefined (sheet deleted). */
export function renameSheetInFormula(formula: string, from: string, to: string | undefined): string {
  return transformRefs(formula, (ref) => {
    if (ref.foreign || !sameSheet(ref.sheet, from)) return undefined;
    return to === undefined ? '#REF!' : formatRef({ ...ref, sheet: to });
  });
}

/** Applies a shift to a space-separated list of ranges such as a `sqref`; undefined when all are gone. */
export function shiftSqref(sqref: string, shift: Shift): string | undefined {
  const parts = sqref
    .split(/\s+/)
    .filter(Boolean)
    .flatMap((part) => {
      const shifted = shiftFormula(part, '', true, shift);
      return shifted.includes('#REF!') ? [] : [shifted];
    });
  return parts.length ? parts.join(' ') : undefined;
}
