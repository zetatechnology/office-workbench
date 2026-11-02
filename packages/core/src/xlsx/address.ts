// A1-style cell reference helpers. Dependency-free so webviews can import them cheaply.
import type { MergeRange } from './types';

/** "AB12" → 27 (0-based column). */
export function columnIndex(ref: string): number {
  let n = 0;
  for (let i = 0; i < ref.length; i++) {
    const code = ref.charCodeAt(i);
    if (code >= 65 && code <= 90) n = n * 26 + (code - 64);
    else if (code >= 97 && code <= 122) n = n * 26 + (code - 96);
    else break;
  }
  return n - 1;
}

/** "AB12" → 11 (0-based row). */
export function rowIndex(ref: string): number {
  const m = /\d+/.exec(ref);
  return m ? Number(m[0]) - 1 : -1;
}

export function parseRange(ref: string): MergeRange | undefined {
  const [a, b = a] = ref.split(':');
  if (!a) return undefined;
  const r1 = rowIndex(a), c1 = columnIndex(a), r2 = rowIndex(b), c2 = columnIndex(b);
  if (r1 < 0 || c1 < 0 || r2 < 0 || c2 < 0) return undefined;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

/** 0 → "A", 27 → "AB". */
export function columnName(index: number): string {
  let n = index + 1;
  let name = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

