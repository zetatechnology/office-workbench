import { readFileSync } from 'node:fs';
import JSZip from 'jszip';

export const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`../../../test-fixtures/${name}`, import.meta.url)));

/** Raw compressed bytes of every ZIP entry, to check that untouched parts are byte-identical. */
export function rawEntries(data: Uint8Array): Map<string, string> {
  const b = Buffer.from(data);
  const entries = new Map<string, string>();
  for (let i = 0; b.readUInt32LE(i) === 0x04034b50; ) {
    const size = b.readUInt32LE(i + 18), nameLen = b.readUInt16LE(i + 26), extraLen = b.readUInt16LE(i + 28);
    const start = i + 30 + nameLen + extraLen;
    entries.set(b.toString('utf8', i + 30, i + 30 + nameLen), b.subarray(start, start + size).toString('base64'));
    i = start + size;
  }
  return entries;
}

/** Entries whose raw bytes differ between two files (or exist in only one of them). */
export function changedEntries(before: Uint8Array, after: Uint8Array): string[] {
  const a = rawEntries(before), b = rawEntries(after);
  return [...new Set([...a.keys(), ...b.keys()])].filter((name) => a.get(name) !== b.get(name)).sort();
}

export async function part(data: Uint8Array, name: string): Promise<string | undefined> {
  return (await JSZip.loadAsync(data)).file(name)?.async('string');
}
