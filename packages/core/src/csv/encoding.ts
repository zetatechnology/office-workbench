// Character encodings of delimited text files: detection on read, and the same encoding on write.
// Shift_JIS (Windows-31J) matters because Excel on Japanese Windows writes CSV in it.

export type TextEncodingName = 'utf-8' | 'utf-8-bom' | 'shift_jis' | 'utf-16le' | 'utf-16be';

export const ENCODING_LABELS: Record<TextEncodingName, string> = {
  'utf-8': 'UTF-8',
  'utf-8-bom': 'UTF-8 with BOM',
  shift_jis: 'Shift_JIS',
  'utf-16le': 'UTF-16 LE',
  'utf-16be': 'UTF-16 BE',
};

export function decodeText(data: Uint8Array): { text: string; encoding: TextEncodingName } {
  if (data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf) {
    return { text: new TextDecoder('utf-8').decode(data.subarray(3)), encoding: 'utf-8-bom' };
  }
  if (data[0] === 0xff && data[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(data.subarray(2)), encoding: 'utf-16le' };
  if (data[0] === 0xfe && data[1] === 0xff) return { text: new TextDecoder('utf-16be').decode(data.subarray(2)), encoding: 'utf-16be' };
  try {
    return { text: new TextDecoder('utf-8', { fatal: true }).decode(data), encoding: 'utf-8' };
  } catch {
    return { text: new TextDecoder('shift_jis').decode(data), encoding: 'shift_jis' };
  }
}

export class EncodingError extends Error {}

/** Encodes text; throws EncodingError naming the first character the encoding cannot represent. */
export function encodeText(text: string, encoding: TextEncodingName): Uint8Array {
  switch (encoding) {
    case 'utf-8':
      return new TextEncoder().encode(text);
    case 'utf-8-bom':
      return concat(Uint8Array.of(0xef, 0xbb, 0xbf), new TextEncoder().encode(text));
    case 'utf-16le':
    case 'utf-16be': {
      const out = new Uint8Array(2 + text.length * 2);
      const view = new DataView(out.buffer);
      const little = encoding === 'utf-16le';
      view.setUint16(0, 0xfeff, little);
      for (let i = 0; i < text.length; i++) view.setUint16(2 + i * 2, text.charCodeAt(i), little);
      return out;
    }
    case 'shift_jis':
      return encodeShiftJis(text);
  }
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array {
  const out = new Uint8Array(a.length + b.length);
  out.set(a);
  out.set(b, a.length);
  return out;
}

let sjisTable: Map<number, number> | undefined;

/** Code point → two-byte Windows-31J code, built by inverting the platform's decoder. */
function shiftJisTable(): Map<number, number> {
  if (sjisTable) return sjisTable;
  const table = new Map<number, number>();
  const decoder = new TextDecoder('shift_jis');
  for (let lead = 0x81; lead <= 0xfc; lead++) {
    // 0xA0–0xDF are single bytes; 0xED/0xEE duplicate the IBM extensions at 0xFA–0xFC (NEC-selected),
    // which encoders are expected to avoid.
    if ((lead >= 0xa0 && lead <= 0xdf) || lead === 0xed || lead === 0xee) continue;
    for (let trail = 0x40; trail <= 0xfc; trail++) {
      if (trail === 0x7f) continue;
      const ch = decoder.decode(Uint8Array.of(lead, trail));
      const cp = ch.codePointAt(0)!;
      if (ch.length === 1 && cp !== 0xfffd && !table.has(cp)) table.set(cp, (lead << 8) | trail);
    }
  }
  // Characters the WHATWG encoder maps onto neighbours of what the decoder produces.
  table.set(0xa5, 0x5c);
  table.set(0x203e, 0x7e);
  table.set(0x2212, table.get(0xff0d)!);
  sjisTable = table;
  return table;
}

function encodeShiftJis(text: string): Uint8Array {
  const table = shiftJisTable();
  const out = new Uint8Array(text.length * 2);
  let n = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (cp < 0x80) out[n++] = cp;
    else if (cp >= 0xff61 && cp <= 0xff9f) out[n++] = cp - 0xff61 + 0xa1;
    else {
      const code = table.get(cp);
      if (code === undefined) throw new EncodingError(`"${ch}" cannot be saved in Shift_JIS.`);
      out[n++] = code >> 8;
      out[n++] = code & 0xff;
    }
  }
  return out.slice(0, n);
}
