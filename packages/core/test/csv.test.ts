import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CsvSession, decodeText, encodeText, EncodingError, parseDelimited, serializeDelimited, type HostMessage } from '../src';

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`../../../test-fixtures/${name}`, import.meta.url)));

async function sheetOf(session: CsvSession) {
  const [message] = await session.handle({ type: 'loadSheet', index: 0 });
  if (message.type !== 'sheet') throw new Error(JSON.stringify(message));
  return message.sheet;
}

const valueAt = (cells: { r: number; c: number; v: string }[], r: number, c: number) => cells.find((x) => x.r === r && x.c === c)?.v;

describe('delimited text', () => {
  it('parses quotes, embedded line breaks and ragged rows', () => {
    const doc = parseDelimited('a,"b,1","say ""hi"""\r\n"multi\r\nline",,\r\nlast', ',');
    expect(doc.rows).toEqual([['a', 'b,1', 'say "hi"'], ['multi\r\nline', '', ''], ['last']]);
    expect(doc.eol).toBe('\r\n');
    expect(doc.trailingNewline).toBe(false);
    expect(serializeDelimited(doc)).toBe('a,"b,1","say ""hi"""\r\n"multi\r\nline",,\r\nlast');
    doc.rows[0][0] = 'x y';
    doc.raws[0] = undefined;
    expect(serializeDelimited(doc).split('\r\n')[0]).toBe('x y,"b,1","say ""hi"""');
  });

  it('detects the encodings Excel writes', () => {
    expect(decodeText(fixture('sample.csv'))).toMatchObject({ encoding: 'shift_jis' });
    expect(decodeText(fixture('sample-utf8.csv'))).toMatchObject({ encoding: 'utf-8-bom' });
    expect(decodeText(fixture('sample-utf16.tsv'))).toMatchObject({ encoding: 'utf-16le' });
    for (const name of ['sample.csv', 'sample-utf8.csv', 'sample-utf16.tsv']) {
      expect(decodeText(fixture(name)).text).toContain('ぶどう（長い商品名のテキスト）');
    }
  });

  it('encodes Shift_JIS (Windows-31J) losslessly and rejects what it cannot represent', () => {
    const text = 'あア漢字①㈱～－ｶﾅ\\~';
    expect(new TextDecoder('shift_jis').decode(encodeText(text, 'shift_jis'))).toBe(text);
    expect(() => encodeText('emoji 😀', 'shift_jis')).toThrow(EncodingError);
    expect(decodeText(encodeText('日本語', 'utf-16be'))).toEqual({ text: '日本語', encoding: 'utf-16be' });
  });
});

describe('CsvSession', () => {
  it('saves unedited files byte for byte', async () => {
    for (const name of ['sample.csv', 'sample-utf8.csv', 'sample-utf16.tsv']) {
      const data = fixture(name);
      expect(Buffer.from(await new CsvSession(name, data).save()).equals(Buffer.from(data))).toBe(true);
    }
  });

  it('presents the file as one sheet with plain values', async () => {
    const session = new CsvSession('sample.csv', fixture('sample.csv'));
    const [info] = await session.handle({ type: 'ready' });
    expect(info).toMatchObject({ type: 'workbook', format: 'csv', details: 'Comma separated · Shift_JIS · CRLF' });
    const sheet = await sheetOf(session);
    expect(sheet.rowCount).toBe(14);
    expect(sheet.colCount).toBe(5);
    expect(sheet.cells.find((c) => c.r === 3 && c.c === 4)).toMatchObject({ v: '11,820.00', t: 'n' });
    expect(valueAt(sheet.cells, 13, 0)).toBe('折り返し\nテキスト');
    const tsv = await sheetOf(new CsvSession('sample-utf16.tsv', fixture('sample-utf16.tsv')));
    expect(valueAt(tsv.cells, 2, 0)).toBe('商品');
  });

  it('edits cells and rewrites only the changed rows', async () => {
    const original = decodeText(fixture('sample.csv')).text.split('\r\n');
    const session = new CsvSession('sample.csv', fixture('sample.csv'));
    const edit = await session.edit({ type: 'setCells', sheet: 0, cells: [{ r: 3, c: 2, input: '200' }, { r: 15, c: 1, input: 'a,b' }] });
    expect(edit.messages[0]).toMatchObject({ type: 'cells', cells: [{ r: 3, c: 2, v: '200', t: 'n' }, { r: 15, c: 1, v: 'a,b' }] });
    let lines = decodeText(await session.save()).text.split('\r\n');
    expect(lines[3]).toBe('りんご,2026/1/15,200,98.50,"11,820.00"');
    expect(lines.slice(4, 14)).toEqual(original.slice(4, 14));
    expect(lines.slice(14)).toEqual(['', ',"a,b"', '']);

    await edit.undo();
    expect(Buffer.from(await session.save()).equals(Buffer.from(fixture('sample.csv')))).toBe(true);
    await edit.redo();
    lines = decodeText(await session.save()).text.split('\r\n');
    expect(lines[3]).toContain(',200,');
  });

  it('inserts and deletes rows and columns with undo', async () => {
    const session = new CsvSession('sample-utf8.csv', fixture('sample-utf8.csv'));
    const rows = await session.edit({ type: 'insertRows', sheet: 0, at: 3, count: 2 });
    expect(rows.messages[0].type).toBe('workbook');
    let sheet = await sheetOf(session);
    expect(valueAt(sheet.cells, 5, 0)).toBe('りんご');
    const cols = await session.edit({ type: 'deleteCols', sheet: 0, at: 0, count: 1 });
    sheet = await sheetOf(session);
    expect(valueAt(sheet.cells, 2, 0)).toBe('日付');
    await cols.undo();
    await rows.undo();
    expect(Buffer.from(await session.save()).equals(Buffer.from(fixture('sample-utf8.csv')))).toBe(true);
    await expect(session.edit({ type: 'addSheet' })).rejects.toThrow(/single sheet/);
  });

  it('refuses to save characters the file encoding cannot hold', async () => {
    const session = new CsvSession('sample.csv', fixture('sample.csv'));
    await session.edit({ type: 'setCells', sheet: 0, cells: [{ r: 0, c: 0, input: '😀' }] });
    await expect(session.save()).rejects.toThrow(/cannot be saved in Shift_JIS/);
  });

  it('ignores undo of edits made before a reset', async () => {
    const session = new CsvSession('x.csv', new TextEncoder().encode('a,b\n'));
    const edit = await session.edit({ type: 'setCells', sheet: 0, cells: [{ r: 0, c: 0, input: 'z' }] });
    session.reset(new TextEncoder().encode('c,d\n'));
    expect(await edit.undo()).toEqual([] as HostMessage[]);
    expect(new TextDecoder().decode(await session.save())).toBe('c,d\n');
  });
});
