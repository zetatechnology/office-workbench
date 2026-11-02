import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseInput, parseRange, Workbook, WorkbookSession, type CellData, type SheetData } from '../src';
import { part, rawEntries } from './helpers';
import { formatRef, renameSheetInFormula, shiftFormula, transformRefs, translateFormula } from '../src/xlsx/refs';

const fixture = () => new Uint8Array(readFileSync(new URL('../../../test-fixtures/sample.xlsx', import.meta.url)));

function cellAt(sheet: SheetData, address: string): CellData | undefined {
  const range = parseRange(address)!;
  return sheet.cells.find((c) => c.r === range.r1 && c.c === range.c1);
}

const at = (address: string) => ({ r: parseRange(address)!.r1, c: parseRange(address)!.c1 });

describe('formula references', () => {
  it('finds references outside strings, names and functions', () => {
    const seen: string[] = [];
    transformRefs(`SUM(A1:B2)+LOG10(C3)+"D4"+'My Sheet'!$E$5+Sheet2!F:F+3:3+Table1[Col]+[1]Ext!G7+売上!H8`, (ref) => {
      seen.push(`${ref.sheet ?? ''}|${formatRef(ref)}|${ref.foreign}`);
      return undefined;
    });
    expect(seen).toEqual([
      '|A1:B2|false',
      '|C3|false',
      "My Sheet|'My Sheet'!$E$5|false",
      'Sheet2|Sheet2!F:F|false',
      '|3:3|false',
      'Ext|Ext!G7|true',
      '売上|売上!H8|false',
    ]);
  });

  it('shifts references for inserted and deleted rows', () => {
    const ins = { axis: 'row', at: 4, count: 2 } as const;
    expect(shiftFormula('SUM(E4:E7)+E3+E5', 'S', true, ins)).toBe('SUM(E4:E9)+E3+E7');
    expect(shiftFormula('S!A5+T!A5+A5', 'S', false, ins)).toBe('S!A7+T!A5+A5');
    const del = { axis: 'row', at: 4, count: -2 } as const;
    expect(shiftFormula('SUM(E4:E7)+E5+E8+$A$10', 'S', true, del)).toBe('SUM(E4:E5)+#REF!+E6+$A$8');
    expect(shiftFormula('SUM(A:A)+SUM(5:6)', 'S', true, del)).toBe('SUM(A:A)+SUM(#REF!)');
    expect(shiftFormula('B1+D1', 'S', true, { axis: 'col', at: 2, count: 1 })).toBe('B1+E1');
  });

  it('translates relative references and renames sheets', () => {
    expect(translateFormula('A1+$B$1+C$1+$D1', 2, 1)).toBe('B3+$B$1+D$1+$D3');
    expect(translateFormula('A1', -1, 0)).toBe('#REF!');
    expect(renameSheetInFormula("SUM(Old!A1,'old'!B2,A3)", 'Old', 'New Name')).toBe("SUM('New Name'!A1,'New Name'!B2,A3)");
    expect(renameSheetInFormula('Gone!A1+1', 'Gone', undefined)).toBe('#REF!+1');
  });
});

describe('cell input', () => {
  it('recognizes values the way Excel does', () => {
    expect(parseInput('', 'General')).toEqual({ kind: 'empty' });
    expect(parseInput('=SUM(A1:A3)', 'General')).toEqual({ kind: 'formula', formula: 'SUM(A1:A3)' });
    expect(parseInput("'123", 'General')).toEqual({ kind: 'string', text: '123' });
    expect(parseInput('123', '@')).toEqual({ kind: 'string', text: '123' });
    expect(parseInput('-1.5e3', 'General')).toEqual({ kind: 'number', value: -1500 });
    expect(parseInput('1,234', 'General')).toEqual({ kind: 'number', value: 1234, numFmtId: 3 });
    expect(parseInput('12.5%', 'General')).toEqual({ kind: 'number', value: 0.125, numFmtId: 10 });
    expect(parseInput('2026/1/15', 'General')).toEqual({ kind: 'number', value: 46037, numFmtId: 14 });
    expect(parseInput('2026-01-15', 'yyyy/m/d')).toEqual({ kind: 'number', value: 46037 });
    expect(parseInput('12:30', 'General')).toEqual({ kind: 'number', value: 0.5208333333333334, numFmtId: 20 });
    expect(parseInput('true', 'General')).toEqual({ kind: 'bool', value: true });
    expect(parseInput('#n/a', 'General')).toEqual({ kind: 'error', code: '#N/A' });
    expect(parseInput('2026/2/30', 'General')).toEqual({ kind: 'string', text: '2026/2/30' });
    expect(parseInput('abc', 'General')).toEqual({ kind: 'string', text: 'abc' });
  });

  it('exposes editable text for stored values', async () => {
    const sheet = await (await Workbook.open(fixture())).sheet(0);
    expect(cellAt(sheet, 'B4')).toMatchObject({ v: '2026/1/15' });
    expect(cellAt(sheet, 'B4')?.input).toBeUndefined();
    expect(cellAt(sheet, 'C6')).toMatchObject({ v: '1,500', input: '1500' });
    expect(cellAt(sheet, 'B10')?.input).toBeUndefined();
    expect(cellAt(sheet, 'E8')).toMatchObject({ input: '=SUM(E4:E7)' });
    const data = await (await Workbook.open(fixture())).sheet(1);
    expect(cellAt(data, 'C2')).toMatchObject({ f: 'ROW()*(COLUMN()-1)+0.5' });
    expect(cellAt(data, 'A3')).toMatchObject({ f: '"Item "&ROW()' });
  });
});

describe('editing workbooks', () => {
  it('saves cell edits and keeps untouched parts byte-identical', async () => {
    const original = fixture();
    const wb = await Workbook.open(original);
    await wb.apply({
      type: 'setCells',
      sheet: 0,
      cells: [
        { ...at('C4'), input: '200' },
        { ...at('A13'), input: '  メモ\n2行目' },
        { ...at('G2'), input: '=SUM(C4:C7)' },
        { ...at('G3'), input: '2026/3/1' },
        { ...at('A10'), input: '' },
      ],
    });
    const saved = await wb.save();

    const before = rawEntries(original), after = rawEntries(saved);
    const changed = [...before.keys()].filter((name) => after.has(name) && before.get(name) !== after.get(name));
    expect(changed.sort()).toEqual(['[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
    expect(after.has('xl/calcChain.xml')).toBe(false);
    expect(await part(saved, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');

    const sheet = await (await Workbook.open(saved)).sheet(0);
    expect(cellAt(sheet, 'C4')).toMatchObject({ v: '200', t: 'n' });
    expect(cellAt(sheet, 'A13')).toMatchObject({ v: '  メモ\n2行目', t: 's' });
    expect(cellAt(sheet, 'G2')).toMatchObject({ f: 'SUM(C4:C7)', v: '' });
    expect(cellAt(sheet, 'G3')).toMatchObject({ v: '2026/3/1', t: 'd' });
    expect(cellAt(sheet, 'A10')).toBeUndefined();
    expect(cellAt(sheet, 'B10')).toMatchObject({ v: '87.5%' });
    const xml = (await part(saved, 'xl/worksheets/sheet1.xml'))!;
    expect(xml).toContain('<dimension ref="A1:G14"/>');
    expect(xml).toContain('x14ac:dyDescent');
  });

  it('reverts cell edits, including expanded shared formulas', async () => {
    const wb = await Workbook.open(fixture());
    const edit = await wb.apply({ type: 'setCells', sheet: 1, cells: [{ ...at('C2'), input: '42' }] });
    expect(edit.change).toMatchObject({ kind: 'cells', sheet: 1 });
    let sheet = await wb.sheet(1);
    expect(cellAt(sheet, 'C2')).toMatchObject({ v: '42' });
    expect(cellAt(sheet, 'D3')).toMatchObject({ f: 'ROW()*(COLUMN()-1)+0.5', v: '9.5' });
    await edit.revert();
    sheet = await wb.sheet(1);
    expect(cellAt(sheet, 'C2')).toMatchObject({ v: '4.5', f: 'ROW()*(COLUMN()-1)+0.5' });
    const xml = (await part(await wb.save(), 'xl/worksheets/sheet2.xml'))!;
    expect(xml).toContain('<c r="C2"><f t="shared" si="1"/><v>4.5</v></c>');
  });

  it('inserts and deletes rows, moving cells, formulas, merges and chart ranges', async () => {
    const wb = await Workbook.open(fixture());
    const insert = await wb.apply({ type: 'insertRows', sheet: 0, at: 4, count: 2 });
    let sheet = await wb.sheet(0);
    expect(cellAt(sheet, 'E10')).toMatchObject({ f: 'SUM(E4:E9)' });
    expect(cellAt(sheet, 'E7')).toMatchObject({ f: 'C7*D7' });
    expect(cellAt(sheet, 'A5')).toBeUndefined();
    expect(sheet.rowHeights[15]).toBe(48);
    let saved = await wb.save();
    expect(await part(saved, 'xl/charts/chart1.xml')).toContain('<c:f>売上!$A$4:$A$9</c:f>');

    await insert.revert();
    sheet = await wb.sheet(0);
    expect(cellAt(sheet, 'E8')).toMatchObject({ f: 'SUM(E4:E7)' });

    await wb.apply({ type: 'deleteRows', sheet: 0, at: 0, count: 1 });
    sheet = await wb.sheet(0);
    expect(sheet.merges).toEqual([]);
    expect(sheet.frozen).toEqual({ rows: 2, cols: 0 });
    expect(cellAt(sheet, 'E7')).toMatchObject({ f: 'SUM(E3:E6)' });
    saved = await wb.save();
    const xml = (await part(saved, 'xl/worksheets/sheet1.xml'))!;
    expect(xml).not.toContain('mergeCells');
    expect(xml).toContain('<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>');
  });

  it('inserts columns and refuses to push data off the sheet', async () => {
    const wb = await Workbook.open(fixture());
    await wb.apply({ type: 'insertCols', sheet: 0, at: 2, count: 1 });
    const sheet = await wb.sheet(0);
    expect(cellAt(sheet, 'F4')).toMatchObject({ f: 'D4*E4' });
    expect(sheet.merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 5 }]);
    await wb.apply({ type: 'setCells', sheet: 0, cells: [{ r: 0, c: 16_383, input: 'edge' }] });
    await expect(wb.apply({ type: 'insertCols', sheet: 0, at: 0, count: 1 })).rejects.toThrow(/pushed off/);
  });

  it('adds, renames and deletes sheets', async () => {
    const wb = await Workbook.open(fixture());
    const add = await wb.apply({ type: 'addSheet' });
    expect(add.change).toEqual({ kind: 'structure', activate: 3 });
    expect(wb.info().sheets.map((s) => s.name)).toEqual(['売上', 'Data', 'Hidden', 'Sheet4']);
    await wb.apply({ type: 'setCells', sheet: 3, cells: [{ r: 0, c: 0, input: '=売上!E8*2' }] });
    await expect(wb.apply({ type: 'renameSheet', sheet: 3, name: 'data' })).rejects.toThrow(/already exists/);

    await wb.apply({ type: 'renameSheet', sheet: 0, name: 'Sales 2026' });
    expect(cellAt(await wb.sheet(3), 'A1')).toMatchObject({ f: "'Sales 2026'!E8*2" });
    let saved = await wb.save();
    expect(await part(saved, 'xl/charts/chart1.xml')).toContain("<c:f>'Sales 2026'!$A$4:$A$7</c:f>");

    await wb.apply({ type: 'deleteSheet', sheet: 0 });
    expect(wb.info().sheets.map((s) => s.name)).toEqual(['Data', 'Hidden', 'Sheet4']);
    expect(cellAt(await wb.sheet(2), 'A1')).toMatchObject({ f: '#REF!*2' });
    saved = await wb.save();
    const zip = await JSZip.loadAsync(saved);
    expect(zip.file('xl/worksheets/sheet1.xml')).toBeNull();
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/xl/worksheets/sheet4.xml');
    expect(await zip.file('[Content_Types].xml')!.async('string')).not.toContain('/xl/worksheets/sheet1.xml');

    const reopened = await Workbook.open(saved);
    expect(reopened.info().sheets.map((s) => s.name)).toEqual(['Data', 'Hidden', 'Sheet4']);
    await expect(reopened.apply({ type: 'deleteSheet', sheet: 0 })).resolves.toBeDefined();
    await expect(reopened.apply({ type: 'deleteSheet', sheet: 1 })).rejects.toThrow(/at least one visible/);
  });

  it('undoes structural edits by restoring the touched parts', async () => {
    const wb = await Workbook.open(fixture());
    const pristine = rawEntries(fixture());
    const steps = [
      await wb.apply({ type: 'addSheet', name: 'New' }),
      await wb.apply({ type: 'renameSheet', sheet: 1, name: 'Numbers' }),
      await wb.apply({ type: 'deleteCols', sheet: 0, at: 1, count: 2 }),
      await wb.apply({ type: 'deleteSheet', sheet: 2 }),
    ];
    for (const step of steps.reverse()) await step.revert();
    expect(wb.info().sheets.map((s) => s.name)).toEqual(['売上', 'Data', 'Hidden']);
    const sheet = await wb.sheet(0);
    expect(cellAt(sheet, 'E4')).toMatchObject({ f: 'C4*D4' });
    const saved = await wb.save();
    expect([...rawEntries(saved).keys()].sort()).toEqual([...pristine.keys()].filter((n) => n !== 'xl/calcChain.xml').sort());
    for (const name of ['xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/charts/chart1.xml']) {
      expect(await part(saved, name)).toBe(await part(fixture(), name));
    }
  });
});

describe('WorkbookSession', () => {
  it('reports cell patches and supports undo/redo', async () => {
    const session = new WorkbookSession('sample.xlsx', fixture());
    const edit = await session.edit({ type: 'setCells', sheet: 0, cells: [{ ...at('C4'), input: '1,000' }] });
    expect(edit.label).toBe('Edit cell');
    const [message] = edit.messages;
    expect(message.type).toBe('cells');
    if (message.type === 'cells') {
      expect(message.cells).toMatchObject([{ r: 3, c: 2, v: '1,000' }]);
      expect(message.styles[message.cells[0].s]).toBeDefined();
    }
    const [undone] = await edit.undo();
    expect(undone.type === 'cells' && undone.cells[0].v).toBe('120');
    const [redone] = await edit.redo();
    expect(redone.type === 'cells' && redone.cells[0].v).toBe('1,000');
    await expect(session.edit({ type: 'renameSheet', sheet: 0, name: 'a/b' })).rejects.toThrow(/cannot contain/);
  });
});
