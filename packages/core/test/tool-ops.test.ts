import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { describeEdit, toEditOp, toMarkdown, Workbook, type SheetInfo } from '../src';

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`../../../test-fixtures/${name}`, import.meta.url)));

const sheets: SheetInfo[] = [
  { index: 0, name: 'Hidden', state: 'hidden' },
  { index: 1, name: '売上', state: 'visible' },
  { index: 2, name: 'Data', state: 'visible' },
];

describe('tool edits', () => {
  it('translates A1 addresses, sheet names and 1-based rows', () => {
    expect(toEditOp({ type: 'setCells', cells: { B3: '42', $c$4: 7, D5: null } }, sheets)).toEqual({
      type: 'setCells',
      sheet: 1,
      cells: [
        { r: 2, c: 1, input: '42' },
        { r: 3, c: 2, input: '7' },
        { r: 4, c: 3, input: '' },
      ],
    });
    expect(toEditOp({ type: 'clearRange', sheet: 'data', range: 'A2:C10' }, sheets)).toEqual({
      type: 'clearRange',
      sheet: 2,
      range: { r1: 1, c1: 0, r2: 9, c2: 2 },
    });
    expect(toEditOp({ type: 'insertRows', row: 5, count: 2 }, sheets)).toEqual({ type: 'insertRows', sheet: 1, at: 4, count: 2 });
    expect(toEditOp({ type: 'deleteColumns', sheet: 'Data', column: 'c' }, sheets)).toEqual({ type: 'deleteCols', sheet: 2, at: 2, count: 1 });
    expect(toEditOp({ type: 'renameSheet', sheet: '売上', name: 'Sales' }, sheets)).toEqual({ type: 'renameSheet', sheet: 1, name: 'Sales' });
  });

  it('explains invalid input so a model can correct it', () => {
    expect(() => toEditOp({ type: 'setCells', sheet: 'Nope', cells: { A1: 'x' } }, sheets)).toThrow('No sheet named "Nope". Sheets: Hidden, 売上, Data');
    expect(() => toEditOp({ type: 'setCells', cells: { 'A1:B2': 'x' } }, sheets)).toThrow(/not a single cell/);
    expect(() => toEditOp({ type: 'insertRows', row: 0 }, sheets)).toThrow(/1-based/);
    expect(() => toEditOp({ type: 'insertColumns', column: '3' }, sheets)).toThrow(/column letter/);
    expect(() => toEditOp({ type: 'deleteRows', row: 2, count: -1 }, sheets)).toThrow(/positive integer/);
  });

  it('describes edits for confirmations', () => {
    expect(describeEdit({ type: 'setCells', sheet: '売上', cells: { C4: '200', G2: '=SUM(C4:C7)' } })).toBe(
      'Set 2 cell(s) on "売上": C4 = "200", G2 = "=SUM(C4:C7)"',
    );
    expect(describeEdit({ type: 'insertColumns', column: 'B', count: 3 })).toBe('Insert 3 column(s) at column B');
  });

  it('applies through the workbook editor', async () => {
    const wb = await Workbook.open(fixture('sample.xlsx'));
    await wb.apply(toEditOp({ type: 'setCells', sheet: '売上', cells: { C4: '200' } }, wb.info().sheets));
    const md = await toMarkdown('xlsx', await wb.save(), { sheet: '売上', range: 'A3:C5' });
    expect(md).toBe('## 売上\n\n|  | A | B | C |\n| --- | --- | --- | --- |\n| 3 | 商品 | 日付 | 数量 |\n| 4 | りんご | 2026/1/15 | 200 |\n| 5 | みかん | 2026/2/3 | 80 |\n');
  });

  it('reads one sheet or range of a workbook', async () => {
    const md = await toMarkdown('xlsx', fixture('sample.xlsx'), { sheet: 'data', range: 'B2:C3' });
    expect(md).toBe('## Data\n\n|  | B | C |\n| --- | --- | --- |\n| 2 | 2.5 | 4.5 |\n| 3 | 3.5 | 6.5 |\n\n### Formulas (Data)\n\n- B2:B3: `=ROW()*(COLUMN()-1)+0.5`\n- C2:C3: `=ROW()*(COLUMN()-1)+0.5`\n');
    await expect(toMarkdown('xlsx', fixture('sample.xlsx'), { sheet: 'x' })).rejects.toThrow('No sheet named "x". Sheets: 売上, Data, Hidden');
  });
});
