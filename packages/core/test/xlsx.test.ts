import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { columnIndex, columnName, parseRange, Workbook, type CellData, type SheetData } from '../src';

const fixture = () => readFileSync(new URL('../../../test-fixtures/sample.xlsx', import.meta.url));

function cellAt(sheet: SheetData, address: string): CellData | undefined {
  const range = parseRange(address)!;
  return sheet.cells.find((c) => c.r === range.r1 && c.c === range.c1);
}

describe('Workbook (Excel-generated fixture)', () => {
  it('lists worksheets with visibility', async () => {
    const wb = await Workbook.open(fixture());
    expect(wb.info().sheets).toEqual([
      { index: 0, name: '売上', state: 'visible' },
      { index: 1, name: 'Data', state: 'visible' },
      { index: 2, name: 'Hidden', state: 'hidden' },
    ]);
  });

  it('formats values with their number formats', async () => {
    const sheet = await (await Workbook.open(fixture())).sheet(0);
    expect(cellAt(sheet, 'A1')).toMatchObject({ v: '2026年 売上レポート', t: 's' });
    expect(cellAt(sheet, 'B4')).toMatchObject({ v: '2026/1/15', t: 'd' });
    expect(cellAt(sheet, 'C6')).toMatchObject({ v: '1,500', t: 'n' });
    expect(cellAt(sheet, 'D4')).toMatchObject({ v: '98.50' });
    expect(cellAt(sheet, 'E4')).toMatchObject({ v: '11,820.00', f: 'C4*D4' });
    expect(cellAt(sheet, 'E8')).toMatchObject({ v: '¥75,195', f: 'SUM(E4:E7)' });
    expect(cellAt(sheet, 'B10')).toMatchObject({ v: '87.5%' });
    expect(cellAt(sheet, 'B11')).toMatchObject({ v: 'TRUE', t: 'b' });
    expect(cellAt(sheet, 'B12')).toMatchObject({ v: '#DIV/0!', t: 'e' });
  });

  it('resolves styles, layout and sheet view', async () => {
    const sheet = await (await Workbook.open(fixture())).sheet(0);
    const header = sheet.styles[cellAt(sheet, 'A3')!.s];
    expect(header).toMatchObject({ bold: true, color: '#ffffff', background: '#1f3f7f', hAlign: 'center' });
    expect(header.borderTop).toBe('1px solid #000000');
    expect(sheet.styles[cellAt(sheet, 'A14')!.s].wrap).toBe(true);
    expect(sheet.styles[cellAt(sheet, 'C14')!.s]).toMatchObject({ italic: true, color: '#ff0000' });
    expect(sheet.merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 4 }]);
    expect(sheet.frozen).toEqual({ rows: 3, cols: 0 });
    expect(sheet.colWidths[0]).toBeGreaterThan(sheet.defaultColWidth);
    expect(sheet.rowHeights[13]).toBe(48);
    expect(sheet.unsupported).toEqual(['1 drawing object(s) (charts, images or shapes)']);
  });

  it('reads formula sheets', async () => {
    const sheet = await (await Workbook.open(fixture())).sheet(1);
    expect(sheet.rowCount).toBe(5000);
    expect(sheet.colCount).toBe(6);
    expect(sheet.cells).toHaveLength(30_000);
    expect(cellAt(sheet, 'A5000')?.v).toBe('Item 5000');
    expect(cellAt(sheet, 'F10')?.v).toBe('50.5');
  });

  it('parses a 100,000-cell sheet quickly', async () => {
    const rows: string[] = [];
    for (let r = 1; r <= 10_000; r++) {
      const cells = Array.from({ length: 10 }, (_, c) => `<c r="${columnName(c)}${r}"><v>${r * c}</v></c>`).join('');
      rows.push(`<row r="${r}">${cells}</row>`);
    }
    const data = await buildWorkbook(`<sheetData>${rows.join('')}</sheetData>`);
    const started = performance.now();
    const sheet = await (await Workbook.open(data)).sheet(0);
    const elapsed = performance.now() - started;
    expect(sheet.cells).toHaveLength(100_000);
    expect(elapsed).toBeLessThan(5000);
  });

  it('rejects files that are not workbooks', async () => {
    await expect(Workbook.open(new Uint8Array([1, 2, 3]))).rejects.toThrow(/not a valid Office Open XML/i);
  });
});

describe('cell addresses', () => {
  it('converts between names and indices', () => {
    expect(columnName(0)).toBe('A');
    expect(columnName(25)).toBe('Z');
    expect(columnName(26)).toBe('AA');
    expect(columnName(16_383)).toBe('XFD');
    expect(columnIndex('XFD1')).toBe(16_383);
    expect(parseRange('B2:A1')).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 });
  });
});

/** A minimal single-sheet workbook package around the given worksheet body. */
async function buildWorkbook(sheetBody: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  );
  zip.file(
    'xl/workbook.xml',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  );
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${sheetBody}</worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}
