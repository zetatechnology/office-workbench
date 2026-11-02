import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { textKindOf, toMarkdown, Workbook } from '../src';

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`../../../test-fixtures/${name}`, import.meta.url)));

describe('documents as Markdown', () => {
  it('converts Word documents to headings, lists and tables', async () => {
    const md = await toMarkdown('docx', fixture('sample.docx'));
    expect(md).toContain('# Office Workbench サンプル文書\n\nこれは Word の表示テスト用の文書です。太字、斜体、赤い文字 and English text.');
    expect(md).toContain('## 箇条書き\n\n- 一つ目の項目\n- 二つ目の項目\n- 三つ目の項目');
    expect(md).toContain('| 項目 | 数量 | 備考 |\n| --- | --- | --- |\n| A | 10 | テスト |');
    expect(md).toContain('**Header:** ヘッダー');
  });

  it('handles numbering, tracked changes and merged table cells', async () => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    zip.file(
      '_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    );
    zip.file(
      'word/_rels/document.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
    );
    zip.file(
      'word/numbering.xml',
      `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    );
    const item = (t: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
    const cell = (body: string, pr = '') => `<w:tc>${pr}<w:p><w:r><w:t>${body}</w:t></w:r></w:p></w:tc>`;
    zip.file(
      'word/document.xml',
      `<w:document xmlns:w="${W}"><w:body>${item('first')}${item('second')}` +
        `<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:del><w:r><w:delText>removed</w:delText></w:r></w:del><w:ins><w:r><w:t>added</w:t></w:r></w:ins></w:p>` +
        `<w:tbl><w:tr>${cell('a|b', '<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>')}${cell('x')}</w:tr><w:tr>${cell('', '<w:tcPr><w:vMerge/></w:tcPr>')}${cell('y')}</w:tr></w:tbl>` +
        '</w:body></w:document>',
    );
    const md = await toMarkdown('docx', await zip.generateAsync({ type: 'uint8array' }));
    expect(md).toBe('1. first\n1. second\n\nkept added\n\n| a\\|b | x |\n| --- | --- |\n|  | y |\n');
  });

  it('converts workbooks to one table per sheet with grouped formulas', async () => {
    const md = await toMarkdown('xlsx', fixture('sample.xlsx'));
    expect(md).toContain('## 売上\n\n|  | A | B | C | D | E |');
    expect(md).toContain('| 4 | りんご | 2026/1/15 | 120 | 98.50 | 11,820.00 |');
    expect(md).toContain('| 1 | 2026年 売上レポート |\n| 3 |');
    expect(md).toContain('| 14 | 折り返し<br>テキスト |  | Red italic |');
    expect(md).toContain('- E4:E7: `=C4*D4`\n- E8: `=SUM(E4:E7)`\n- B12: `=1/0`');
    expect(md).toContain('- B1:B5000: `=ROW()*(COLUMN()-1)+0.5`');
    expect(md).toContain('## Hidden (hidden)\n\n|  | A |\n| --- | --- |\n| 1 | secret |');
  });

  it('shows edits as small line changes', async () => {
    const wb = await Workbook.open(fixture('sample.xlsx'));
    await wb.apply({ type: 'setCells', sheet: 0, cells: [{ r: 3, c: 2, input: '200' }] });
    const before = (await toMarkdown('xlsx', fixture('sample.xlsx'))).split('\n');
    const after = (await toMarkdown('xlsx', await wb.save())).split('\n');
    expect(after.filter((line, i) => line !== before[i])).toEqual(['| 4 | りんご | 2026/1/15 | 200 | 98.50 | 11,820.00 |']);
  });

  it('converts presentations slide by slide with notes', async () => {
    const md = await toMarkdown('pptx', fixture('sample.pptx'));
    expect(md).toContain('## Slide 1\n\nOffice Workbench\n\nPowerPoint 表示テスト\n\n> **Notes:** タイトルスライドのノートです。');
    expect(md).toContain('- 第一レベルの項目\n  - 第二レベルの項目');
    expect(md).toContain('| 地域 | 売上 | 前年比 |\n| --- | --- | --- |\n| 東日本 | 1,200 | 105% |');
    expect(md).toContain('## Slide 5 (hidden)');
  });

  it('recognizes convertible file types', () => {
    expect(['a.DOCX', 'b.xlsm', 'c.pptx', 'd.pdf', 'e.csv'].map(textKindOf)).toEqual(['docx', 'xlsx', 'pptx', undefined, undefined]);
  });
});
