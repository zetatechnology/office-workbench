import { describe, expect, it } from 'vitest';
import {
  describeDocumentEdit, describePresentationEdit, parsePresentation, toDocxOps, toOutline, toPptxOps, WordDocument,
} from '../src';
import { fixture } from './helpers';

describe('Word tool edits', () => {
  it('numbers paragraphs in reading order, including table cells', async () => {
    const outline = await toOutline('docx', fixture('sample.docx'));
    const lines = outline.split('\n');
    expect(lines[0]).toBe('1. [heading 1] "Office Workbench サンプル文書"');
    expect(lines[3]).toBe('4. [bulleted list, level 1] "一つ目の項目"');
    expect(lines[8]).toBe('9. [table 1, row 1, column 2] "数量"');
    expect(outline).toContain('[not editable: Paragraphs with pictures or shapes] ""');
  });

  it('translates paragraph numbers and text replacements', async () => {
    const blocks = await (await WordDocument.open(fixture('sample.docx'))).blocks();
    expect(toDocxOps({ type: 'setParagraph', paragraph: 9, text: '個数' }, blocks)).toEqual([{ type: 'setParagraphText', id: '7/0/1/1', text: '個数' }]);
    expect(toDocxOps({ type: 'insertParagraph', after: 4, text: 'x' }, blocks)).toEqual([{ type: 'insertParagraph', after: '3', text: 'x' }]);
    expect(toDocxOps({ type: 'replaceText', find: '項目', replace: 'item' }, blocks).map((op) => op.type === 'setParagraphText' && op.text)).toEqual([
      '一つ目のitem',
      '二つ目のitem',
      '三つ目のitem',
      'item',
    ]);
    expect(() => toDocxOps({ type: 'deleteParagraph', paragraph: 0 }, blocks)).toThrow(/paragraph number from 1 to/);
    expect(() => toDocxOps({ type: 'replaceText', find: 'nowhere', replace: '' }, blocks)).toThrow('"nowhere" was not found.');
    expect(describeDocumentEdit({ type: 'insertParagraph', after: 3, text: 'abc' })).toBe('Insert "abc" after paragraph 3');
  });
});

describe('PowerPoint tool edits', () => {
  it('lists editable shapes and table cells per slide', async () => {
    const outline = await toOutline('pptx', fixture('sample.pptx'));
    expect(outline).toContain('Slide 1:\n  shape 2: ["Office Workbench"]\n  shape 3: ["PowerPoint 表示テスト"]\n  notes (not editable): "タイトルスライドのノートです。"');
    expect(outline).toContain('  shape 5 (table, 3 rows × 3 columns):\n    row 1, column 1: ["地域"]');
    expect(outline).toContain('Slide 5 (hidden):');
  });

  it('translates slide numbers, shapes and replacements', async () => {
    const { presentation } = await parsePresentation(fixture('sample.pptx'));
    expect(toPptxOps({ type: 'setText', slide: 2, shape: '3', text: 'a\nb' }, presentation)).toEqual([
      { type: 'setShapeText', slide: 1, shape: '3', paragraphs: ['a', 'b'] },
    ]);
    expect(toPptxOps({ type: 'setTableCell', slide: 4, shape: '5', row: 2, column: 3, paragraphs: ['110%'] }, presentation)).toEqual([
      { type: 'setTableCellText', slide: 3, shape: '5', row: 1, col: 2, paragraphs: ['110%'] },
    ]);
    expect(toPptxOps({ type: 'replaceText', find: '日本', replace: 'Japan' }, presentation)).toEqual([
      { type: 'setTableCellText', slide: 3, shape: '5', row: 1, col: 0, paragraphs: ['東Japan'] },
      { type: 'setTableCellText', slide: 3, shape: '5', row: 2, col: 0, paragraphs: ['西Japan'] },
    ]);
    expect(toPptxOps({ type: 'moveSlide', slide: 5, to: 1 }, presentation)).toEqual([{ type: 'moveSlide', slide: 4, to: 0 }]);
    expect(() => toPptxOps({ type: 'setText', slide: 2, shape: '9', text: 'x' }, presentation)).toThrow('Slide 2 has no editable shape "9". Editable shapes: 2, 3.');
    expect(() => toPptxOps({ type: 'setText', slide: 4, shape: '5', text: 'x' }, presentation)).toThrow(/use setTableCell/);
    expect(() => toPptxOps({ type: 'deleteSlide', slide: 6 }, presentation)).toThrow('slide must be a slide number from 1 to 5 (got 6).');
    expect(describePresentationEdit({ type: 'setText', slide: 1, shape: '2', paragraphs: ['A', 'B'] })).toBe('Set the text of shape 2 on slide 1 to "A / B"');
  });
});
