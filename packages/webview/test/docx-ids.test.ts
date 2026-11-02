import { describe, expect, it } from 'vitest';
import { moveId, siblingId } from '../src/docx/ids';

describe('Word outline ids', () => {
  it('moves ids past an insert in the same body or cell only', () => {
    const insert = { type: 'insertParagraph', after: '3', text: '' } as const;
    expect(moveId('2', insert)).toBe('2');
    expect(moveId('3', insert)).toBe('3');
    expect(moveId('4', insert)).toBe('5');
    // A table after the insert moves as a whole; paragraphs inside cells keep their own positions.
    expect(moveId('7/1/2/0', insert)).toBe('8/1/2/0');
    const inCell = { type: 'insertParagraph', after: '7/1/2/0', text: '' } as const;
    expect(moveId('7/1/2/1', inCell)).toBe('7/1/2/2');
    expect(moveId('7/1/3/1', inCell)).toBe('7/1/3/1');
    expect(moveId('9', inCell)).toBe('9');
  });

  it('removes the deleted paragraph and moves the ones after it back', () => {
    const del = { type: 'deleteParagraph', id: '6' } as const;
    expect(moveId('6', del)).toBeUndefined();
    expect(moveId('7', del)).toBe('6');
    expect(moveId('5', del)).toBe('5');
    expect(moveId('4', { type: 'setParagraphText', id: '1', text: '' })).toBe('4');
    expect(siblingId('7/1/2/0', 1)).toBe('7/1/2/1');
  });
});
