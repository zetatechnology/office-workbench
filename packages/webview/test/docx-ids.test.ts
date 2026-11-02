import { describe, expect, it } from 'vitest';
import { remap, siblingId } from '../src/docx/ids';

describe('Word outline ids', () => {
  it('moves ids past an insert in the same body or cell only', () => {
    const insert = { type: 'insertParagraph', after: '3', text: '' } as const;
    expect(remap({ type: 'setParagraphText', id: '2', text: 'a' }, insert)).toEqual({ type: 'setParagraphText', id: '2', text: 'a' });
    expect(remap({ type: 'setParagraphText', id: '3', text: 'a' }, insert)).toMatchObject({ id: '3' });
    expect(remap({ type: 'setParagraphText', id: '4', text: 'a' }, insert)).toMatchObject({ id: '5' });
    // A table after the insert moves as a whole; paragraphs inside cells keep their own positions.
    expect(remap({ type: 'deleteParagraph', id: '7/1/2/0' }, insert)).toEqual({ type: 'deleteParagraph', id: '8/1/2/0' });
    const inCell = { type: 'insertParagraph', after: '7/1/2/0', text: '' } as const;
    expect(remap({ type: 'setParagraphText', id: '7/1/2/1', text: 'a' }, inCell)).toMatchObject({ id: '7/1/2/2' });
    expect(remap({ type: 'setParagraphText', id: '7/1/3/1', text: 'a' }, inCell)).toMatchObject({ id: '7/1/3/1' });
    expect(remap({ type: 'setParagraphText', id: '9', text: 'a' }, inCell)).toMatchObject({ id: '9' });
  });

  it('drops edits of a deleted paragraph and moves the ones after it back', () => {
    const del = { type: 'deleteParagraph', id: '6' } as const;
    // A second Backspace in the same (already deleted) paragraph must not delete the next one.
    expect(remap({ type: 'deleteParagraph', id: '6' }, del)).toBeUndefined();
    expect(remap({ type: 'insertParagraph', after: '7', text: 'x' }, del)).toMatchObject({ after: '6' });
    expect(remap({ type: 'setParagraphText', id: '5', text: 'a' }, del)).toMatchObject({ id: '5' });
    expect(remap({ type: 'setParagraphText', id: '4', text: 'a' }, { type: 'setParagraphText', id: '1', text: '' })).toMatchObject({ id: '4' });
    expect(siblingId('7/1/2/0', 1)).toBe('7/1/2/1');
  });
});
