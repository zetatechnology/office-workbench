// Word edits as language models express them: paragraphs by their number in a numbered outline
// (body and table cells in reading order, 1-based), translated into the editor's DocxEditOps.
import { EditError } from '../errors';
import type { DocxBlock, DocxEditOp, DocxParagraph } from './types';

export type DocumentToolEdit =
  | { type: 'setParagraph'; paragraph: number; text: string }
  | { type: 'insertParagraph'; after: number; text: string }
  | { type: 'deleteParagraph'; paragraph: number }
  /** Replaces text inside every editable paragraph that contains `find`. */
  | { type: 'replaceText'; find: string; replace: string };

/** Paragraphs in reading order, with where they are (for the outline). */
function numbered(blocks: DocxBlock[]): { paragraph: DocxParagraph; where?: string }[] {
  const out: { paragraph: DocxParagraph; where?: string }[] = [];
  let tables = 0;
  const walk = (list: DocxBlock[], where?: string) => {
    for (const block of list) {
      if (block.type === 'paragraph') out.push({ paragraph: block, where });
      else if (block.type === 'table') {
        const n = ++tables;
        block.rows.forEach((row, r) =>
          row.cells.forEach((cell, c) => {
            if (!cell.merged) walk(cell.blocks, `table ${n}, row ${r + 1}, column ${c + 1}`);
          }),
        );
      }
    }
  };
  walk(blocks);
  return out;
}

/** The paragraphs as numbered lines, for a model to refer to. */
export function documentOutline(blocks: DocxBlock[]): string {
  return numbered(blocks)
    .map(({ paragraph: p, where }, i) => {
      const tags = [
        where,
        p.heading ? `heading ${p.heading}` : undefined,
        p.list ? `${p.list.ordered ? 'numbered' : 'bulleted'} list, level ${p.list.level + 1}` : undefined,
        p.locked ? `not editable: ${p.locked.replace(/ cannot be edited here.*$/, '')}` : undefined,
      ].filter(Boolean);
      return `${i + 1}. ${tags.length ? `[${tags.join('; ')}] ` : ''}${JSON.stringify(p.text)}`;
    })
    .join('\n');
}

/** Converts one tool edit against the document's current outline (numbers shift as paragraphs change). */
export function toDocxOps(edit: DocumentToolEdit, blocks: DocxBlock[]): DocxEditOp[] {
  const list = numbered(blocks).map((n) => n.paragraph);
  const at = (n: number, name: string) => {
    if (!Number.isInteger(n) || n < 1 || n > list.length) {
      throw new EditError(`${name} must be a paragraph number from 1 to ${list.length} (got ${n}).`);
    }
    return list[n - 1];
  };
  const text = (value: unknown) => {
    if (typeof value !== 'string') throw new EditError('text must be a string.');
    return value;
  };
  switch (edit.type) {
    case 'setParagraph':
      return [{ type: 'setParagraphText', id: at(edit.paragraph, 'paragraph').id, text: text(edit.text) }];
    case 'insertParagraph':
      return [{ type: 'insertParagraph', after: at(edit.after, 'after').id, text: text(edit.text) }];
    case 'deleteParagraph':
      return [{ type: 'deleteParagraph', id: at(edit.paragraph, 'paragraph').id }];
    case 'replaceText': {
      if (!edit.find) throw new EditError('find must not be empty.');
      const replace = text(edit.replace ?? '');
      const matches = list.filter((p) => p.text.includes(edit.find));
      const editable = matches.filter((p) => !p.locked);
      if (!editable.length) {
        throw new EditError(
          matches.length
            ? `"${edit.find}" appears only in paragraphs that cannot be edited here (${matches[0].locked})`
            : `"${edit.find}" was not found.`,
        );
      }
      return editable.map((p) => ({ type: 'setParagraphText', id: p.id, text: p.text.split(edit.find).join(replace) }));
    }
    default:
      throw new EditError(`Unknown edit type "${(edit as { type?: string }).type}".`);
  }
}

export function describeDocumentEdit(edit: DocumentToolEdit): string {
  const short = (s: string | undefined) => JSON.stringify((s ?? '').length > 60 ? `${s!.slice(0, 57)}...` : (s ?? ''));
  switch (edit.type) {
    case 'setParagraph':
      return `Set paragraph ${edit.paragraph} to ${short(edit.text)}`;
    case 'insertParagraph':
      return `Insert ${short(edit.text)} after paragraph ${edit.after}`;
    case 'deleteParagraph':
      return `Delete paragraph ${edit.paragraph}`;
    case 'replaceText':
      return `Replace ${short(edit.find)} with ${short(edit.replace)}`;
    default:
      return 'Unknown edit';
  }
}
