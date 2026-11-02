// Paragraph ids of the Word outline are positions ("3", or "7/1/2/0" inside a table cell), so an
// inserted or deleted paragraph moves the ids after it in the same body or cell.
import type { DocxEditOp } from '@office-workbench/core';

/** `op` with its ids moved past an applied insert or delete; undefined when its paragraph was deleted. */
export function remap(op: DocxEditOp, applied: DocxEditOp): DocxEditOp | undefined {
  const move = (id: string) => moveId(id, applied);
  switch (op.type) {
    case 'setParagraphText':
    case 'deleteParagraph': {
      const id = move(op.id);
      return id === undefined ? undefined : { ...op, id };
    }
    case 'insertParagraph': {
      const after = move(op.after);
      return after === undefined ? undefined : { ...op, after };
    }
  }
}

function moveId(id: string, applied: DocxEditOp): string | undefined {
  if (applied.type === 'setParagraphText') return id;
  const at = applied.type === 'insertParagraph' ? applied.after : applied.id;
  if (applied.type === 'deleteParagraph' && (id === at || id.startsWith(`${at}/`))) return undefined;
  const container = at.split('/').slice(0, -1);
  const index = Number(at.split('/').pop());
  const parts = id.split('/');
  // Only ids in the same body or cell, at or below the changed position, move.
  if (parts.length <= container.length || container.some((seg, i) => parts[i] !== seg)) return id;
  const own = Number(parts[container.length]);
  if (own <= index) return id;
  parts[container.length] = String(own + (applied.type === 'insertParagraph' ? 1 : -1));
  return parts.join('/');
}

export function siblingId(id: string, delta: number): string {
  const parts = id.split('/');
  parts[parts.length - 1] = String(Number(parts[parts.length - 1]) + delta);
  return parts.join('/');
}
