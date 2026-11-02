// PowerPoint edits as language models express them: 1-based slide numbers and the shape ids of an
// outline listing, translated into the editor's PptxEditOps.
import { EditError } from '../errors';
import type { PptxEditOp, Presentation, TextBody } from './types';

export type PresentationToolEdit =
  /** `paragraphs` (or `text`, split at line breaks) replaces all text of the shape. */
  | { type: 'setText'; slide: number; shape: string; paragraphs?: string[]; text?: string }
  | { type: 'setTableCell'; slide: number; shape: string; row: number; column: number; paragraphs?: string[]; text?: string }
  /** Replaces text in every editable shape and table cell (of one slide, when given) that contains `find`. */
  | { type: 'replaceText'; find: string; replace: string; slide?: number }
  | { type: 'duplicateSlide' | 'deleteSlide'; slide: number }
  | { type: 'moveSlide'; slide: number; to: number };

const paragraphsOf = (body: TextBody | undefined) => body?.paragraphs.map((p) => p.runs.map((r) => (r.br ? '\n' : r.text)).join('')) ?? [];

/** Slides with their editable shapes and table cells, for a model to refer to. */
export function presentationOutline(presentation: Presentation): string {
  const out: string[] = [];
  presentation.slides.forEach((slide, i) => {
    out.push(`Slide ${i + 1}${slide.hidden ? ' (hidden)' : ''}:`);
    for (const element of slide.elements) {
      if (element.type === 'shape' && element.textId) {
        out.push(`  shape ${element.textId}: ${JSON.stringify(paragraphsOf(element.text))}`);
      } else if (element.type === 'table' && element.textId) {
        out.push(`  shape ${element.textId} (table, ${element.rows.length} rows × ${element.columns.length} columns):`);
        element.rows.forEach((row, r) =>
          row.cells.forEach((cell, c) => {
            if (!cell.merged) out.push(`    row ${r + 1}, column ${c + 1}: ${JSON.stringify(paragraphsOf(cell.text))}`);
          }),
        );
      }
    }
    if (slide.notes) out.push(`  notes (not editable): ${JSON.stringify(slide.notes)}`);
  });
  return out.join('\n');
}

/** Converts one tool edit against the presentation as it is now (slide numbers shift as slides change). */
export function toPptxOps(edit: PresentationToolEdit, presentation: Presentation): PptxEditOp[] {
  const count = presentation.slides.length;
  const slideIndex = (n: number | undefined, name = 'slide') => {
    if (!Number.isInteger(n) || n! < 1 || n! > count) throw new EditError(`${name} must be a slide number from 1 to ${count} (got ${n}).`);
    return n! - 1;
  };
  const paragraphs = (e: { paragraphs?: string[]; text?: string }) => {
    if (Array.isArray(e.paragraphs) && e.paragraphs.every((p) => typeof p === 'string')) return e.paragraphs;
    if (typeof e.text === 'string') return e.text.replace(/\r\n?/g, '\n').split('\n');
    throw new EditError('Provide paragraphs (a list of strings) or text.');
  };
  const shapeOn = (slide: number, shape: string) => {
    const element = presentation.slides[slide].elements.find((e) => (e.type === 'shape' || e.type === 'table') && e.textId === String(shape));
    if (!element) {
      const ids = presentation.slides[slide].elements.flatMap((e) => ((e.type === 'shape' || e.type === 'table') && e.textId ? [e.textId] : []));
      throw new EditError(`Slide ${slide + 1} has no editable shape "${shape}". Editable shapes: ${ids.join(', ') || 'none'}.`);
    }
    return element;
  };
  switch (edit.type) {
    case 'setText': {
      const slide = slideIndex(edit.slide);
      if (shapeOn(slide, edit.shape).type === 'table') throw new EditError(`Shape ${edit.shape} is a table; use setTableCell.`);
      return [{ type: 'setShapeText', slide, shape: String(edit.shape), paragraphs: paragraphs(edit) }];
    }
    case 'setTableCell': {
      const slide = slideIndex(edit.slide);
      const table = shapeOn(slide, edit.shape);
      if (table.type !== 'table') throw new EditError(`Shape ${edit.shape} is not a table; use setText.`);
      const row = edit.row - 1, col = edit.column - 1;
      if (!table.rows[row]?.cells[col]) {
        throw new EditError(`The table has ${table.rows.length} rows and ${table.columns.length} columns (got row ${edit.row}, column ${edit.column}).`);
      }
      return [{ type: 'setTableCellText', slide, shape: String(edit.shape), row, col, paragraphs: paragraphs(edit) }];
    }
    case 'replaceText': {
      if (!edit.find) throw new EditError('find must not be empty.');
      const only = edit.slide === undefined ? undefined : slideIndex(edit.slide);
      const swap = (list: string[]) => list.map((p) => p.split(edit.find).join(edit.replace ?? ''));
      const ops: PptxEditOp[] = [];
      presentation.slides.forEach((slide, s) => {
        if (only !== undefined && s !== only) return;
        for (const element of slide.elements) {
          if (element.type === 'shape' && element.textId) {
            const current = paragraphsOf(element.text);
            if (current.some((p) => p.includes(edit.find))) ops.push({ type: 'setShapeText', slide: s, shape: element.textId, paragraphs: swap(current) });
          } else if (element.type === 'table' && element.textId) {
            element.rows.forEach((row, r) =>
              row.cells.forEach((cell, c) => {
                const current = paragraphsOf(cell.text);
                if (!cell.merged && current.some((p) => p.includes(edit.find))) {
                  ops.push({ type: 'setTableCellText', slide: s, shape: element.textId!, row: r, col: c, paragraphs: swap(current) });
                }
              }),
            );
          }
        }
      });
      if (!ops.length) throw new EditError(`"${edit.find}" was not found in any editable shape or table cell.`);
      return ops;
    }
    case 'duplicateSlide':
    case 'deleteSlide':
      return [{ type: edit.type, slide: slideIndex(edit.slide) }];
    case 'moveSlide':
      return [{ type: 'moveSlide', slide: slideIndex(edit.slide), to: slideIndex(edit.to, 'to') }];
    default:
      throw new EditError(`Unknown edit type "${(edit as { type?: string }).type}".`);
  }
}

export function describePresentationEdit(edit: PresentationToolEdit): string {
  const text = (e: { paragraphs?: string[]; text?: string }) => {
    const value = e.paragraphs?.join(' / ') ?? e.text ?? '';
    return JSON.stringify(value.length > 60 ? `${value.slice(0, 57)}...` : value);
  };
  switch (edit.type) {
    case 'setText':
      return `Set the text of shape ${edit.shape} on slide ${edit.slide} to ${text(edit)}`;
    case 'setTableCell':
      return `Set row ${edit.row}, column ${edit.column} of table ${edit.shape} on slide ${edit.slide} to ${text(edit)}`;
    case 'replaceText':
      return `Replace ${JSON.stringify(edit.find)} with ${JSON.stringify(edit.replace ?? '')}${edit.slide ? ` on slide ${edit.slide}` : ''}`;
    case 'duplicateSlide':
      return `Duplicate slide ${edit.slide}`;
    case 'deleteSlide':
      return `Delete slide ${edit.slide}`;
    case 'moveSlide':
      return `Move slide ${edit.slide} to position ${edit.to}`;
    default:
      return 'Unknown edit';
  }
}
