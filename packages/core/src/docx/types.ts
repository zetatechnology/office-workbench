/**
 * The editable outline of a Word document: body paragraphs and tables in reading order.
 * Ids are positions (element index in the body or a table cell, joined with "/" through
 * table/row/cell), so they change when paragraphs are inserted or deleted before them.
 */
export type DocxBlock = DocxParagraph | DocxTable | DocxOther;

export interface DocxParagraph {
  type: 'paragraph';
  id: string;
  /** Plain text; tabs as "\t" and line breaks as "\n". */
  text: string;
  /** Display name of the paragraph style (e.g. "heading 1"). */
  style?: string;
  heading?: number;
  list?: { ordered: boolean; level: number };
  /** Why the paragraph cannot be edited here (fields, hyperlinks, tracked changes, ...). */
  locked?: string;
}

export interface DocxTable {
  type: 'table';
  id: string;
  rows: { cells: DocxTableCell[] }[];
}

export interface DocxTableCell {
  id: string;
  blocks: DocxBlock[];
  colSpan: number;
  /** Continuation of a vertically merged cell (shown empty, not editable). */
  merged: boolean;
}

/** Content shown read-only, such as content controls. */
export interface DocxOther {
  type: 'other';
  id: string;
  label: string;
  text: string;
}

/** Edits of a Word document; each one is a single undo step. */
export type DocxEditOp =
  | { type: 'setParagraphText'; id: string; text: string }
  /** Inserts a paragraph after `after` with its paragraph properties; `update` also replaces the text of `after` (splitting it). */
  | { type: 'insertParagraph'; after: string; text: string; update?: string }
  | { type: 'deleteParagraph'; id: string };
