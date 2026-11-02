/** All geometry is in CSS pixels at 96 dpi (1 EMU = 1/9525 px). */
export interface Presentation {
  width: number;
  height: number;
  slides: Slide[];
  /** Media referenced by `src` fields, keyed by part path. Bytes are sent separately. */
  media: MediaRef[];
}

export interface MediaRef {
  path: string;
  contentType: string;
}

export interface Slide {
  index: number;
  /** `sldId@id`: stays the same when slides are moved, inserted or deleted. */
  id: string;
  hidden: boolean;
  background: Fill;
  elements: SlideElement[];
  notes?: string;
}

export type Fill =
  | { type: 'none' }
  | { type: 'solid'; color: string }
  | { type: 'gradient'; css: string }
  | { type: 'image'; src: string };

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  /** Degrees, clockwise. */
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export interface Line {
  color: string;
  width: number;
  dash?: string;
  headEnd?: string;
  tailEnd?: string;
}

export type SlideElement = ShapeElement | PictureElement | TableElement | UnsupportedElement;

export interface ShapeElement {
  type: 'shape';
  box: Box;
  /** SVG path data in the shape's local coordinates (0..w, 0..h); absent for text-only boxes. */
  path?: string;
  /** Open paths (lines, arcs) are stroked only. */
  closed: boolean;
  fill: Fill;
  line?: Line;
  /** Present when the shape has visible text, and always for shapes whose text can be edited. */
  text?: TextBody;
  /** `cNvPr@id` of a shape on the slide itself whose text can be edited (not layout/master shapes). */
  textId?: string;
}

export interface PictureElement {
  type: 'picture';
  box: Box;
  /** Media part path, or undefined when the image format cannot be shown. */
  src?: string;
  /** Crop fractions (0..1) from each edge. */
  crop?: { l: number; t: number; r: number; b: number };
  line?: Line;
}

export interface TableElement {
  type: 'table';
  box: Box;
  /** `cNvPr@id` of the table's frame when its cells can be edited. */
  textId?: string;
  columns: number[];
  rows: TableRow[];
}

export interface TableRow {
  height: number;
  cells: TableCell[];
}

export interface TableCell {
  text: TextBody;
  fill: Fill;
  colSpan: number;
  rowSpan: number;
  /** Covered by a neighbouring merged cell. */
  merged: boolean;
  borders: { l?: Line; t?: Line; r?: Line; b?: Line };
}

export interface UnsupportedElement {
  type: 'unsupported';
  box: Box;
  label: string;
}

export interface TextBody {
  paragraphs: Paragraph[];
  anchor: 't' | 'ctr' | 'b';
  /** Pixels. */
  insets: { l: number; t: number; r: number; b: number };
  wrap: boolean;
  vertical?: 'vert' | 'vert270' | 'eaVert';
}

export interface Paragraph {
  align: 'l' | 'ctr' | 'r' | 'just' | 'dist';
  level: number;
  runs: Run[];
  bullet?: string;
  /** Pixels. */
  marginLeft: number;
  indent: number;
  spaceBefore: number;
  spaceAfter: number;
  /** Multiplier (1 = single) or fixed pixels. */
  lineSpacing?: { percent: number } | { px: number };
  /** Font size of the paragraph end mark; keeps empty paragraphs at the right height. */
  endSize: number;
}

export interface Run {
  text: string;
  /** Pixels. */
  size: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  font?: string;
  /** Superscript (+) / subscript (-) offset in percent. */
  baseline?: number;
  /** A line break (`<a:br/>`) instead of text. */
  br?: boolean;
}

/**
 * Edits of a presentation; each one is a single undo step. Shapes are identified by their
 * `cNvPr@id` on the slide. A paragraph's "\n" (or "\v") is a line break within the paragraph.
 */
export type PptxEditOp =
  /** `slideId` (the slide's `id`), when given, wins over `slide`: it still finds the slide after others moved. */
  | { type: 'setShapeText'; slide: number; slideId?: string; shape: string; paragraphs: string[] }
  | { type: 'setTableCellText'; slide: number; slideId?: string; shape: string; row: number; col: number; paragraphs: string[] }
  | { type: 'duplicateSlide'; slide: number }
  | { type: 'deleteSlide'; slide: number }
  | { type: 'moveSlide'; slide: number; to: number };
