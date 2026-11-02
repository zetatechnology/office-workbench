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
  text?: TextBody;
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
