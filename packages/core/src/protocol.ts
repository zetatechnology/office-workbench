// Messages exchanged between the extension host and the viewer webviews.
import type { DocxBlock, DocxEditOp } from './docx/types';
import type { PptxEditOp, Presentation } from './pptx/types';
import type { CellAddress, CellData, CellStyle, EditOp, SheetData, WorkbookInfo } from './xlsx/types';

export type ViewerKind = 'docx' | 'xlsx' | 'pptx' | 'pdf';

/** Any edit a view can send; each editor accepts the operations of its own format. */
export type AnyEditOp = EditOp | DocxEditOp | PptxEditOp;

export type WebviewMessage =
  | { type: 'ready' }
  | { type: 'loadSheet'; index: number }
  | { type: 'edit'; op: AnyEditOp };

export interface MediaPayload {
  path: string;
  contentType: string;
  data: Uint8Array;
}

export type HostMessage =
  /** The document to render; `blocks` is its editable outline when it can be edited. */
  | { type: 'docx'; fileName: string; data: Uint8Array; blocks?: DocxBlock[] }
  | { type: 'pdf'; fileName: string; data: Uint8Array }
  /** Workbook structure; sent on load and after structural edits. Views reload their sheets. */
  | {
      type: 'workbook';
      fileName: string;
      info: WorkbookInfo;
      editable: boolean;
      /** csv: a single sheet of plain text values (no formulas, formats or sheet operations). */
      format: 'xlsx' | 'csv';
      /** Short description of the file shown in the footer (e.g. delimiter and encoding). */
      details?: string;
      activate?: number;
    }
  | { type: 'sheet'; sheet: SheetData }
  /** Updated cells of one sheet after an edit, undo or redo. */
  | { type: 'cells'; sheet: number; cells: CellData[]; removed: CellAddress[]; styles: Record<number, CellStyle> }
  | {
      type: 'presentation';
      fileName: string;
      presentation: Presentation;
      media: MediaPayload[];
      editable?: boolean;
      /** Sent after an edit: media is omitted (the view keeps what it has). */
      update?: boolean;
      /** Slide to show (after slide operations). */
      activate?: number;
    }
  /** A non-fatal problem to show briefly (e.g. an edit that was refused). */
  | { type: 'notice'; message: string }
  | { type: 'error'; message: string };
