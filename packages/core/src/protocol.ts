// Messages exchanged between the extension host and the viewer webviews.
import type { Presentation } from './pptx/types';
import type { CellAddress, CellData, CellStyle, EditOp, SheetData, WorkbookInfo } from './xlsx/types';

export type ViewerKind = 'docx' | 'xlsx' | 'pptx' | 'pdf';

export type WebviewMessage =
  | { type: 'ready' }
  | { type: 'loadSheet'; index: number }
  | { type: 'edit'; op: EditOp };

export interface MediaPayload {
  path: string;
  contentType: string;
  data: Uint8Array;
}

export type HostMessage =
  | { type: 'docx'; fileName: string; data: Uint8Array }
  | { type: 'pdf'; fileName: string; data: Uint8Array }
  /** Workbook structure; sent on load and after structural edits. Views reload their sheets. */
  | { type: 'workbook'; fileName: string; info: WorkbookInfo; editable: boolean; activate?: number }
  | { type: 'sheet'; sheet: SheetData }
  /** Updated cells of one sheet after an edit, undo or redo. */
  | { type: 'cells'; sheet: number; cells: CellData[]; removed: CellAddress[]; styles: Record<number, CellStyle> }
  | { type: 'presentation'; fileName: string; presentation: Presentation; media: MediaPayload[] }
  /** A non-fatal problem to show briefly (e.g. an edit that was refused). */
  | { type: 'notice'; message: string }
  | { type: 'error'; message: string };
