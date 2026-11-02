// Messages exchanged between the extension host and the viewer webviews.
import type { Presentation } from './pptx/types';
import type { SheetData, WorkbookInfo } from './xlsx/types';

export type ViewerKind = 'docx' | 'xlsx' | 'pptx' | 'pdf';

export type WebviewMessage =
  | { type: 'ready' }
  | { type: 'loadSheet'; index: number };

export interface MediaPayload {
  path: string;
  contentType: string;
  data: Uint8Array;
}

export type HostMessage =
  | { type: 'docx'; fileName: string; data: Uint8Array }
  | { type: 'pdf'; fileName: string; data: Uint8Array }
  | { type: 'workbook'; fileName: string; info: WorkbookInfo }
  | { type: 'sheet'; sheet: SheetData }
  | { type: 'presentation'; fileName: string; presentation: Presentation; media: MediaPayload[] }
  | { type: 'error'; message: string };
