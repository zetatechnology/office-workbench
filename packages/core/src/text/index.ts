// Office documents as Markdown text, for reading in a text editor and for line-based diffs.
import { docxToMarkdown } from './docx';
import { pptxToMarkdown } from './pptx';
import { xlsxToMarkdown, type SheetSelection } from './xlsx';

export type { SheetSelection };

export type TextKind = 'docx' | 'xlsx' | 'pptx';

/** The kind of document by file name, or undefined when it cannot be converted. */
export function textKindOf(fileName: string): TextKind | undefined {
  const ext = fileName.slice(fileName.lastIndexOf('.') + 1).toLowerCase();
  if (ext === 'docx' || ext === 'docm') return 'docx';
  if (ext === 'xlsx' || ext === 'xlsm') return 'xlsx';
  if (ext === 'pptx' || ext === 'pptm') return 'pptx';
  return undefined;
}

/** `selection` narrows workbooks to one sheet and/or range; other kinds ignore it. */
export function toMarkdown(kind: TextKind, data: Uint8Array, selection?: SheetSelection): Promise<string> {
  switch (kind) {
    case 'docx':
      return docxToMarkdown(data);
    case 'xlsx':
      return xlsxToMarkdown(data, selection);
    case 'pptx':
      return pptxToMarkdown(data);
  }
}
