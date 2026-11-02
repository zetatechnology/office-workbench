// One open Word document behind its views: the rendered file plus its editable outline.
import type { HostMessage } from '../protocol';
import { DocumentSession } from '../session';
import { WordDocument } from './document';
import type { DocxEditOp } from './types';

export class WordSession extends DocumentSession<WordDocument, DocxEditOp> {
  constructor(fileName: string, data: Uint8Array) {
    super(fileName, data, (bytes) => WordDocument.open(bytes));
  }

  protected async view(document: WordDocument): Promise<HostMessage[]> {
    return [{ type: 'docx', fileName: this.fileName, data: await document.bytes(), blocks: await document.blocks() }];
  }
}
