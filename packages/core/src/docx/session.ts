// One open Word document behind its views: the rendered file plus its editable outline. After an
// edit only the outline is sent; building the file for the page view waits for the view to ask.
import type { HostMessage, WebviewMessage } from '../protocol';
import { DocumentSession } from '../session';
import { WordDocument } from './document';
import type { DocxBlock, DocxEditOp } from './types';

export class WordSession extends DocumentSession<WordDocument, DocxEditOp> {
  constructor(fileName: string, data: Uint8Array) {
    super(fileName, data, (bytes) => WordDocument.open(bytes));
  }

  /** The editable outline as it is now. */
  blocks(): Promise<DocxBlock[]> {
    return this.inspect((document) => document.blocks());
  }

  protected async view(document: WordDocument, update: object | undefined): Promise<HostMessage[]> {
    if (update) return [{ type: 'docxOutline', blocks: await document.blocks() }];
    return this.full(document);
  }

  protected async request(document: WordDocument, message: WebviewMessage): Promise<HostMessage[]> {
    return message.type === 'preview' ? this.full(document) : [];
  }

  private async full(document: WordDocument): Promise<HostMessage[]> {
    return [{ type: 'docx', fileName: this.fileName, data: await document.bytes(), blocks: await document.blocks() }];
  }
}
