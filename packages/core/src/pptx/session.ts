// One open presentation behind its views. After an edit the whole slide model is sent again,
// without image bytes (the views keep the media they already have).
import type { HostMessage } from '../protocol';
import type { Presentation } from './types';
import { DocumentSession } from '../session';
import { PresentationDocument } from './edit';
import type { PptxEditOp } from './types';

export class PresentationSession extends DocumentSession<PresentationDocument, PptxEditOp> {
  constructor(fileName: string, data: Uint8Array) {
    super(fileName, data, (bytes) => PresentationDocument.open(bytes));
  }

  /** The slide model as it is now, without image bytes. */
  presentation(): Promise<Presentation> {
    return this.inspect(async (document) => (await document.read(false)).presentation);
  }

  protected async view(document: PresentationDocument, update: { activate?: number } | undefined): Promise<HostMessage[]> {
    const { presentation, media } = await document.read(!update);
    return [
      {
        type: 'presentation',
        fileName: this.fileName,
        presentation,
        media: presentation.media.flatMap((ref) => {
          const data = media.get(ref.path);
          return data ? [{ ...ref, data }] : [];
        }),
        editable: true,
        update: !!update,
        activate: update?.activate,
      },
    ];
  }
}
