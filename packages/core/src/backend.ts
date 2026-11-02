// Host-side request handling for a single open file, independent of the VS Code API so the
// same logic serves the extension and the browser preview harness.
import { parsePresentation } from './pptx/presentation';
import type { HostMessage, ViewerKind, WebviewMessage } from './protocol';
import { Workbook } from './xlsx/workbook';

/** Files larger than this are refused rather than risking the extension host's memory. */
export const MAX_FILE_BYTES = 200 * 1024 * 1024;

export class ViewerBackend {
  private workbook?: Promise<Workbook>;

  constructor(
    private readonly kind: ViewerKind,
    private readonly fileName: string,
    private readonly readFile: () => Promise<Uint8Array>,
  ) {}

  /** Drops cached state so the next request re-reads the file (e.g. after it changed on disk). */
  invalidate(): void {
    this.workbook = undefined;
  }

  async handle(message: WebviewMessage): Promise<HostMessage[]> {
    try {
      switch (message.type) {
        case 'ready':
          this.invalidate();
          return [await this.load()];
        case 'loadSheet': {
          const workbook = await this.openWorkbook();
          return [{ type: 'sheet', sheet: await workbook.sheet(message.index) }];
        }
      }
    } catch (e) {
      return [{ type: 'error', message: e instanceof Error ? e.message : String(e) }];
    }
  }

  private async load(): Promise<HostMessage> {
    switch (this.kind) {
      case 'docx':
        return { type: 'docx', fileName: this.fileName, data: await this.read() };
      case 'pdf':
        return { type: 'pdf', fileName: this.fileName, data: await this.read() };
      case 'xlsx':
        return { type: 'workbook', fileName: this.fileName, info: (await this.openWorkbook()).info() };
      case 'pptx': {
        const { presentation, media } = await parsePresentation(await this.read());
        return {
          type: 'presentation',
          fileName: this.fileName,
          presentation,
          media: presentation.media.flatMap((ref) => {
            const data = media.get(ref.path);
            return data ? [{ ...ref, data }] : [];
          }),
        };
      }
    }
  }

  private openWorkbook(): Promise<Workbook> {
    if (!this.workbook) {
      this.workbook = this.read().then((data) => Workbook.open(data));
      this.workbook.catch(() => (this.workbook = undefined));
    }
    return this.workbook;
  }

  private async read(): Promise<Uint8Array> {
    const data = await this.readFile();
    if (data.byteLength > MAX_FILE_BYTES) {
      throw new Error(`The file is too large to preview (${Math.round(data.byteLength / 1024 / 1024)} MB).`);
    }
    return data;
  }
}
