// One open workbook as seen by its views: answers view requests and turns edits into undoable
// steps that report which messages to broadcast. Independent of the VS Code API.
import type { HostMessage, WebviewMessage } from '../protocol';
import { EditError, Workbook, type EditChange } from './workbook';
import type { EditOp } from './types';

export interface SessionEdit {
  label: string;
  /** Messages for every view of the workbook after the edit. */
  messages: HostMessage[];
  undo(): Promise<HostMessage[]>;
  redo(): Promise<HostMessage[]>;
}

/** One open spreadsheet-like file (workbook or CSV) behind the grid views. */
export interface SheetSession {
  /** Requests from a view; edits go through `edit` instead. */
  handle(message: WebviewMessage): Promise<HostMessage[]>;
  /** Applies an edit. Throws EditError (with a message for the user) when it is refused. */
  edit(op: EditOp): Promise<SessionEdit>;
  save(): Promise<Uint8Array>;
  /** Replaces the contents, e.g. after revert or a change on disk. */
  reset(data: Uint8Array): void;
}

export class WorkbookSession implements SheetSession {
  private workbook: Promise<Workbook>;
  /** Bumped by reset(); undo/redo of edits made before then do nothing. */
  private generation = 0;
  /** Operations run one at a time: edits await part loads, and must not interleave. */
  private queue: Promise<unknown> = Promise.resolve();

  constructor(
    private readonly fileName: string,
    data: Uint8Array,
    private readonly editable = true,
  ) {
    this.workbook = Workbook.open(data);
  }

  /** Replaces the contents, e.g. after revert or a change on disk. */
  reset(data: Uint8Array): void {
    this.generation++;
    this.workbook = Workbook.open(data);
  }

  /** Requests from a view; edits go through `edit` instead. */
  handle(message: WebviewMessage): Promise<HostMessage[]> {
    return this.run(() => this.respond(message));
  }

  private async respond(message: WebviewMessage): Promise<HostMessage[]> {
    try {
      const workbook = await this.workbook;
      switch (message.type) {
        case 'ready':
          return [this.workbookMessage(workbook)];
        case 'loadSheet':
          return [{ type: 'sheet', sheet: await workbook.sheet(message.index) }];
        case 'edit':
          return [{ type: 'notice', message: 'This workbook is read-only.' }];
      }
    } catch (e) {
      return [{ type: 'error', message: e instanceof Error ? e.message : String(e) }];
    }
  }

  /** Applies an edit. Throws EditError (with a message for the user) when it is refused. */
  edit(op: EditOp): Promise<SessionEdit> {
    return this.run(async () => {
      if (!this.editable) throw new EditError('This workbook is read-only.');
      const workbook = await this.workbook;
      const generation = this.generation;
      let applied = await workbook.apply(op);
      const step = (fn: () => Promise<HostMessage[]>) => () =>
        this.run(() => (generation === this.generation ? fn() : Promise.resolve([])));
      return {
        label: applied.label,
        messages: await this.messagesFor(workbook, applied.change),
        undo: step(async () => this.messagesFor(workbook, await applied.revert())),
        redo: step(async () => {
          applied = await workbook.apply(op);
          return this.messagesFor(workbook, applied.change);
        }),
      };
    });
  }

  save(): Promise<Uint8Array> {
    return this.run(async () => (await this.workbook).save());
  }

  private run<T>(fn: () => Promise<T>): Promise<T> {
    const result = this.queue.then(fn);
    this.queue = result.catch(() => undefined);
    return result;
  }

  private async messagesFor(workbook: Workbook, change: EditChange): Promise<HostMessage[]> {
    if (change.kind === 'cells') {
      return [{ type: 'cells', sheet: change.sheet, ...(await workbook.cells(change.sheet, change.cells)) }];
    }
    return [this.workbookMessage(workbook, change.activate)];
  }

  private workbookMessage(workbook: Workbook, activate?: number): HostMessage {
    return { type: 'workbook', fileName: this.fileName, info: workbook.info(), editable: this.editable, format: 'xlsx', activate };
  }
}
