// The shape shared by every open, editable document behind its views, and a base class for the
// documents (Word, PowerPoint) whose views are rebuilt from the whole model after each edit.
// Independent of the VS Code API.
import type { HostMessage, WebviewMessage } from './protocol';
import { EditError } from './errors';

export interface SessionEdit {
  label: string;
  /** Messages for every view of the document after the edit. */
  messages: HostMessage[];
  undo(): Promise<HostMessage[]>;
  redo(): Promise<HostMessage[]>;
}

/** One open file behind its editor views. */
export interface EditSession<Op> {
  /** Requests from a view; edits go through `edit` instead. */
  handle(message: WebviewMessage): Promise<HostMessage[]>;
  /** Applies an edit. Throws EditError (with a message for the user) when it is refused. */
  edit(op: Op): Promise<SessionEdit>;
  save(): Promise<Uint8Array>;
  /** Replaces the contents, e.g. after revert or a change on disk. */
  reset(data: Uint8Array): void;
}

/** An edit applied to a document model. */
export interface AppliedChange {
  label: string;
  /** Slide (or similar position) the views should show afterwards. */
  activate?: number;
  /** Restores the state before the edit; valid only while later edits have been reverted first. */
  revert(): Promise<{ activate?: number } | void>;
}

export interface EditableDocument<Op> {
  apply(op: Op): Promise<AppliedChange>;
  save(): Promise<Uint8Array>;
}

/**
 * Session over a document that applies operations itself. Operations run one at a time (edits
 * await part loads and must not interleave), and undo/redo of edits made before a reset do nothing.
 */
export abstract class DocumentSession<Doc extends EditableDocument<Op>, Op> implements EditSession<Op> {
  private document: Promise<Doc>;
  private generation = 0;
  private queue: Promise<unknown> = Promise.resolve();

  constructor(
    protected readonly fileName: string,
    data: Uint8Array,
    private readonly open: (data: Uint8Array) => Promise<Doc>,
  ) {
    this.document = open(data);
  }

  /** Every message a view needs to show the document; `update` is set after an edit. */
  protected abstract view(document: Doc, update: { activate?: number } | undefined): Promise<HostMessage[]>;

  reset(data: Uint8Array): void {
    this.generation++;
    this.document = this.open(data);
  }

  handle(message: WebviewMessage): Promise<HostMessage[]> {
    return this.run(async () => {
      try {
        const document = await this.document;
        if (message.type === 'ready') return this.view(document, undefined);
        return [];
      } catch (e) {
        return [{ type: 'error', message: e instanceof Error ? e.message : String(e) }];
      }
    });
  }

  edit(op: Op): Promise<SessionEdit> {
    return this.run(async () => {
      let document: Doc;
      try {
        document = await this.document;
      } catch (e) {
        throw new EditError(e instanceof Error ? e.message : String(e));
      }
      const generation = this.generation;
      let applied = await document.apply(op);
      const step = (fn: () => Promise<HostMessage[]>) => () =>
        this.run(() => (generation === this.generation ? fn() : Promise.resolve([])));
      return {
        label: applied.label,
        messages: await this.view(document, { activate: applied.activate }),
        undo: step(async () => this.view(document, (await applied.revert()) ?? {})),
        redo: step(async () => {
          applied = await document.apply(op);
          return this.view(document, { activate: applied.activate });
        }),
      };
    });
  }

  save(): Promise<Uint8Array> {
    return this.run(async () => (await this.document).save());
  }

  private run<T>(fn: () => Promise<T>): Promise<T> {
    const result = this.queue.then(fn);
    this.queue = result.catch(() => undefined);
    return result;
  }
}
