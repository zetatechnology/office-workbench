import * as vscode from 'vscode';
import { EditError, EncodingError, MAX_FILE_BYTES, type EditOp, type HostMessage, type SheetSession, type WebviewMessage } from '@office-workbench/core';
import { getWebviewHtml } from './html';

/** An open workbook or CSV file shared by all of its editor panels. */
export class SheetDocument implements vscode.CustomDocument {
  readonly panels = new Set<vscode.WebviewPanel>();
  /** Edits applied since the last save or revert (undo counts down), to know if it is clean. */
  unsavedEdits = 0;
  /** File change events up to this time come from our own save. */
  ignoreChangesUntil = 0;
  private readonly disposables: vscode.Disposable[] = [];

  constructor(
    readonly uri: vscode.Uri,
    readonly fileName: string,
    readonly session: SheetSession,
    onDiskChange: () => void,
  ) {
    const folder = vscode.Uri.joinPath(uri, '..');
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(folder, fileName));
    let timer: ReturnType<typeof setTimeout> | undefined;
    const schedule = () => {
      clearTimeout(timer);
      timer = setTimeout(onDiskChange, 300);
    };
    this.disposables.push(watcher, watcher.onDidChange(schedule), watcher.onDidCreate(schedule), {
      dispose: () => clearTimeout(timer),
    });
  }

  broadcast(messages: HostMessage[]): void {
    for (const panel of this.panels) for (const message of messages) void panel.webview.postMessage(message);
  }

  /** Registers cleanup to run when VS Code disposes the document. */
  onDispose(fn: () => void): void {
    this.disposables.push({ dispose: fn });
  }

  dispose(): void {
    for (const d of this.disposables) d.dispose();
  }
}

/** Open sheet documents by URI, shared by the providers and the language model tools. */
export class SheetDocuments {
  private readonly documents = new Map<string, { document: SheetDocument; provider: SheetEditorProvider }>();

  get(uri: vscode.Uri): { document: SheetDocument; provider: SheetEditorProvider } | undefined {
    return this.documents.get(uri.toString());
  }

  add(document: SheetDocument, provider: SheetEditorProvider): void {
    const key = document.uri.toString();
    this.documents.set(key, { document, provider });
    document.onDispose(() => {
      if (this.documents.get(key)?.document === document) this.documents.delete(key);
    });
  }
}

/** Editable custom editor for Excel workbooks and CSV/TSV files, both shown in the sheet grid. */
export class SheetEditorProvider implements vscode.CustomEditorProvider<SheetDocument> {
  private readonly changes = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<SheetDocument>>();
  readonly onDidChangeCustomDocument = this.changes.event;

  constructor(
    private readonly context: vscode.ExtensionContext,
    private readonly createSession: (fileName: string, data: Uint8Array) => SheetSession,
    private readonly registry: SheetDocuments,
  ) {}

  async openCustomDocument(uri: vscode.Uri, openContext: vscode.CustomDocumentOpenContext): Promise<SheetDocument> {
    // After a restart with unsaved changes, VS Code hands back the backup written by backupCustomDocument.
    const source = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : uri;
    const data = await readBytes(source);
    const fileName = uri.path.slice(uri.path.lastIndexOf('/') + 1);
    const document: SheetDocument = new SheetDocument(uri, fileName, this.createSession(fileName, data), () =>
      void this.onDiskChange(document),
    );
    if (openContext.backupId) document.unsavedEdits = 1;
    this.registry.add(document, this);
    return document;
  }

  resolveCustomEditor(document: SheetDocument, panel: vscode.WebviewPanel): void {
    const webviewRoot = vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview');
    panel.webview.options = { enableScripts: true, localResourceRoots: [webviewRoot] };
    panel.webview.html = getWebviewHtml(panel.webview, webviewRoot, 'xlsx');
    document.panels.add(panel);

    const post = (messages: HostMessage[]) => {
      for (const message of messages) void panel.webview.postMessage(message);
    };
    const subscription = panel.webview.onDidReceiveMessage(async (message: WebviewMessage) => {
      if (message.type !== 'edit') return post(await document.session.handle(message));
      try {
        await this.applyEdit(document, message.op);
      } catch (e) {
        if (!(e instanceof EditError)) console.error(e);
        post([{ type: 'notice', message: e instanceof Error ? e.message : String(e) }]);
      }
    });
    panel.onDidDispose(() => {
      document.panels.delete(panel);
      subscription.dispose();
    });
  }

  /**
   * Applies an edit as one undo step and updates every view of the document. Throws EditError when
   * the edit is refused.
   */
  async applyEdit(document: SheetDocument, op: EditOp): Promise<void> {
    const edit = await document.session.edit(op);
    document.unsavedEdits++;
    document.broadcast(edit.messages);
    this.changes.fire({
      document,
      label: edit.label,
      undo: async () => {
        document.unsavedEdits--;
        document.broadcast(await edit.undo());
      },
      redo: async () => {
        document.unsavedEdits++;
        document.broadcast(await edit.redo());
      },
    });
  }

  saveCustomDocument(document: SheetDocument): Promise<void> {
    return this.write(document, document.uri);
  }

  async saveCustomDocumentAs(document: SheetDocument, destination: vscode.Uri): Promise<void> {
    await this.write(document, destination);
  }

  async revertCustomDocument(document: SheetDocument): Promise<void> {
    await this.reload(document);
  }

  async backupCustomDocument(document: SheetDocument, context: vscode.CustomDocumentBackupContext): Promise<vscode.CustomDocumentBackup> {
    await vscode.workspace.fs.writeFile(context.destination, await document.session.save());
    return {
      id: context.destination.toString(),
      delete: () => void vscode.workspace.fs.delete(context.destination).then(undefined, () => undefined),
    };
  }

  private async write(document: SheetDocument, target: vscode.Uri): Promise<void> {
    let data: Uint8Array;
    try {
      data = await document.session.save();
    } catch (e) {
      // e.g. a character the CSV file's encoding cannot represent; VS Code reports the failed save.
      if (e instanceof EncodingError) throw new Error(`${e.message} Remove that character to save the file.`);
      throw e;
    }
    if (target.toString() === document.uri.toString()) {
      document.ignoreChangesUntil = Date.now() + 2000;
      document.unsavedEdits = 0;
    }
    await vscode.workspace.fs.writeFile(target, data);
  }

  private async reload(document: SheetDocument): Promise<void> {
    document.session.reset(await readBytes(document.uri));
    document.unsavedEdits = 0;
    document.broadcast(await document.session.handle({ type: 'ready' }));
  }

  /** Picks up changes made by other programs (Excel, git) unless there are unsaved edits here. */
  private async onDiskChange(document: SheetDocument): Promise<void> {
    if (Date.now() < document.ignoreChangesUntil || document.unsavedEdits !== 0) return;
    try {
      await this.reload(document);
    } catch {
      // The file may be mid-write or deleted; the next change event will retry.
    }
  }
}

async function readBytes(uri: vscode.Uri): Promise<Uint8Array> {
  // readFile returns a Node Buffer; postMessage and JSZip expect a plain Uint8Array.
  const data = new Uint8Array(await vscode.workspace.fs.readFile(uri));
  if (data.byteLength > MAX_FILE_BYTES) {
    throw new Error(`The file is too large to open (${Math.round(data.byteLength / 1024 / 1024)} MB).`);
  }
  return data;
}
