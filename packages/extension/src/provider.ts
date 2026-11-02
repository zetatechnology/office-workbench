import * as vscode from 'vscode';
import { ViewerBackend, type HostMessage, type ViewerKind, type WebviewMessage } from '@office-workbench/core';
import { getWebviewHtml } from './html';

/** Read-only custom editor (PDF; the Office formats and CSV use OfficeEditorProvider). */
export class OfficeViewerProvider implements vscode.CustomReadonlyEditorProvider {
  constructor(
    private readonly context: vscode.ExtensionContext,
    private readonly kind: ViewerKind,
  ) {}

  openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
    return { uri, dispose: () => undefined };
  }

  resolveCustomEditor(document: vscode.CustomDocument, panel: vscode.WebviewPanel): void {
    const webviewRoot = vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview');
    panel.webview.options = { enableScripts: true, localResourceRoots: [webviewRoot] };
    panel.webview.html = getWebviewHtml(panel.webview, webviewRoot, this.kind);

    const uri = document.uri;
    const fileName = uri.path.slice(uri.path.lastIndexOf('/') + 1);
    // readFile returns a Node Buffer, which postMessage serializes as JSON ({type: 'Buffer', data: [...]})
    // rather than transferring it as binary; copy into a plain Uint8Array.
    const backend = new ViewerBackend(this.kind, fileName, async () => new Uint8Array(await vscode.workspace.fs.readFile(uri)));
    const disposables: vscode.Disposable[] = [];

    const post = (messages: HostMessage[]) => {
      for (const message of messages) void panel.webview.postMessage(message);
    };
    const handle = async (message: WebviewMessage) => post(await backend.handle(message));

    disposables.push(panel.webview.onDidReceiveMessage((message: WebviewMessage) => void handle(message)));

    // Reload when the file changes on disk (e.g. saved from Office or checked out by git).
    let reloadTimer: ReturnType<typeof setTimeout> | undefined;
    const scheduleReload = () => {
      clearTimeout(reloadTimer);
      reloadTimer = setTimeout(() => {
        backend.invalidate();
        void handle({ type: 'ready' });
      }, 300);
    };
    const folder = vscode.Uri.joinPath(uri, '..');
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(folder, fileName));
    disposables.push(watcher, watcher.onDidChange(scheduleReload), watcher.onDidCreate(scheduleReload));

    panel.onDidDispose(() => {
      clearTimeout(reloadTimer);
      for (const d of disposables) d.dispose();
    });
  }
}
