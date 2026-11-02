import * as vscode from 'vscode';
import { MAX_FILE_BYTES, textKindOf, toMarkdown } from '@office-workbench/core';

/**
 * Read-only Markdown views of Office files: `office-text:<path>.md?<source uri>`. The source may be
 * a file on disk or a `git:` revision, which is how the text diff against HEAD is built.
 */
const SCHEME = 'office-text';

export function registerTextDocuments(context: vscode.ExtensionContext): void {
  const changes = new vscode.EventEmitter<vscode.Uri>();
  const provider: vscode.TextDocumentContentProvider = {
    onDidChange: changes.event,
    async provideTextDocumentContent(uri) {
      const source = vscode.Uri.parse(uri.query);
      const kind = textKindOf(source.path);
      if (!kind) return `This file type cannot be shown as text: ${source.path}`;
      let data: Uint8Array;
      try {
        data = new Uint8Array(await vscode.workspace.fs.readFile(source));
      } catch {
        // Missing on one side of a diff (added, deleted or untracked file).
        return '';
      }
      if (data.byteLength > MAX_FILE_BYTES) return 'The file is too large to convert.';
      try {
        return await toMarkdown(kind, data);
      } catch (e) {
        return `This file could not be converted: ${e instanceof Error ? e.message : String(e)}`;
      }
    },
  };

  // Refresh open views when their source file changes (saved here, in Office, or by git).
  const watcher = vscode.workspace.createFileSystemWatcher('**/*.{docx,docm,xlsx,xlsm,pptx,pptm}');
  const refresh = (file: vscode.Uri) => {
    for (const doc of vscode.workspace.textDocuments) {
      if (doc.uri.scheme === SCHEME && vscode.Uri.parse(doc.uri.query).toString() === file.toString()) changes.fire(doc.uri);
    }
  };

  context.subscriptions.push(
    changes,
    watcher,
    watcher.onDidChange(refresh),
    watcher.onDidCreate(refresh),
    watcher.onDidDelete(refresh),
    vscode.workspace.registerTextDocumentContentProvider(SCHEME, provider),
    vscode.commands.registerCommand('officeWorkbench.openAsMarkdown', async (arg?: unknown) => {
      const uri = targetUri(arg);
      if (!uri) return;
      const doc = await vscode.workspace.openTextDocument(textUri(uri));
      await vscode.window.showTextDocument(doc, { preview: false });
    }),
    vscode.commands.registerCommand('officeWorkbench.compareWithHead', async (arg?: unknown) => {
      const uri = targetUri(arg);
      if (!uri) return;
      const git = await gitApi();
      if (!git?.getRepository(uri)) {
        void vscode.window.showWarningMessage('This file is not in a Git repository.');
        return;
      }
      const name = uri.path.slice(uri.path.lastIndexOf('/') + 1);
      await vscode.commands.executeCommand(
        'vscode.diff',
        textUri(git.toGitUri(uri, 'HEAD')),
        textUri(uri),
        `${name} (HEAD ↔ Working Tree, as text)`,
      );
    }),
  );
}

function textUri(source: vscode.Uri): vscode.Uri {
  return vscode.Uri.from({ scheme: SCHEME, path: `${source.path}.md`, query: source.toString() });
}

/** The file a command applies to: its argument (explorer, editor title, SCM) or the active editor. */
function targetUri(arg: unknown): vscode.Uri | undefined {
  if (arg instanceof vscode.Uri) return arg;
  const resource = (arg as { resourceUri?: unknown } | undefined)?.resourceUri;
  if (resource instanceof vscode.Uri) return resource;
  const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
  if (input instanceof vscode.TabInputCustom || input instanceof vscode.TabInputText) return input.uri;
  void vscode.window.showInformationMessage('Open or select a Word, Excel or PowerPoint file first.');
  return undefined;
}

/** The subset of the built-in Git extension's API used here. */
interface GitApi {
  getRepository(uri: vscode.Uri): unknown;
  toGitUri(uri: vscode.Uri, ref: string): vscode.Uri;
}

async function gitApi(): Promise<GitApi | undefined> {
  const extension = vscode.extensions.getExtension<{ getAPI(version: 1): GitApi }>('vscode.git');
  if (!extension) return undefined;
  const exports = extension.isActive ? extension.exports : await extension.activate();
  return exports.getAPI(1);
}
