import * as vscode from 'vscode';
import { statSync } from 'node:fs';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { buildPrintHtml, findBrowser, PAGE_SIZES, printToPdf, type PageSize } from './markdown-pdf';

const PRINT_TIMEOUT_MS = 120_000;

/**
 * "Export Markdown to PDF": renders the Markdown with VS Code's built-in Markdown engine (the same
 * markdown-it setup as the preview) and prints it with a locally installed Chromium-based browser.
 */
export function registerMarkdownPdfExport(context: vscode.ExtensionContext): void {
  context.subscriptions.push(
    vscode.commands.registerCommand('officeWorkbench.exportMarkdownToPdf', async (arg?: unknown) => {
      try {
        await exportToPdf(arg);
      } catch (e) {
        void vscode.window.showErrorMessage(`Could not export to PDF: ${e instanceof Error ? e.message : String(e)}`);
      }
    }),
  );
}

async function exportToPdf(arg: unknown): Promise<void> {
  const document = await targetDocument(arg);
  if (!document) {
    void vscode.window.showInformationMessage('Open or select a Markdown file first.');
    return;
  }

  const config = vscode.workspace.getConfiguration('officeWorkbench.markdownPdf', document.uri);
  const browser = resolveBrowser(config.get<string>('browserPath', ''));
  if (!browser) {
    const action = await vscode.window.showErrorMessage(
      'Exporting to PDF needs Google Chrome, Microsoft Edge or Chromium. Install one, or set its path in "officeWorkbench.markdownPdf.browserPath".',
      'Open Settings',
    );
    if (action) void vscode.commands.executeCommand('workbench.action.openSettings', 'officeWorkbench.markdownPdf.browserPath');
    return;
  }

  const source = document.uri.scheme === 'file' ? document.uri : undefined;
  const target = await vscode.window.showSaveDialog({
    defaultUri: defaultTarget(document),
    filters: { PDF: ['pdf'] },
    saveLabel: 'Export',
  });
  if (!target) return;

  let body: string;
  try {
    body = await vscode.commands.executeCommand<string>('markdown.api.render', document);
  } catch {
    throw new Error('the built-in Markdown extension is not available (is "Markdown Language Features" disabled?).');
  }
  const baseDir = source ? path.dirname(source.fsPath) : vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
  const pageSize = config.get<string>('pageSize', 'A4');
  const html = buildPrintHtml({
    title: titleOf(document),
    body: body ?? '',
    baseHref: baseDir ? withTrailingSlash(pathToFileURL(baseDir).href) : undefined,
    lang: vscode.env.language,
    pageSize: (PAGE_SIZES as readonly string[]).includes(pageSize) ? (pageSize as PageSize) : 'A4',
  });

  const written = await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: `Exporting ${titleOf(document)} to PDF…`, cancellable: true },
    async (_progress, token) => {
      const work = await mkdtemp(path.join(os.tmpdir(), 'office-workbench-pdf-'));
      try {
        const htmlFile = path.join(work, 'document.html');
        await writeFile(htmlFile, html, 'utf8');
        const abort = new AbortController();
        const subscription = token.onCancellationRequested(() => abort.abort());
        let pdf: Uint8Array;
        try {
          pdf = await printToPdf(browser, path.join(work, 'profile'), pathToFileURL(htmlFile).href, {
            timeoutMs: PRINT_TIMEOUT_MS,
            signal: abort.signal,
          });
        } catch (e) {
          if (token.isCancellationRequested) return false;
          throw e;
        } finally {
          subscription.dispose();
        }
        await vscode.workspace.fs.writeFile(target, pdf);
        return true;
      } finally {
        // The browser may still hold the profile for a moment after exiting (Windows).
        await rm(work, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }).catch(() => undefined);
      }
    },
  );

  if (!written) return; // Cancelled.
  const open = await vscode.window.showInformationMessage(`Exported ${path.basename(target.fsPath)}.`, 'Open');
  if (open) await vscode.commands.executeCommand('vscode.open', target);
}

/** The Markdown document a command applies to: its argument (explorer, editor title) or the active editor. */
async function targetDocument(arg: unknown): Promise<vscode.TextDocument | undefined> {
  const uri = arg instanceof vscode.Uri ? arg : undefined;
  // Prefer the open document, so unsaved changes are exported too.
  const doc = uri
    ? (vscode.workspace.textDocuments.find((d) => d.uri.toString() === uri.toString()) ?? (await vscode.workspace.openTextDocument(uri)))
    : vscode.window.activeTextEditor?.document;
  if (!doc) return undefined;
  if (doc.languageId === 'markdown' || /\.(md|markdown|mdown|mkd)$/i.test(doc.uri.path)) return doc;
  return undefined;
}

function resolveBrowser(configured: string): string | undefined {
  const isFile = (file: string) => {
    try {
      return statSync(file).isFile();
    } catch {
      return false;
    }
  };
  const custom = configured.trim();
  if (custom) {
    if (!isFile(custom)) throw new Error(`the browser set in "officeWorkbench.markdownPdf.browserPath" was not found: ${custom}`);
    return custom;
  }
  return findBrowser(process.platform, process.env, isFile);
}

function defaultTarget(document: vscode.TextDocument): vscode.Uri | undefined {
  if (document.uri.scheme === 'file') return vscode.Uri.file(document.uri.fsPath.replace(/\.[^./\\]*$/, '') + '.pdf');
  const folder = vscode.workspace.workspaceFolders?.[0]?.uri;
  return folder ? vscode.Uri.joinPath(folder, `${titleOf(document)}.pdf`) : undefined;
}

function titleOf(document: vscode.TextDocument): string {
  const name = document.uri.path.slice(document.uri.path.lastIndexOf('/') + 1);
  return name.replace(/\.[^.]*$/, '') || name;
}

function withTrailingSlash(url: string): string {
  return url.endsWith('/') ? url : `${url}/`;
}
