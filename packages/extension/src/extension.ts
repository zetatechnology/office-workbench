import * as vscode from 'vscode';
import { CsvSession, PresentationSession, WordSession, WorkbookSession, type ViewerKind } from '@office-workbench/core';
import { OfficeEditorProvider, OfficeDocuments } from './editor';
import { OfficeViewerProvider } from './provider';
import { registerLanguageModelTools, type OfficeWorkbenchApi } from './lm-tools';
import { registerTextDocuments } from './text-documents';

/** Formats that are view-only; the Office formats and CSV files have editable editors. */
const VIEW_ONLY: ViewerKind[] = ['pdf'];
const CSV_VIEW_TYPE = 'officeWorkbench.csv';

export function activate(context: vscode.ExtensionContext): OfficeWorkbenchApi {
  registerTextDocuments(context);
  const documents = new OfficeDocuments();
  const api = registerLanguageModelTools(context, documents);
  // Webviews are rebuilt from the host on demand; avoid keeping large documents in hidden tabs.
  const options = { webviewOptions: { retainContextWhenHidden: false }, supportsMultipleEditorsPerDocument: true };
  for (const kind of VIEW_ONLY) {
    context.subscriptions.push(
      vscode.window.registerCustomEditorProvider(`officeWorkbench.${kind}`, new OfficeViewerProvider(context, kind), options),
    );
  }
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(
      'officeWorkbench.xlsx',
      new OfficeEditorProvider(context, 'xlsx', (fileName, data) => new WorkbookSession(fileName, data), documents),
      options,
    ),
    // CSV files are often edited as text, so the table view is an option rather than the default.
    vscode.window.registerCustomEditorProvider(
      CSV_VIEW_TYPE,
      new OfficeEditorProvider(context, 'xlsx', (fileName, data) => new CsvSession(fileName, data), documents),
      options,
    ),
    vscode.window.registerCustomEditorProvider(
      'officeWorkbench.docx',
      new OfficeEditorProvider(context, 'docx', (fileName, data) => new WordSession(fileName, data), documents),
      options,
    ),
    vscode.window.registerCustomEditorProvider(
      'officeWorkbench.pptx',
      new OfficeEditorProvider(context, 'pptx', (fileName, data) => new PresentationSession(fileName, data), documents),
      options,
    ),
    vscode.commands.registerCommand('officeWorkbench.openAsTable', (uri?: vscode.Uri) => {
      const target = uri ?? vscode.window.activeTextEditor?.document.uri;
      if (target) void vscode.commands.executeCommand('vscode.openWith', target, CSV_VIEW_TYPE);
    }),
    vscode.commands.registerCommand('officeWorkbench.openAsText', (uri?: vscode.Uri) => {
      const target = uri ?? activeCustomEditorUri();
      if (target) void vscode.commands.executeCommand('vscode.openWith', target, 'default');
    }),
  );
  return api;
}

function activeCustomEditorUri(): vscode.Uri | undefined {
  const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
  return input instanceof vscode.TabInputCustom ? input.uri : undefined;
}

export function deactivate(): void {}
