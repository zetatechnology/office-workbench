import * as vscode from 'vscode';
import type { ViewerKind } from '@office-workbench/core';
import { OfficeViewerProvider } from './provider';
import { WorkbookEditorProvider } from './workbook-editor';

/** Formats that are view-only so far; Excel workbooks have their own editable provider. */
const VIEW_ONLY: ViewerKind[] = ['docx', 'pptx', 'pdf'];

export function activate(context: vscode.ExtensionContext): void {
  // Webviews are rebuilt from the host on demand; avoid keeping large documents in hidden tabs.
  const options = { webviewOptions: { retainContextWhenHidden: false }, supportsMultipleEditorsPerDocument: true };
  for (const kind of VIEW_ONLY) {
    context.subscriptions.push(
      vscode.window.registerCustomEditorProvider(`officeWorkbench.${kind}`, new OfficeViewerProvider(context, kind), options),
    );
  }
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider('officeWorkbench.xlsx', new WorkbookEditorProvider(context), options),
  );
}

export function deactivate(): void {}
