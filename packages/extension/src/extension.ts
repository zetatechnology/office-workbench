import * as vscode from 'vscode';
import type { ViewerKind } from '@office-workbench/core';
import { OfficeViewerProvider } from './provider';

const KINDS: ViewerKind[] = ['docx', 'xlsx', 'pptx', 'pdf'];

export function activate(context: vscode.ExtensionContext): void {
  for (const kind of KINDS) {
    context.subscriptions.push(
      vscode.window.registerCustomEditorProvider(`officeWorkbench.${kind}`, new OfficeViewerProvider(context, kind), {
        // Webviews are rebuilt from the host on demand; avoid keeping large documents in hidden tabs.
        webviewOptions: { retainContextWhenHidden: false },
        supportsMultipleEditorsPerDocument: true,
      }),
    );
  }
}

export function deactivate(): void {}
