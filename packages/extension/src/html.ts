import * as vscode from 'vscode';
import type { ViewerKind } from '@office-workbench/core';

export function getWebviewHtml(webview: vscode.Webview, root: vscode.Uri, kind: ViewerKind): string {
  const nonce = createNonce();
  const script = webview.asWebviewUri(vscode.Uri.joinPath(root, `${kind}.js`));
  const style = webview.asWebviewUri(vscode.Uri.joinPath(root, `${kind}.css`));
  const assets = webview.asWebviewUri(root);
  const src = webview.cspSource;
  // Inline styles are required by the renderers (docx-preview injects <style> blocks); scripts stay nonce-only.
  const csp = [
    "default-src 'none'",
    `img-src ${src} blob: data:`,
    `font-src ${src} blob: data:`,
    `style-src ${src} 'unsafe-inline'`,
    `script-src 'nonce-${nonce}'`,
  ];
  if (kind === 'pdf') {
    // pdf.js runs its worker from a blob (cross-origin worker scripts are blocked in webviews),
    // fetches CMaps / fonts / wasm decoders from the extension, and compiles those decoders.
    csp[csp.length - 1] += " 'wasm-unsafe-eval'";
    csp.push('worker-src blob:', `connect-src ${src}`);
  }
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="${csp.join('; ')}">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="stylesheet" href="${style}">
</head>
<body data-assets="${assets}">
  <div id="app"></div>
  <script nonce="${nonce}" src="${script}"></script>
</body>
</html>`;
}

function createNonce(): string {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let out = '';
  for (let i = 0; i < 32; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}
