// Browser harness for the viewer webviews: serves the built webview bundles with a mocked
// `acquireVsCodeApi` that forwards messages to the real core backend over HTTP.
// Usage: npm run build && npm run preview [-- --dir <folder with Office files>] [--port 5178]
import * as esbuild from 'esbuild';
import { createServer } from 'node:http';
import { readFile, readdir, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const option = (name, fallback) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] : fallback;
};
const fileDir = path.resolve(option('dir', path.join(root, 'test-fixtures')));
const port = Number(option('port', process.env.PORT ?? 5178));
const webviewDir = path.join(root, 'packages/extension/dist/webview');

// Bundle the core package for Node so this script can use the TypeScript sources directly.
const cacheDir = path.join(root, 'node_modules/.cache/office-workbench');
await mkdir(cacheDir, { recursive: true });
const coreBundle = path.join(cacheDir, 'core.mjs');
await esbuild.build({
  entryPoints: [path.join(root, 'packages/core/src/index.ts')],
  outfile: coreBundle,
  bundle: true,
  platform: 'node',
  format: 'esm',
  logLevel: 'warning',
});
const { ViewerBackend } = await import(pathToFileURL(coreBundle).href);

const KINDS = { docx: 'docx', docm: 'docx', xlsx: 'xlsx', xlsm: 'xlsx', pptx: 'pptx', pptm: 'pptx', pdf: 'pdf' };
const kindOf = (name) => KINDS[path.extname(name).slice(1).toLowerCase()];
const backends = new Map();

const THEMES = {
  light: {
    '--vscode-editor-background': '#ffffff',
    '--vscode-foreground': '#3b3b3b',
    '--vscode-descriptionForeground': '#717171',
    '--vscode-editorGroupHeader-tabsBackground': '#f8f8f8',
    '--vscode-panel-border': '#e5e5e5',
    '--vscode-editorWidget-background': '#f0f0f0',
    '--vscode-focusBorder': '#005fb8',
    '--vscode-toolbar-hoverBackground': 'rgba(184, 184, 184, 0.31)',
    '--vscode-errorForeground': '#f85149',
    '--vscode-font-family': '"Segoe WPC", "Segoe UI", sans-serif',
    '--vscode-font-size': '13px',
  },
  dark: {
    '--vscode-editor-background': '#1f1f1f',
    '--vscode-foreground': '#cccccc',
    '--vscode-descriptionForeground': '#9d9d9d',
    '--vscode-editorGroupHeader-tabsBackground': '#181818',
    '--vscode-panel-border': '#2b2b2b',
    '--vscode-editorWidget-background': '#2a2a2a',
    '--vscode-focusBorder': '#0078d4',
    '--vscode-toolbar-hoverBackground': 'rgba(90, 93, 94, 0.31)',
    '--vscode-errorForeground': '#f85149',
    '--vscode-font-family': '"Segoe WPC", "Segoe UI", sans-serif',
    '--vscode-font-size': '13px',
  },
};

function viewerHtml(file, kind, theme) {
  const vars = Object.entries(THEMES[theme] ?? THEMES.light).map(([k, v]) => `${k}: ${v};`).join(' ');
  const nonce = 'preview';
  // Same policy as the extension (see packages/extension/src/html.ts), with 'self' standing in for cspSource.
  const pdf = kind === 'pdf' ? " 'wasm-unsafe-eval'; worker-src blob:" : '';
  const csp = `default-src 'none'; img-src 'self' blob: data:; font-src 'self' blob: data:; style-src 'self' 'unsafe-inline'; connect-src 'self'; script-src 'nonce-${nonce}'${pdf}`;
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="${csp}">
  <title>${file} – preview</title>
  <style>html { ${vars} }</style>
  <link rel="stylesheet" href="/webview/${kind}.css">
</head>
<body data-assets="/webview">
  <div id="app"></div>
  <script nonce="${nonce}">
    const file = ${JSON.stringify(file)};
    const revive = (value) => {
      if (value && typeof value === 'object') {
        if (typeof value.__bytes === 'string') return Uint8Array.from(atob(value.__bytes), (c) => c.charCodeAt(0));
        for (const key of Object.keys(value)) value[key] = revive(value[key]);
      }
      return value;
    };
    window.acquireVsCodeApi = () => ({
      postMessage(message) {
        fetch('/api?file=' + encodeURIComponent(file), { method: 'POST', body: JSON.stringify(message) })
          .then((r) => r.json())
          .then((list) => list.forEach((m) => window.dispatchEvent(new MessageEvent('message', { data: revive(m) }))));
      },
      getState() { try { return JSON.parse(sessionStorage.getItem(file) ?? 'null'); } catch { return null; } },
      setState(state) { try { sessionStorage.setItem(file, JSON.stringify(state)); } catch {} },
    });
  </script>
  <script nonce="${nonce}" src="/webview/${kind}.js"></script>
</body>
</html>`;
}

function encode(value) {
  if (value instanceof Uint8Array) return { __bytes: Buffer.from(value).toString('base64') };
  if (Array.isArray(value)) return value.map(encode);
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, encode(v)]));
  return value;
}

const server = createServer(async (req, res) => {
  try {
    const url = new URL(req.url ?? '/', `http://localhost:${port}`);
    if (url.pathname === '/') {
      const files = (await readdir(fileDir)).filter((f) => kindOf(f) && !f.startsWith('~$'));
      const items = files.map((f) => `<li><a href="/view?file=${encodeURIComponent(f)}">${f}</a> (<a href="/view?file=${encodeURIComponent(f)}&theme=dark">dark</a>)</li>`);
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
      return res.end(`<!DOCTYPE html><meta charset="utf-8"><title>Office Workbench preview</title><h1>Office Workbench preview</h1><p>${fileDir}</p><ul>${items.join('')}</ul>`);
    }
    if (url.pathname === '/view') {
      const file = path.basename(url.searchParams.get('file') ?? '');
      const kind = kindOf(file);
      if (!kind) throw new Error('Unsupported file type');
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
      return res.end(viewerHtml(file, kind, url.searchParams.get('theme') ?? 'light'));
    }
    if (url.pathname.startsWith('/webview/')) {
      const name = path.normalize(decodeURIComponent(url.pathname.slice('/webview/'.length)));
      if (name.startsWith('..') || path.isAbsolute(name)) throw new Error('Invalid path');
      const body = await readFile(path.join(webviewDir, name));
      const type = name.endsWith('.js') ? 'text/javascript' : name.endsWith('.css') ? 'text/css' : name.endsWith('.wasm') ? 'application/wasm' : 'application/octet-stream';
      res.writeHead(200, { 'content-type': type });
      return res.end(body);
    }
    if (url.pathname === '/api' && req.method === 'POST') {
      const file = path.basename(url.searchParams.get('file') ?? '');
      const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      const message = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      if (message.type === 'ready') backends.delete(file);
      let backend = backends.get(file);
      if (!backend) {
        backend = new ViewerBackend(kindOf(file), file, async () => new Uint8Array(await readFile(path.join(fileDir, file))));
        backends.set(file, backend);
      }
      const replies = await backend.handle(message);
      res.writeHead(200, { 'content-type': 'application/json' });
      return res.end(JSON.stringify(encode(replies)));
    }
    res.writeHead(404).end('Not found');
  } catch (e) {
    res.writeHead(500, { 'content-type': 'text/plain' }).end(String(e?.stack ?? e));
  }
});

server.listen(port, () => console.log(`Office Workbench preview: http://localhost:${port}/ (files from ${fileDir})`));
