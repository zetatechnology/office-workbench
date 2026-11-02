// Browser harness for the viewer webviews: serves the built webview bundles with a mocked
// `acquireVsCodeApi` that forwards messages to the real core backend over HTTP.
// Workbooks, CSV files, documents and presentations are editable: Ctrl+Z / Ctrl+Y undo and redo,
// Ctrl+S writes the result to a temp folder (never over the source file).
// Usage: npm run build && npm run preview [-- --dir <folder with Office files>] [--port 5178]
import * as esbuild from 'esbuild';
import { createServer } from 'node:http';
import { readFile, readdir, mkdir, writeFile } from 'node:fs/promises';
import os from 'node:os';
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
const { ViewerBackend, WorkbookSession, CsvSession, WordSession, PresentationSession } = await import(pathToFileURL(coreBundle).href);
const saveDir = path.join(os.tmpdir(), 'office-workbench-preview');

const KINDS = { docx: 'docx', docm: 'docx', xlsx: 'xlsx', xlsm: 'xlsx', csv: 'xlsx', tsv: 'xlsx', tab: 'xlsx', pptx: 'pptx', pptm: 'pptx', pdf: 'pdf' };
const isDelimited = (name) => /\.(csv|tsv|tab)$/i.test(name);
const kindOf = (name) => KINDS[path.extname(name).slice(1).toLowerCase()];
const backends = new Map();
/** Editable files: file → { session, undo: [], redo: [] }. */
const sessions = new Map();
const SESSIONS = { xlsx: WorkbookSession, docx: WordSession, pptx: PresentationSession };

async function handleSession(file, message) {
  if (message.type === 'ready') {
    const data = new Uint8Array(await readFile(path.join(fileDir, file)));
    const Session = isDelimited(file) ? CsvSession : SESSIONS[kindOf(file)];
    sessions.set(file, { session: new Session(file, data), undo: [], redo: [] });
  }
  const state = sessions.get(file);
  switch (message.type) {
    case 'edit':
      try {
        const edit = await state.session.edit(message.op);
        state.undo.push(edit);
        state.redo = [];
        return edit.messages;
      } catch (e) {
        return [{ type: 'notice', message: e.message }];
      }
    case 'undo':
    case 'redo': {
      const edit = state[message.type].pop();
      if (!edit) return [];
      state[message.type === 'undo' ? 'redo' : 'undo'].push(edit);
      return edit[message.type]();
    }
    case 'save': {
      await mkdir(saveDir, { recursive: true });
      const target = path.join(saveDir, file);
      await writeFile(target, await state.session.save());
      console.log(`Saved ${target}`);
      return [{ type: 'notice', message: `Saved to ${target}` }];
    }
    default:
      return state.session.handle(message);
  }
}

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
    // Requests are queued so replies arrive in order, as with VS Code's postMessage.
    let queue = Promise.resolve();
    const api = {
      postMessage(message) {
        queue = queue
          .then(() => fetch('/api?file=' + encodeURIComponent(file), { method: 'POST', body: JSON.stringify(message) }))
          .then((r) => r.json())
          .then((list) => list.forEach((m) => window.dispatchEvent(new MessageEvent('message', { data: revive(m) }))))
          .catch((e) => console.error(e));
      },
      getState() { try { return JSON.parse(sessionStorage.getItem(file) ?? 'null'); } catch { return null; } },
      setState(state) { try { sessionStorage.setItem(file, JSON.stringify(state)); } catch {} },
    };
    window.acquireVsCodeApi = () => api;
    // ?raf=timeout: animation frames never fire while a browser window is occluded; use timers instead.
    if (new URLSearchParams(location.search).get('raf') === 'timeout') {
      window.requestAnimationFrame = (cb) => setTimeout(() => cb(performance.now()), 16);
      window.cancelAnimationFrame = (id) => clearTimeout(id);
    }
    // Stand-ins for VS Code's undo/redo/save keybindings (preview only). Like VS Code's, this
    // listens after the page's own handlers, which may keep undo keys for typing in progress.
    window.addEventListener('keydown', (e) => {
      const key = e.key.toLowerCase();
      if (!(e.ctrlKey || e.metaKey) || !['z', 'y', 's'].includes(key)) return;
      if (document.activeElement?.matches?.('.cell-editor.active, input')) return;
      e.preventDefault();
      api.postMessage({ type: key === 's' ? 'save' : key === 'y' || e.shiftKey ? 'redo' : 'undo' });
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
      if (kindOf(file) !== 'pdf') {
        res.writeHead(200, { 'content-type': 'application/json' });
        return res.end(JSON.stringify(encode(await handleSession(file, message))));
      }
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
