// Helpers for exporting Markdown to PDF: the printable HTML page and the headless browser that prints it.
// Kept free of the vscode API so they can be unit-tested.
import { spawn } from 'node:child_process';
import path from 'node:path';
import type { Readable, Writable } from 'node:stream';

export const PAGE_SIZES = ['A4', 'A3', 'A5', 'Letter', 'Legal'] as const;
export type PageSize = (typeof PAGE_SIZES)[number];

export interface PrintHtmlOptions {
  /** Document title (PDF metadata). */
  title: string;
  /** Rendered Markdown body (HTML). */
  body: string;
  /** file: URL of the Markdown file's folder (with a trailing slash), so relative images resolve. */
  baseHref?: string;
  /** BCP 47 language, which picks the CJK glyph variants. */
  lang: string;
  pageSize: PageSize;
}

export function escapeHtml(text: string): string {
  return text.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
}

/**
 * A standalone HTML page for printing. Scripts are blocked by the CSP: raw HTML in the Markdown runs
 * in a real browser here, unlike the sandboxed preview. Images may come from local files, data: or https.
 */
export function buildPrintHtml(options: PrintHtmlOptions): string {
  const size = PAGE_SIZES.includes(options.pageSize) ? options.pageSize : 'A4';
  const csp = "default-src 'none'; img-src file: data: https:; style-src 'unsafe-inline'; font-src file: data:";
  const base = options.baseHref ? `<base href="${escapeHtml(options.baseHref)}">\n` : '';
  return `<!DOCTYPE html>
<html lang="${escapeHtml(options.lang)}">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
${base}<title>${escapeHtml(options.title)}</title>
<style>
@page { size: ${size}; margin: 18mm 16mm; }
${PRINT_CSS}
</style>
</head>
<body>
${neutralizeMetaTags(options.body)}
</body>
</html>
`;
}

/**
 * Turns the start of every <meta> tag in the rendered Markdown into text: a `<meta http-equiv="refresh">`
 * would navigate away from the page (and its CSP) while it is being printed. Only the `<meta` is
 * replaced, so tags with unbalanced quotes are neutralized too. Code blocks are already escaped.
 */
export function neutralizeMetaTags(html: string): string {
  return html.replace(/<(meta[\s/>])/gi, '&lt;$1');
}

const PRINT_CSS = `
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body {
  margin: 0; color: #1f2328; background: #fff; font-size: 10.5pt; line-height: 1.7; overflow-wrap: break-word;
  /* BIZ UDPGothic before Yu Gothic / Meiryo: with those, Chrome on Windows records some kanji (日 見 文) as
     Kangxi radicals (U+2F00 block) in the PDF text, which breaks searching and copying. */
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Hiragino Sans", "Hiragino Kaku Gothic ProN",
    "BIZ UDPGothic", "Yu Gothic UI", "Yu Gothic", Meiryo, "Noto Sans CJK JP", "Noto Sans JP", "Helvetica Neue", Arial, sans-serif;
}
h1, h2, h3, h4, h5, h6 { margin: 1.4em 0 0.6em; line-height: 1.3; font-weight: 600; break-after: avoid; }
h1 { font-size: 1.9em; padding-bottom: 0.25em; border-bottom: 1px solid #d1d9e0; }
h2 { font-size: 1.5em; padding-bottom: 0.25em; border-bottom: 1px solid #d1d9e0; }
h3 { font-size: 1.25em; }
h4 { font-size: 1em; }
h5 { font-size: 0.9em; }
h6 { font-size: 0.85em; color: #59636e; }
body > :first-child { margin-top: 0; }
p, ul, ol, dl, table, pre, blockquote { margin: 0 0 0.9em; }
ul, ol { padding-left: 2em; }
li + li { margin-top: 0.2em; }
li > p { margin: 0.4em 0; }
a { color: #0969da; text-decoration: none; }
img { max-width: 100%; break-inside: avoid; }
hr { height: 0; margin: 1.5em 0; border: 0; border-top: 1px solid #d1d9e0; }
blockquote { padding: 0 1em; color: #59636e; border-left: 0.25em solid #d1d9e0; }
code, kbd, pre {
  font-family: "SFMono-Regular", Menlo, Consolas, "Liberation Mono", "BIZ UDGothic", "MS Gothic", "Noto Sans Mono CJK JP", monospace;
  font-size: 0.9em;
}
:not(pre) > code { padding: 0.15em 0.35em; background: #eff1f3; border-radius: 4px; }
pre { padding: 0.8em 1em; background: #f6f8fa; border: 1px solid #e1e4e8; border-radius: 6px; white-space: pre-wrap; line-height: 1.45; }
pre code { padding: 0; background: none; font-size: 1em; }
table { border-collapse: collapse; max-width: 100%; }
th, td { padding: 0.35em 0.8em; border: 1px solid #d1d9e0; }
th { font-weight: 600; background: #f6f8fa; }
tr { break-inside: avoid; }
thead { display: table-header-group; }
.hljs-comment, .hljs-quote { color: #6e7781; font-style: italic; }
.hljs-keyword, .hljs-selector-tag, .hljs-type, .hljs-doctag { color: #cf222e; }
.hljs-string, .hljs-regexp, .hljs-addition, .hljs-attribute { color: #0a3069; }
.hljs-number, .hljs-literal, .hljs-variable, .hljs-template-variable, .hljs-attr { color: #0550ae; }
.hljs-title, .hljs-section, .hljs-name, .hljs-selector-id, .hljs-selector-class { color: #8250df; }
.hljs-built_in, .hljs-symbol, .hljs-meta { color: #953800; }
.hljs-deletion { color: #82071e; background: #ffebe9; }
.hljs-emphasis { font-style: italic; }
.hljs-strong { font-weight: 600; }
`;

/**
 * Where to look for a Chromium-based browser (Chrome, Edge, Chromium, Brave), in order of preference.
 * Absolute paths are checked as they are; bare names are looked up on PATH.
 */
export function browserCandidates(platform: NodeJS.Platform, env: NodeJS.ProcessEnv): string[] {
  if (platform === 'win32') {
    const roots = [env['PROGRAMFILES'], env['PROGRAMFILES(X86)'], env['LOCALAPPDATA']].filter((r): r is string => !!r);
    const apps = [
      'Google\\Chrome\\Application\\chrome.exe',
      'Microsoft\\Edge\\Application\\msedge.exe',
      'Chromium\\Application\\chrome.exe',
      'BraveSoftware\\Brave-Browser\\Application\\brave.exe',
    ];
    return apps.flatMap((app) => roots.map((root) => path.win32.join(root, app)));
  }
  if (platform === 'darwin') {
    const apps = [
      'Google Chrome.app/Contents/MacOS/Google Chrome',
      'Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
      'Chromium.app/Contents/MacOS/Chromium',
      'Brave Browser.app/Contents/MacOS/Brave Browser',
    ];
    const roots = ['/Applications', ...(env['HOME'] ? [path.posix.join(env['HOME'], 'Applications')] : [])];
    return apps.flatMap((app) => roots.map((root) => path.posix.join(root, app)));
  }
  return [
    'google-chrome',
    'google-chrome-stable',
    'chromium',
    'chromium-browser',
    'microsoft-edge',
    'microsoft-edge-stable',
    'brave-browser',
  ];
}

/** The first candidate that exists, looking bare names up on PATH. */
export function findBrowser(
  platform: NodeJS.Platform,
  env: NodeJS.ProcessEnv,
  isFile: (file: string) => boolean,
): string | undefined {
  const p = platform === 'win32' ? path.win32 : path.posix;
  const dirs = (env['PATH'] ?? '').split(platform === 'win32' ? ';' : ':').filter(Boolean);
  for (const candidate of browserCandidates(platform, env)) {
    if (p.isAbsolute(candidate)) {
      if (isFile(candidate)) return candidate;
      continue;
    }
    for (const dir of dirs) {
      const file = p.join(dir, candidate);
      if (isFile(file)) return file;
    }
  }
  return undefined;
}

/** Command-line arguments for a headless browser controlled over the DevTools pipe (file descriptors 3 and 4). */
export function browserArgs(profileDir: string): string[] {
  return [
    '--headless=new',
    '--remote-debugging-pipe',
    // A profile of its own, so it runs alongside an open browser window instead of handing off to it.
    `--user-data-dir=${profileDir}`,
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-extensions',
    '--disable-gpu',
    '--disable-background-networking',
    '--disable-component-update',
    '--disable-sync',
    // Edge on Windows otherwise relaunches itself at startup, and the process holding the pipe exits.
    '--edge-skip-compat-layer-relaunch',
    'about:blank',
  ];
}

export interface PrintOptions {
  /** Extra browser arguments (tests). */
  extraArgs?: string[];
  timeoutMs?: number;
  signal?: AbortSignal;
}

/**
 * Prints `url` to PDF with a Chromium-based browser through the DevTools protocol, which behaves the
 * same across Chrome and Edge versions (the --print-to-pdf switch does not in every headless mode).
 */
export async function printToPdf(browser: string, profileDir: string, url: string, options: PrintOptions = {}): Promise<Uint8Array> {
  const child = spawn(browser, [...browserArgs(profileDir), ...(options.extraArgs ?? [])], {
    stdio: ['ignore', 'ignore', 'pipe', 'pipe', 'pipe'],
    windowsHide: true,
  });
  const cdp = new PipeConnection(child.stdio[3] as Writable, child.stdio[4] as Readable);
  let stderr = '';
  child.stderr?.on('data', (chunk: Buffer) => (stderr = (stderr + chunk.toString()).slice(-4000)));
  const name = path.basename(browser);
  const exited = new Promise<never>((_, reject) => {
    child.once('error', (e) => reject(new Error(`could not start ${browser}: ${e.message}`)));
    child.once('exit', (code) => {
      const detail = lastLine(stderr);
      reject(
        new Error(
          `${name} closed before the PDF was printed (exit code ${code}).` +
            (detail ? ` ${detail}` : '') +
            ' Try another browser in "officeWorkbench.markdownPdf.browserPath".',
        ),
      );
    });
  });
  exited.catch(() => undefined);
  let timer: ReturnType<typeof setTimeout> | undefined;
  let onAbort: (() => void) | undefined;
  const stop = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error('the browser timed out while printing.')), options.timeoutMs ?? 120_000);
    onAbort = () => reject(new Error('cancelled'));
    options.signal?.addEventListener('abort', onAbort, { once: true });
  });
  stop.catch(() => undefined);

  try {
    return await Promise.race([exited, stop, (async () => {
      const { targetId } = await cdp.send<{ targetId: string }>('Target.createTarget', { url: 'about:blank' });
      const { sessionId } = await cdp.send<{ sessionId: string }>('Target.attachToTarget', { targetId, flatten: true });
      await cdp.send('Page.enable', {}, sessionId);
      // Besides the page's CSP: also covers any page the document manages to navigate to.
      await cdp.send('Emulation.setScriptExecutionDisabled', { value: true }, sessionId);
      // Let only the first load of the page through; any later navigation of a frame is refused.
      let pageRequested = false;
      cdp.on('Fetch.requestPaused', sessionId, (params) => {
        const { requestId, request } = params as { requestId: string; request: { url: string } };
        const allow = !pageRequested && request.url === url;
        pageRequested ||= allow;
        const reply = allow
          ? cdp.send('Fetch.continueRequest', { requestId }, sessionId)
          : cdp.send('Fetch.failRequest', { requestId, errorReason: 'BlockedByClient' }, sessionId);
        reply.catch(() => undefined);
      });
      await cdp.send('Fetch.enable', { patterns: [{ resourceType: 'Document', requestStage: 'Request' }] }, sessionId);
      const loaded = cdp.waitFor('Page.loadEventFired', sessionId);
      const nav = await cdp.send<{ errorText?: string }>('Page.navigate', { url }, sessionId);
      if (nav.errorText) throw new Error(`the browser could not open the page: ${nav.errorText}`);
      await loaded;
      const { data } = await cdp.send<{ data: string }>(
        'Page.printToPDF',
        { printBackground: true, preferCSSPageSize: true, displayHeaderFooter: false, generateDocumentOutline: true },
        sessionId,
      );
      return new Uint8Array(Buffer.from(data, 'base64'));
    })()]);
  } finally {
    clearTimeout(timer);
    if (onAbort) options.signal?.removeEventListener('abort', onAbort);
    let grace: ReturnType<typeof setTimeout> | undefined;
    await Promise.race([cdp.send('Browser.close').catch(() => undefined), new Promise((r) => (grace = setTimeout(r, 2000)))]);
    clearTimeout(grace);
    if (child.exitCode === null) child.kill();
  }
}

function lastLine(text: string): string {
  return text.trim().split(/\r?\n/).at(-1) ?? '';
}

/** A DevTools protocol connection over --remote-debugging-pipe: NUL-terminated JSON messages. */
class PipeConnection {
  private nextId = 1;
  private buffer = '';
  private readonly pending = new Map<number, { resolve(value: unknown): void; reject(error: Error): void }>();
  private readonly listeners = new Set<(message: CdpMessage) => void>();

  constructor(
    private readonly output: Writable,
    input: Readable,
  ) {
    input.setEncoding('utf8');
    input.on('data', (chunk: string) => {
      this.buffer += chunk;
      let end: number;
      while ((end = this.buffer.indexOf('\0')) >= 0) {
        const raw = this.buffer.slice(0, end);
        this.buffer = this.buffer.slice(end + 1);
        this.dispatch(JSON.parse(raw) as CdpMessage);
      }
    });
    output.on('error', () => undefined);
  }

  send<T = unknown>(method: string, params: object = {}, sessionId?: string): Promise<T> {
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as (value: unknown) => void, reject });
      this.output.write(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }) + '\0');
    });
  }

  /** Calls `handler` with the parameters of every `method` event of the session. */
  on(method: string, sessionId: string, handler: (params: unknown) => void): void {
    this.listeners.add((message) => {
      if (message.method === method && message.sessionId === sessionId) handler(message.params);
    });
  }

  waitFor(method: string, sessionId: string): Promise<void> {
    return new Promise((resolve) => {
      const listener = (message: CdpMessage) => {
        if (message.method === method && message.sessionId === sessionId) {
          this.listeners.delete(listener);
          resolve();
        }
      };
      this.listeners.add(listener);
    });
  }

  private dispatch(message: CdpMessage): void {
    if (message.id !== undefined) {
      const call = this.pending.get(message.id);
      this.pending.delete(message.id);
      if (message.error) call?.reject(new Error(message.error.message));
      else call?.resolve(message.result);
      return;
    }
    for (const listener of [...this.listeners]) listener(message);
  }
}

interface CdpMessage {
  id?: number;
  method?: string;
  sessionId?: string;
  params?: unknown;
  result?: unknown;
  error?: { message: string };
}
