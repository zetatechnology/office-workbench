import { describe, expect, it } from 'vitest';
import { browserArgs, browserCandidates, buildPrintHtml, findBrowser, neutralizeMetaTags } from '../src/markdown-pdf';

describe('buildPrintHtml', () => {
  const base = { title: 'Notes', body: '<h1>見出し</h1>', lang: 'ja', pageSize: 'A4' as const };

  it('wraps the body in a page with the paper size, language and base URL', () => {
    const html = buildPrintHtml({ ...base, baseHref: 'file:///home/me/docs/' });
    expect(html).toContain('<html lang="ja">');
    expect(html).toContain('<base href="file:///home/me/docs/">');
    expect(html).toContain('@page { size: A4;');
    expect(html).toContain('<h1>見出し</h1>');
  });

  it('blocks scripts with a content security policy', () => {
    const html = buildPrintHtml(base);
    expect(html).toMatch(/Content-Security-Policy" content="default-src 'none';/);
    expect(html).not.toContain('script-src');
    expect(html).not.toContain('<base');
  });

  it('escapes the title and falls back to A4 for unknown sizes', () => {
    const html = buildPrintHtml({ ...base, title: '<a & "b">', pageSize: 'B9' as never });
    expect(html).toContain('<title>&#60;a &#38; &#34;b&#34;&#62;</title>');
    expect(html).toContain('@page { size: A4;');
  });
});

describe('neutralizeMetaTags', () => {
  it('turns the start of meta tags into text, whatever their quotes', () => {
    const body = '<p>a</p><META http-equiv="refresh" content="0;url=file:///x>y">\n<meta/><meta\tcharset=x>';
    expect(neutralizeMetaTags(body)).toBe('<p>a</p>&lt;META http-equiv="refresh" content="0;url=file:///x>y">\n&lt;meta/>&lt;meta\tcharset=x>');
  });

  it('handles a tag with an unclosed quote', () => {
    const body = "<div>\n<meta http-equiv=refresh content='0;url=file:///other.html' x='>\n</div><p>don't</p>";
    expect(neutralizeMetaTags(body)).not.toMatch(/<meta/i);
    expect(neutralizeMetaTags(body)).toContain('<p>don\'t</p>');
  });

  it('leaves other tags and escaped code alone, and is applied by buildPrintHtml', () => {
    expect(neutralizeMetaTags('<metadata><code>&lt;meta&gt;</code>')).toBe('<metadata><code>&lt;meta&gt;</code>');
    const html = buildPrintHtml({ title: 't', body: '<meta http-equiv="refresh" content="0">', lang: 'en', pageSize: 'A4' });
    expect(html).toContain('&lt;meta http-equiv="refresh"');
  });
});

describe('findBrowser', () => {
  it('checks the Windows install folders, Chrome first', () => {
    const env = { PROGRAMFILES: 'C:\\Program Files', 'PROGRAMFILES(X86)': 'C:\\Program Files (x86)', LOCALAPPDATA: 'C:\\Users\\me\\AppData\\Local' };
    const candidates = browserCandidates('win32', env);
    expect(candidates[0]).toBe('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe');
    const edge = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    expect(findBrowser('win32', env, (f) => f === edge)).toBe(edge);
  });

  it('checks the macOS application folders', () => {
    const edge = '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge';
    expect(findBrowser('darwin', { HOME: '/Users/me' }, (f) => f === edge)).toBe(edge);
  });

  it('looks bare names up on PATH on Linux', () => {
    const env = { PATH: '/usr/local/bin:/usr/bin' };
    expect(findBrowser('linux', env, (f) => f === '/usr/bin/chromium')).toBe('/usr/bin/chromium');
    expect(findBrowser('linux', env, () => false)).toBeUndefined();
  });
});

describe('browserArgs', () => {
  it('runs headless with the DevTools pipe and a private profile', () => {
    const args = browserArgs('/tmp/x/profile');
    expect(args).toContain('--headless=new');
    expect(args).toContain('--remote-debugging-pipe');
    expect(args).toContain('--user-data-dir=/tmp/x/profile');
    expect(args).toContain('--edge-skip-compat-layer-relaunch');
  });
});
