// In-place text editing on rendered content: the element becomes contenteditable, and its
// paragraphs are read back as plain text when editing ends. Enter starts a paragraph, Shift+Enter
// a line break, Esc cancels, Ctrl+Enter or clicking elsewhere commits. Typing, IME composition
// and native undo are left to the browser.

import { send } from './host';

/** Elements that start a paragraph when read back. */
const BLOCKS = new Set(['DIV', 'P', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6']);

export interface TextEditOptions {
  /** Called with the paragraphs when editing ends with changed text. */
  commit(paragraphs: string[]): void;
  /** Called when editing ends without a change (or is cancelled); the caller restores the view. */
  close(): void;
  /** Point to put the caret at (the double-click position). */
  caret?: { x: number; y: number };
}

export class TextEdit {
  readonly initial: string[];
  private done = false;
  private readonly cleanup: (() => void)[] = [];

  constructor(
    readonly host: HTMLElement,
    private readonly options: TextEditOptions,
  ) {
    for (const bullet of host.querySelectorAll<HTMLElement>('.bullet')) bullet.contentEditable = 'false';
    this.initial = readParagraphs(host);
    host.contentEditable = 'true';
    host.spellcheck = false;
    host.classList.add('editing');
    host.focus({ preventScroll: true });
    placeCaret(host, options.caret);

    const on = <K extends keyof HTMLElementEventMap>(type: K, fn: (e: HTMLElementEventMap[K]) => void) => {
      host.addEventListener(type, fn);
      this.cleanup.push(() => host.removeEventListener(type, fn));
    };
    on('keydown', (e) => this.onKey(e));
    on('blur', () => this.end(true));
    on('paste', (e) => {
      e.preventDefault();
      const lines = (e.clipboardData?.getData('text/plain') ?? '').replace(/\r\n?/g, '\n').split('\n');
      lines.forEach((line, i) => {
        if (i) document.execCommand('insertParagraph');
        if (line) document.execCommand('insertText', false, line);
      });
    });
    on('drop', (e) => e.preventDefault());
  }

  /** Whether the text differs from what editing started with. */
  get dirty(): boolean {
    return !this.done && !sameParagraphs(readParagraphs(this.host), this.initial);
  }

  get active(): boolean {
    return !this.done;
  }

  /** Ends editing; with `commit`, changed text is passed on. */
  end(commit: boolean): void {
    if (this.done) return;
    this.done = true;
    const paragraphs = readParagraphs(this.host);
    for (const fn of this.cleanup) fn();
    this.host.contentEditable = 'false';
    this.host.classList.remove('editing');
    window.getSelection()?.removeAllRanges();
    if (commit && !sameParagraphs(paragraphs, this.initial)) this.options.commit(paragraphs);
    else this.options.close();
  }

  private onKey(e: KeyboardEvent): void {
    if (e.isComposing || e.keyCode === 229) return;
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'Escape') {
      e.preventDefault();
      this.end(false);
    } else if (e.key === 'Enter' && mod) {
      e.preventDefault();
      this.end(true);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      document.execCommand('insertText', false, '\t');
    } else if (mod && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
      // Character formatting is not editable here; keep the browser from adding markup.
      e.preventDefault();
    }
  }
}

/**
 * Paragraphs of rendered or edited content: each block element is a paragraph, `<br>` a line
 * break ("\n"), except a trailing `<br>` that only keeps an empty line open.
 */
export function readParagraphs(root: HTMLElement): string[] {
  const paragraphs: string[] = [];
  let line: string | null = null;
  const end = () => {
    if (line !== null) paragraphs.push(line);
    line = null;
  };
  const visit = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      line = (line ?? '') + (node.textContent ?? '').replace(/​/g, '');
      return;
    }
    if (!(node instanceof HTMLElement) || node.classList.contains('bullet')) return;
    if (node.tagName === 'BR') {
      line = (line ?? '') + (endsBlock(node, root) ? '' : '\n');
      return;
    }
    const block = BLOCKS.has(node.tagName);
    if (block) {
      end();
      line = '';
    }
    for (const child of Array.from(node.childNodes)) visit(child);
    if (block) end();
  };
  for (const child of Array.from(root.childNodes)) visit(child);
  end();
  return paragraphs.length ? paragraphs : [''];
}

function endsBlock(br: HTMLElement, root: HTMLElement): boolean {
  for (let node: Node | null = br; node && node !== root; node = node.parentNode) {
    for (let next = node.nextSibling; next; next = next.nextSibling) {
      if (next.nodeType !== Node.TEXT_NODE || (next.textContent ?? '').replace(/​/g, '')) return false;
    }
    if (node.parentNode instanceof HTMLElement && BLOCKS.has(node.parentNode.tagName)) return true;
  }
  return true;
}

function sameParagraphs(a: string[], b: string[]): boolean {
  return a.length === b.length && a.every((p, i) => p === b[i]);
}

function placeCaret(host: HTMLElement, point: { x: number; y: number } | undefined): void {
  const selection = window.getSelection();
  if (!selection) return;
  let range: Range | null = null;
  if (point) {
    range = document.caretRangeFromPoint?.(point.x, point.y) ?? null;
    if (range && !host.contains(range.startContainer)) range = null;
  }
  if (!range) {
    range = document.createRange();
    range.selectNodeContents(host);
    range.collapse(false);
  }
  selection.removeAllRanges();
  selection.addRange(range);
}

/**
 * Keeps VS Code's undo/redo keys away from an editor with unsaved typing, so they undo the
 * typing (natively) rather than an earlier document edit. Without typing, they reach VS Code.
 */
export function guardUndoKeys(isTyping: () => boolean): void {
  window.addEventListener(
    'keydown',
    (e) => {
      const key = e.key.toLowerCase();
      if ((e.ctrlKey || e.metaKey) && (key === 'z' || key === 'y') && isTyping()) e.stopImmediatePropagation();
    },
    true,
  );
}

/**
 * Ctrl+S with typing in progress: `commit` sends the typing (returning whether there was any), then
 * the host saves once it has applied it. Without typing, the key goes to VS Code as usual.
 */
export function commitOnSave(commit: () => boolean): void {
  window.addEventListener(
    'keydown',
    (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== 's' || !commit()) return;
      // VS Code would save before the host has applied the typing: ask the host to save after it.
      e.preventDefault();
      e.stopImmediatePropagation();
      send({ type: 'save' });
    },
    true,
  );
}
