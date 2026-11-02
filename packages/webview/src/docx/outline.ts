// The edit panel of the Word viewer: every paragraph of the body and of table cells as a text
// box. Edits are sent as operations when a paragraph loses focus (or on Enter / Backspace for
// splitting and deleting); the host replies with a new outline and preview.
import type { DocxBlock, DocxEditOp, DocxParagraph, DocxTable } from '@office-workbench/core';
import { el } from '../common/host';

export interface OutlineOptions {
  send(op: DocxEditOp): void;
}

interface Focus {
  id: string;
  start: number;
  end: number;
}

export class Outline {
  readonly element = el('div', 'docx-outline');
  private readonly list = el('div', 'outline-list');
  /** Paragraph text boxes by id, with the text the document has for them. */
  private boxes = new Map<string, { box: HTMLTextAreaElement; text: string }>();
  /** Where to put the caret after the next update (the paragraph an Enter or Backspace moved to). */
  private pending?: Focus;

  constructor(private readonly options: OutlineOptions) {
    const header = el('div', 'outline-header', 'Edit paragraphs');
    const help = el('div', 'outline-help', 'Enter: new paragraph · Shift+Enter: line break · Backspace in an empty paragraph: delete it · Esc: undo typing');
    this.element.append(header, help, this.list);
  }

  /** Whether the focused paragraph has typing that is not in the document yet. */
  get typing(): boolean {
    const focused = this.focused();
    return !!focused && focused.box.value !== focused.text;
  }

  /** Sends the focused paragraph's typing, keeping the focus there. */
  commitTyping(): void {
    const focused = this.focused();
    if (focused && focused.box.value !== focused.text) {
      this.commit(focused.id, focused.box);
    }
  }

  /** An edit was refused: the caret stays where it is. */
  refused(): void {
    this.pending = undefined;
  }

  render(blocks: DocxBlock[]): void {
    // Typing in progress survives an update when its paragraph still holds the same text.
    const active = this.focused();
    const draft = active && { id: active.id, value: active.box.value, text: active.text, start: active.box.selectionStart, end: active.box.selectionEnd };
    const scroll = this.list.scrollTop;
    this.boxes = new Map();
    this.list.replaceChildren(...blocks.map((b) => this.block(b)));
    this.list.scrollTop = scroll;

    const target = this.pending ?? (draft && { id: draft.id, start: draft.start, end: draft.end });
    this.pending = undefined;
    const entry = target && this.boxes.get(target.id);
    if (!entry) return;
    if (draft && draft.id === target.id && draft.value !== draft.text && entry.text === draft.text) {
      entry.box.value = draft.value;
      fit(entry.box);
    }
    entry.box.focus({ preventScroll: true });
    entry.box.setSelectionRange(Math.min(target.start, entry.box.value.length), Math.min(target.end, entry.box.value.length));
    entry.box.scrollIntoView({ block: 'nearest' });
  }

  private block(block: DocxBlock): HTMLElement {
    switch (block.type) {
      case 'paragraph':
        return this.paragraph(block);
      case 'table':
        return this.table(block);
      case 'other': {
        const node = el('div', 'outline-row outline-other');
        node.append(el('span', 'outline-marker', '▣'), el('div', 'outline-locked', block.text || block.label));
        node.title = `${block.label} (not editable here)`;
        return node;
      }
    }
  }

  private paragraph(p: DocxParagraph): HTMLElement {
    const row = el('div', 'outline-row');
    const marker = el('span', 'outline-marker', p.heading ? `H${p.heading}` : p.list ? (p.list.ordered ? '1.' : '•') : '¶');
    marker.title = p.style ?? 'Normal';
    if (p.list) row.style.paddingLeft = `${p.list.level * 16}px`;
    row.classList.toggle('heading', !!p.heading);
    if (p.locked) {
      // Paragraphs without text (a picture, a page break) show why they are locked instead.
      const text = el('div', 'outline-locked', p.text || p.locked);
      row.title = p.locked;
      row.append(marker, text, el('span', 'outline-lock', '🔒'));
      return row;
    }
    const box = el('textarea', 'outline-text');
    box.value = p.text;
    box.rows = 1;
    box.spellcheck = false;
    box.dataset.id = p.id;
    this.boxes.set(p.id, { box, text: p.text });
    box.addEventListener('input', () => fit(box));
    box.addEventListener('keydown', (e) => this.onKey(e, p, box));
    box.addEventListener('blur', () => this.commit(p.id, box));
    row.append(marker, box);
    requestAnimationFrame(() => fit(box));
    return row;
  }

  private table(table: DocxTable): HTMLElement {
    const wrap = el('div', 'outline-table');
    const t = el('table');
    for (const row of table.rows) {
      const tr = el('tr');
      for (const cell of row.cells) {
        const td = el('td');
        if (cell.colSpan > 1) td.colSpan = cell.colSpan;
        if (cell.merged) td.classList.add('merged');
        else td.append(...cell.blocks.map((b) => this.block(b)));
        tr.append(td);
      }
      t.append(tr);
    }
    wrap.append(t);
    return wrap;
  }

  private onKey(e: KeyboardEvent, p: DocxParagraph, box: HTMLTextAreaElement): void {
    if (e.isComposing || e.keyCode === 229) return;
    const { selectionStart: start, selectionEnd: end, value } = box;
    if (e.key === 'Enter' && !e.shiftKey && !e.altKey && !e.ctrlKey && !e.metaKey) {
      // Like Word: the text after the caret moves to a new paragraph below.
      e.preventDefault();
      const before = value.slice(0, start);
      const after = value.slice(end);
      this.pending = { id: siblingId(p.id, 1), start: 0, end: 0 };
      this.boxes.get(p.id)!.text = value;
      this.options.send({ type: 'insertParagraph', after: p.id, text: after, update: before === p.text ? undefined : before });
    } else if (e.key === 'Backspace' && !value && start === 0 && end === 0) {
      const previous = this.neighbour(p.id, -1);
      if (!previous) return;
      e.preventDefault();
      this.pending = { id: previous.id, start: previous.text.length, end: previous.text.length };
      this.options.send({ type: 'deleteParagraph', id: p.id });
    } else if (e.key === 'Escape') {
      e.preventDefault();
      box.value = p.text;
      fit(box);
    } else if (e.key === 'Tab' && !e.ctrlKey) {
      e.preventDefault();
      box.setRangeText('\t', start, end, 'end');
      fit(box);
    } else if ((e.key === 'ArrowUp' && start === 0 && end === 0) || (e.key === 'ArrowDown' && start === value.length)) {
      const target = this.neighbour(p.id, e.key === 'ArrowUp' ? -1 : 1);
      if (!target) return;
      e.preventDefault();
      target.box.focus();
      const at = e.key === 'ArrowUp' ? target.box.value.length : 0;
      target.box.setSelectionRange(at, at);
    }
  }

  private commit(id: string, box: HTMLTextAreaElement): void {
    const entry = this.boxes.get(id);
    if (entry?.box !== box || box.value === entry.text) return;
    entry.text = box.value;
    this.options.send({ type: 'setParagraphText', id, text: box.value });
  }

  /** The editable paragraph before or after `id` in document order. */
  private neighbour(id: string, delta: -1 | 1): { id: string; text: string; box: HTMLTextAreaElement } | undefined {
    const ids = [...this.boxes.keys()];
    const at = ids.indexOf(id) + delta;
    const target = ids[at];
    return target === undefined ? undefined : { id: target, ...this.boxes.get(target)! };
  }

  private focused(): { id: string; box: HTMLTextAreaElement; text: string } | undefined {
    const active = document.activeElement;
    if (!(active instanceof HTMLTextAreaElement) || !active.dataset.id) return undefined;
    const entry = this.boxes.get(active.dataset.id);
    return entry?.box === active ? { id: active.dataset.id, ...entry } : undefined;
  }
}

function siblingId(id: string, delta: number): string {
  const parts = id.split('/');
  parts[parts.length - 1] = String(Number(parts[parts.length - 1]) + delta);
  return parts.join('/');
}

/** Grows a text box to its content. */
function fit(box: HTMLTextAreaElement): void {
  box.style.height = 'auto';
  box.style.height = `${box.scrollHeight}px`;
}
