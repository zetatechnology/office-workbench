// The edit panel of the Word viewer: every paragraph of the body and of table cells as a text
// box. Edits are sent as operations when a paragraph loses focus (or on Enter / Backspace for
// splitting and deleting); the host answers each one ("edited") after sending the new outline.
//
// The panel shows its own edits at once: Enter puts the new paragraph on screen (typing goes
// straight into it) and Backspace removes the paragraph, and the ids of the paragraphs after it are
// renumbered. Ids are positions, so the ids on screen are always the ones the document will have
// once the edits sent so far are applied in order. Outlines from the host replace the screen only
// when none of this panel's edits are outstanding. After an insert or delete, later edits wait for
// its answer; if it is refused, they are dropped and the screen goes back to the host's outline.
import type { DocxBlock, DocxEditOp, DocxParagraph, DocxTable } from '@office-workbench/core';
import { el } from '../common/host';
import { moveId, siblingId } from './ids';

export interface OutlineOptions {
  send(op: DocxEditOp): void;
}

export class Outline {
  readonly element = el('div', 'docx-outline');
  private readonly list = el('div', 'outline-list');
  /** Paragraph text boxes by id, in document order. */
  private boxes = new Map<string, HTMLTextAreaElement>();
  /** The text each box's paragraph has (or will have once the edits sent so far are applied). */
  private readonly texts = new WeakMap<HTMLTextAreaElement, string>();
  /** Edits sent and not yet answered, oldest first. */
  private readonly sent: DocxEditOp[] = [];
  /** An insert or delete waiting for its answer, and the edits made meanwhile. */
  private waiting?: DocxEditOp;
  private held: DocxEditOp[] = [];
  /** The newest outline from the host, shown once this panel's edits are all answered. */
  private latest?: DocxBlock[];
  /** An edit was refused: typing on screen may belong to paragraphs that no longer line up. */
  private refused = false;

  constructor(private readonly options: OutlineOptions) {
    const header = el('div', 'outline-header', 'Edit paragraphs');
    const help = el('div', 'outline-help', 'Enter: new paragraph · Shift+Enter: line break · Backspace in an empty paragraph: delete it · Esc: undo typing');
    this.element.append(header, help, this.list);
  }

  /** Whether the focused paragraph has typing that is not in the document yet. */
  get typing(): boolean {
    const box = this.focused();
    return !!box && box.value !== this.texts.get(box);
  }

  /** Sends the focused paragraph's typing, keeping the focus there; returns whether there was any. */
  commitTyping(): boolean {
    const box = this.focused();
    if (!box || box.value === this.texts.get(box)) return false;
    this.commit(box);
    return true;
  }

  /** A new outline from the host (after an edit here or elsewhere, or an undo). */
  update(blocks: DocxBlock[]): void {
    this.latest = blocks;
    if (!this.outstanding) this.render(blocks);
  }

  /** The whole document was (re)loaded: anything in flight no longer applies. */
  reset(blocks: DocxBlock[]): void {
    this.sent.length = 0;
    this.waiting = undefined;
    this.held = [];
    this.refused = false;
    this.latest = blocks;
    this.render(blocks);
  }

  /** The host answered the oldest edit this panel sent. */
  answered(ok: boolean): void {
    const op = this.sent.shift();
    if (!ok) this.refused = true;
    if (op && op === this.waiting) {
      this.waiting = undefined;
      const held = this.held;
      this.held = [];
      // Edits made after a refused insert or delete were made on a screen that showed it.
      if (ok) for (const next of held) this.dispatch(next);
    }
    if (!this.outstanding && this.latest) this.render(this.latest);
  }

  private get outstanding(): boolean {
    return this.sent.length > 0 || this.held.length > 0;
  }

  private dispatch(op: DocxEditOp): void {
    if (this.waiting) {
      this.held.push(op);
      return;
    }
    this.sent.push(op);
    if (op.type !== 'setParagraphText') this.waiting = op;
    this.options.send(op);
  }

  private render(blocks: DocxBlock[]): void {
    // Typing in progress survives when its paragraph still holds the text the typing started from.
    const active = this.focused();
    const draft = active && {
      id: active.dataset.id!,
      value: active.value,
      text: this.texts.get(active),
      start: active.selectionStart,
      end: active.selectionEnd,
    };
    const keepTyping = !this.refused;
    this.refused = false;
    const scroll = this.list.scrollTop;
    this.boxes = new Map();
    this.list.replaceChildren(...blocks.map((b) => this.block(b)));
    this.list.scrollTop = scroll;

    const box = draft && this.boxes.get(draft.id);
    if (!box) return;
    if (keepTyping && draft.value !== draft.text && this.texts.get(box) === draft.text) {
      box.value = draft.value;
      fit(box);
    }
    box.focus({ preventScroll: true });
    box.setSelectionRange(Math.min(draft.start, box.value.length), Math.min(draft.end, box.value.length));
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
    const box = this.textBox(p.id, p.text);
    this.boxes.set(p.id, box);
    row.append(marker, box);
    return row;
  }

  private textBox(id: string, text: string): HTMLTextAreaElement {
    const box = el('textarea', 'outline-text');
    box.value = text;
    box.rows = 1;
    box.spellcheck = false;
    box.dataset.id = id;
    this.texts.set(box, text);
    box.addEventListener('input', () => fit(box));
    box.addEventListener('keydown', (e) => this.onKey(e, box));
    box.addEventListener('blur', () => this.commit(box));
    requestAnimationFrame(() => fit(box));
    return box;
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

  private onKey(e: KeyboardEvent, box: HTMLTextAreaElement): void {
    if (e.isComposing || e.keyCode === 229) return;
    const id = box.dataset.id!;
    const { selectionStart: start, selectionEnd: end, value } = box;
    if (e.key === 'Enter' && !e.shiftKey && !e.altKey && !e.ctrlKey && !e.metaKey) {
      // Like Word: the text after the caret moves to a new paragraph below, which gets the caret.
      e.preventDefault();
      const before = value.slice(0, start);
      const after = value.slice(end);
      const op: DocxEditOp = { type: 'insertParagraph', after: id, text: after };
      if (before !== this.texts.get(box)) op.update = before;
      this.dispatch(op);
      box.value = before;
      this.texts.set(box, before);
      fit(box);
      this.renumber(op);
      const row = box.closest('.outline-row') as HTMLElement;
      const added = row.cloneNode(false) as HTMLElement;
      const created = this.textBox(siblingId(id, 1), after);
      added.append(row.querySelector('.outline-marker')!.cloneNode(true), created);
      row.after(added);
      this.reindex();
      created.focus();
      created.setSelectionRange(0, 0);
    } else if (e.key === 'Backspace' && !value && start === 0 && end === 0) {
      const previous = this.neighbour(id, -1);
      if (!previous) return;
      e.preventDefault();
      const op: DocxEditOp = { type: 'deleteParagraph', id };
      this.dispatch(op);
      // Without its id the box cannot commit typing (its blur fires on removal).
      box.removeAttribute('data-id');
      box.closest('.outline-row')!.remove();
      this.renumber(op);
      this.reindex();
      previous.focus();
      previous.setSelectionRange(previous.value.length, previous.value.length);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      box.value = this.texts.get(box) ?? box.value;
      fit(box);
    } else if (e.key === 'Tab' && !e.ctrlKey) {
      e.preventDefault();
      box.setRangeText('\t', start, end, 'end');
      fit(box);
    } else if ((e.key === 'ArrowUp' && start === 0 && end === 0) || (e.key === 'ArrowDown' && start === value.length)) {
      const target = this.neighbour(id, e.key === 'ArrowUp' ? -1 : 1);
      if (!target) return;
      e.preventDefault();
      target.focus();
      const at = e.key === 'ArrowUp' ? target.value.length : 0;
      target.setSelectionRange(at, at);
    }
  }

  private commit(box: HTMLTextAreaElement): void {
    const id = box.dataset.id;
    if (!id || this.boxes.get(id) !== box || box.value === this.texts.get(box)) return;
    this.texts.set(box, box.value);
    this.dispatch({ type: 'setParagraphText', id, text: box.value });
  }

  /** Moves the ids on screen past an insert or delete just made here. */
  private renumber(op: DocxEditOp): void {
    for (const box of this.boxes.values()) {
      if (!box.dataset.id) continue; // the deleted paragraph
      const moved = moveId(box.dataset.id, op);
      if (moved !== undefined) box.dataset.id = moved;
    }
  }

  /** Rebuilds the id → box map from the boxes on screen, in document order. */
  private reindex(): void {
    this.boxes = new Map(
      Array.from(this.list.querySelectorAll<HTMLTextAreaElement>('textarea.outline-text')).map((box) => [box.dataset.id!, box]),
    );
  }

  /** The editable paragraph before or after `id` in document order. */
  private neighbour(id: string, delta: -1 | 1): HTMLTextAreaElement | undefined {
    const ids = [...this.boxes.keys()];
    const at = ids.indexOf(id);
    return at < 0 ? undefined : this.boxes.get(ids[at + delta]);
  }

  private focused(): HTMLTextAreaElement | undefined {
    const active = document.activeElement;
    if (!(active instanceof HTMLTextAreaElement) || !active.dataset.id) return undefined;
    return this.boxes.get(active.dataset.id) === active ? active : undefined;
  }
}

/** Grows a text box to its content. */
function fit(box: HTMLTextAreaElement): void {
  box.style.height = 'auto';
  box.style.height = `${box.scrollHeight}px`;
}
