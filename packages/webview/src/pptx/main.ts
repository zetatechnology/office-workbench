import type { MediaPayload, PptxEditOp, Presentation, ShapeElement, TableElement } from '@office-workbench/core';
import { el, getState, onHostMessage, patchState, send, showMessage } from '../common/host';
import { showMenu } from '../common/menu';
import { commitOnSave, guardUndoKeys, TextEdit } from '../common/text-edit';
import { showToast } from '../common/toast';
import { renderSlide, type MediaUrls } from './render';
import '../common/base.css';
import './pptx.css';

interface State {
  slide: number;
}

const THUMB_WIDTH = 168;

const app = document.getElementById('app')!;
let presentation: Presentation | undefined;
let media: MediaUrls = new Map();
let current = getState<State>().slide ?? 0;
let editable = false;
let editor: TextEdit | undefined;
/** An update that arrived while text was being typed; shown when editing ends. */
let deferred: Presentation | undefined;

const toolbar = el('div', 'toolbar');
const prev = el('button', 'tool', '‹');
const next = el('button', 'tool', '›');
const counter = el('span', 'label');
const hint = el('span', 'label', 'Double-click text to edit');
prev.title = 'Previous slide';
next.title = 'Next slide';
toolbar.append(prev, counter, next, el('span', 'spacer'), hint);
const body = el('div', 'pptx-body');
const thumbs = el('div', 'thumbs');
const main = el('div', 'stage-column');
const stage = el('div', 'stage');
const notes = el('div', 'notes');
main.append(stage, notes);
body.append(thumbs, main);

prev.onclick = () => go(current - 1);
next.onclick = () => go(current + 1);
new ResizeObserver(() => layoutStage()).observe(stage);
window.addEventListener('keydown', (e) => {
  if (editor?.active || (e.target instanceof HTMLElement && e.target.isContentEditable)) return;
  const delta: Record<string, number> = { ArrowUp: -1, ArrowLeft: -1, PageUp: -1, ArrowDown: 1, ArrowRight: 1, PageDown: 1, ' ': 1 };
  if (e.key === 'Home') go(0);
  else if (e.key === 'End') go((presentation?.slides.length ?? 1) - 1);
  else if (delta[e.key]) go(current + delta[e.key]);
  else return;
  e.preventDefault();
});
guardUndoKeys(() => !!editor?.dirty);
commitOnSave(() => editor?.end(true));
stage.addEventListener('dblclick', (e) => startEditing(e));

onHostMessage((message) => {
  if (message.type === 'error') return showMessage(app, message.message, 'error');
  if (message.type === 'notice') {
    showToast(message.message);
    // A refused edit leaves the typed text on screen; show the slide as it is.
    if (!editor?.active) go(current);
    return;
  }
  if (message.type !== 'presentation') return;
  editable = !!message.editable;
  hint.hidden = !editable;
  if (!message.update) setMedia(message.media);
  if (message.activate !== undefined) current = message.activate;
  if (editor?.active) {
    if (editor.dirty) {
      deferred = message.presentation;
      return;
    }
    editor.end(false);
  }
  show(message.presentation);
});

showMessage(app, 'Loading presentation…');
send({ type: 'ready' });

function show(next: Presentation): void {
  presentation = next;
  deferred = undefined;
  if (!presentation.slides.length) return showMessage(app, 'This presentation has no slides.');
  current = Math.min(current, presentation.slides.length - 1);
  renderThumbs();
  if (!app.contains(body)) app.replaceChildren(toolbar, body);
  go(current);
}

function setMedia(payloads: MediaPayload[]): void {
  for (const url of media.values()) URL.revokeObjectURL(url);
  media = new Map(payloads.map((m) => [m.path, URL.createObjectURL(new Blob([m.data as BlobPart], { type: m.contentType }))]));
}

function renderThumbs(): void {
  if (!presentation) return;
  const p = presentation;
  const scale = THUMB_WIDTH / p.width;
  thumbs.replaceChildren(
    ...p.slides.map((slide, i) => {
      const item = el('button', 'thumb');
      item.classList.toggle('hidden-slide', slide.hidden);
      item.title = slide.hidden ? `Slide ${i + 1} (hidden)` : `Slide ${i + 1}`;
      const frame = el('div', 'thumb-frame');
      frame.style.width = `${THUMB_WIDTH}px`;
      frame.style.height = `${p.height * scale}px`;
      const rendered = renderSlide(p, slide, media);
      rendered.style.transform = `scale(${scale})`;
      frame.append(rendered);
      item.append(el('span', 'thumb-number', String(i + 1)), frame);
      item.onclick = () => go(i);
      if (editable) addSlideActions(item, i, p.slides.length);
      return item;
    }),
  );
}

/** Context menu and drag and drop for one slide thumbnail. */
function addSlideActions(item: HTMLElement, index: number, count: number): void {
  const edit = (op: PptxEditOp) => send({ type: 'edit', op });
  item.oncontextmenu = (e) => {
    e.preventDefault();
    go(index);
    showMenu(e.clientX, e.clientY, [
      { label: 'Duplicate Slide', run: () => edit({ type: 'duplicateSlide', slide: index }) },
      { label: 'Delete Slide', run: () => edit({ type: 'deleteSlide', slide: index }), disabled: count === 1 },
      null,
      { label: 'Move Up', run: () => edit({ type: 'moveSlide', slide: index, to: index - 1 }), disabled: index === 0 },
      { label: 'Move Down', run: () => edit({ type: 'moveSlide', slide: index, to: index + 1 }), disabled: index === count - 1 },
    ]);
  };
  item.draggable = true;
  item.ondragstart = (e) => {
    e.dataTransfer?.setData('application/x-office-slide', String(index));
    if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move';
  };
  const target = (e: DragEvent) => {
    // Dropping on the lower half of a thumbnail puts the slide after it.
    const rect = item.getBoundingClientRect();
    return e.clientY > rect.top + rect.height / 2 ? index + 1 : index;
  };
  item.ondragover = (e) => {
    if (!e.dataTransfer?.types.includes('application/x-office-slide')) return;
    e.preventDefault();
    const after = target(e) > index;
    item.classList.toggle('drop-before', !after);
    item.classList.toggle('drop-after', after);
  };
  item.ondragleave = () => item.classList.remove('drop-before', 'drop-after');
  item.ondrop = (e) => {
    item.classList.remove('drop-before', 'drop-after');
    const from = Number(e.dataTransfer?.getData('application/x-office-slide'));
    if (!Number.isInteger(from)) return;
    e.preventDefault();
    let to = target(e);
    if (to > from) to--;
    if (to !== from) edit({ type: 'moveSlide', slide: from, to });
  };
}

function go(index: number): void {
  if (!presentation) return;
  if (editor?.active) editor.end(true);
  current = Math.min(Math.max(index, 0), presentation.slides.length - 1);
  patchState<State>({ slide: current });
  const slide = presentation.slides[current];
  stage.replaceChildren(renderSlide(presentation, slide, media));
  stage.classList.toggle('editable', editable);
  layoutStage();
  notes.textContent = slide.notes ?? '';
  notes.hidden = !slide.notes;
  counter.textContent = `Slide ${current + 1} / ${presentation.slides.length}${slide.hidden ? ' (hidden)' : ''}`;
  prev.disabled = current === 0;
  next.disabled = current === presentation.slides.length - 1;
  thumbs.querySelectorAll('.thumb').forEach((t, i) => t.classList.toggle('active', i === current));
  thumbs.children[current]?.scrollIntoView({ block: 'nearest' });
}

function layoutStage(): void {
  const slideEl = stage.firstElementChild as HTMLElement | null;
  if (!slideEl || !presentation) return;
  const margin = 24;
  const scale = Math.max(0.05, Math.min((stage.clientWidth - margin * 2) / presentation.width, (stage.clientHeight - margin * 2) / presentation.height));
  slideEl.style.transform = `scale(${scale})`;
  slideEl.style.setProperty('--scale', String(scale));
  slideEl.style.left = `${(stage.clientWidth - presentation.width * scale) / 2}px`;
  slideEl.style.top = `${(stage.clientHeight - presentation.height * scale) / 2}px`;
}

// --- Text editing ----------------------------------------------------------------------------

/** Double-click on a shape or table cell of the slide: edit its text where it is. */
function startEditing(e: MouseEvent): void {
  if (!editable || !presentation || editor?.active) return;
  const target = e.target as HTMLElement;
  const frame = target.closest<HTMLElement>('[data-shape-id]');
  if (!frame || !stage.contains(frame)) return;
  const slide = presentation.slides[current];
  const shape = frame.dataset.shapeId!;
  const element = slide.elements.find((x): x is ShapeElement | TableElement => (x.type === 'shape' || x.type === 'table') && x.textId === shape);
  if (!element) return;

  let host: HTMLElement | null;
  let op: (paragraphs: string[]) => PptxEditOp;
  if (element.type === 'table') {
    host = target.closest<HTMLElement>('td[data-row]');
    if (!host) return;
    const row = Number(host.dataset.row), col = Number(host.dataset.col);
    op = (paragraphs) => ({ type: 'setTableCellText', slide: current, shape, row, col, paragraphs });
  } else {
    host = frame.querySelector<HTMLElement>(':scope > .text');
    op = (paragraphs) => ({ type: 'setShapeText', slide: current, shape, paragraphs });
  }
  if (!host) return;
  e.preventDefault();
  frame.classList.add('editing-frame');
  editor = new TextEdit(host, {
    caret: { x: e.clientX, y: e.clientY },
    commit: (paragraphs) => {
      frame.classList.remove('editing-frame');
      send({ type: 'edit', op: op(paragraphs) });
      if (deferred) show(deferred);
    },
    close: () => {
      frame.classList.remove('editing-frame');
      if (deferred) show(deferred);
      else go(current);
    },
  });
}
