import { renderAsync } from 'docx-preview';
import type { DocxBlock } from '@office-workbench/core';
import { el, getState, onHostMessage, patchState, send, showMessage } from '../common/host';
import { commitOnSave, guardUndoKeys } from '../common/text-edit';
import { showToast } from '../common/toast';
import { Outline } from './outline';
import '../common/base.css';
import './docx.css';

interface State {
  zoom: number;
  scrollTop: number;
  editing: boolean;
}

const ZOOM_STEPS = [0.25, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 3];

const app = document.getElementById('app')!;
let zoom = getState<State>().zoom ?? 1;
let editing = getState<State>().editing ?? false;
let blocks: DocxBlock[] | undefined;
/** Renders are queued so a quick series of edits shows the latest document last. */
let rendering: Promise<void> = Promise.resolve();
const PREVIEW_DELAY = 600;
let previewTimer: ReturnType<typeof setTimeout> | undefined;
/** Page data requested and not received yet (those replies leave the edit panel alone). */
let previewsAsked = 0;

const toolbar = el('div', 'toolbar');
const zoomOut = el('button', 'tool', '−');
const zoomLabel = el('button', 'tool zoom-label');
const zoomIn = el('button', 'tool', '+');
const fitWidth = el('button', 'tool', 'Fit width');
const editToggle = el('button', 'tool edit-toggle', 'Edit');
const pageLabel = el('span', 'label');
zoomOut.title = 'Zoom out';
zoomIn.title = 'Zoom in';
zoomLabel.title = 'Reset zoom';
editToggle.title = 'Show the paragraphs for editing next to the page view';
toolbar.append(zoomOut, zoomLabel, zoomIn, fitWidth, editToggle, el('span', 'spacer'), pageLabel);

const split = el('div', 'docx-split');
const host = el('div', 'docx-host');
const content = el('div', 'docx-content');
host.append(content);
const outline = new Outline({ send: (op) => send({ type: 'edit', op }) });
split.append(host, outline.element);
guardUndoKeys(() => outline.typing);
commitOnSave(() => outline.commitTyping());

const pages = () => Array.from(content.querySelectorAll<HTMLElement>('section.docx'));
const applyZoom = (value: number) => {
  zoom = Math.min(5, Math.max(0.1, value));
  content.style.setProperty('zoom', String(zoom));
  zoomLabel.textContent = `${Math.round(zoom * 100)}%`;
  patchState<State>({ zoom });
};
const step = (dir: 1 | -1) => {
  const next = dir > 0 ? ZOOM_STEPS.find((z) => z > zoom + 0.001) : [...ZOOM_STEPS].reverse().find((z) => z < zoom - 0.001);
  applyZoom(next ?? zoom);
};
const updatePageLabel = () => {
  const list = pages();
  const middle = host.getBoundingClientRect().top + host.clientHeight / 3;
  const current = list.findIndex((p) => p.getBoundingClientRect().bottom > middle);
  pageLabel.textContent = list.length ? `Page ${Math.max(current, 0) + 1} / ${list.length}` : '';
};

zoomOut.onclick = () => step(-1);
zoomIn.onclick = () => step(1);
zoomLabel.onclick = () => applyZoom(1);
fitWidth.onclick = () => {
  const first = pages()[0];
  // The wrapper's 24px padding is zoomed along with the page, so fit page + padding.
  const pageWidth = first ? first.getBoundingClientRect().width / zoom : 0;
  if (pageWidth) applyZoom(host.clientWidth / (pageWidth + 48));
};
editToggle.onclick = () => setEditing(!editing);
host.addEventListener(
  'wheel',
  (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    step(e.deltaY < 0 ? 1 : -1);
  },
  { passive: false },
);
host.addEventListener('scroll', () => {
  updatePageLabel();
  patchState<State>({ scrollTop: host.scrollTop });
});

onHostMessage((message) => {
  switch (message.type) {
    case 'error':
      return showMessage(app, message.message, 'error');
    case 'notice':
      return showToast(message.message);
    case 'edited':
      return outline.answered(message.ok);
    case 'docxOutline':
      blocks = message.blocks;
      outline.render(blocks);
      schedulePreview();
      return;
    case 'docx': {
      if (previewsAsked > 0) previewsAsked--;
      else {
        // Loaded, reverted or changed on disk.
        blocks = message.blocks;
        if (blocks) outline.reset(blocks);
        setEditing(editing);
      }
      editToggle.hidden = !blocks;
      const data = message.data;
      rendering = rendering.then(() => render(data));
      return;
    }
  }
});

/** After edits, the page view is rebuilt once they pause (building the file and laying out pages is slow). */
function schedulePreview(): void {
  host.classList.add('stale');
  clearTimeout(previewTimer);
  previewTimer = setTimeout(() => {
    previewsAsked++;
    send({ type: 'preview' });
  }, PREVIEW_DELAY);
}

showMessage(app, 'Loading document…');
send({ type: 'ready' });

function setEditing(value: boolean): void {
  editing = value && !!blocks;
  patchState<State>({ editing });
  editToggle.classList.toggle('active', editing);
  outline.element.hidden = !editing;
}

async function render(data: Uint8Array): Promise<void> {
  const firstRender = !app.contains(split);
  const previousScroll = firstRender ? (getState<State>().scrollTop ?? 0) : host.scrollTop;
  // Render into a hidden layer, then swap, so the page view does not flash empty after each edit.
  // (docx-preview measures text for tab stops, so the layer must be laid out.)
  const next = el('div', 'docx-content docx-staging');
  if (firstRender) app.replaceChildren(toolbar, split);
  host.append(next);
  try {
    await renderAsync(data, next, undefined, {
      className: 'docx',
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      renderComments: false,
      renderChanges: false,
      useBase64URL: false,
      experimental: true,
    });
  } catch (e) {
    next.remove();
    if (firstRender) showMessage(app, `This document could not be displayed.\n${(e as Error).message}`, 'error');
    else showToast(`The page view could not be updated: ${(e as Error).message}`);
    return;
  }
  next.remove();
  content.replaceChildren(...Array.from(next.childNodes));
  host.classList.remove('stale');
  applyZoom(zoom);
  host.scrollTop = previousScroll;
  updatePageLabel();
}
