import type { MediaPayload, Presentation } from '@office-workbench/core';
import { el, getState, onHostMessage, patchState, send, showMessage } from '../common/host';
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

const toolbar = el('div', 'toolbar');
const prev = el('button', 'tool', '‹');
const next = el('button', 'tool', '›');
const counter = el('span', 'label');
prev.title = 'Previous slide';
next.title = 'Next slide';
toolbar.append(prev, counter, next, el('span', 'spacer'));
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
  const delta: Record<string, number> = { ArrowUp: -1, ArrowLeft: -1, PageUp: -1, ArrowDown: 1, ArrowRight: 1, PageDown: 1, ' ': 1 };
  if (e.key === 'Home') go(0);
  else if (e.key === 'End') go((presentation?.slides.length ?? 1) - 1);
  else if (delta[e.key]) go(current + delta[e.key]);
  else return;
  e.preventDefault();
});

onHostMessage((message) => {
  if (message.type === 'error') return showMessage(app, message.message, 'error');
  if (message.type !== 'presentation') return;
  presentation = message.presentation;
  setMedia(message.media);
  if (!presentation.slides.length) return showMessage(app, 'This presentation has no slides.');
  current = Math.min(current, presentation.slides.length - 1);
  renderThumbs();
  app.replaceChildren(toolbar, body);
  go(current);
});

showMessage(app, 'Loading presentation…');
send({ type: 'ready' });

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
      return item;
    }),
  );
}

function go(index: number): void {
  if (!presentation) return;
  current = Math.min(Math.max(index, 0), presentation.slides.length - 1);
  patchState<State>({ slide: current });
  const slide = presentation.slides[current];
  stage.replaceChildren(renderSlide(presentation, slide, media));
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
  slideEl.style.left = `${(stage.clientWidth - presentation.width * scale) / 2}px`;
  slideEl.style.top = `${(stage.clientHeight - presentation.height * scale) / 2}px`;
}
