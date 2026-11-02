import { renderAsync } from 'docx-preview';
import { el, getState, onHostMessage, patchState, send, showMessage } from '../common/host';
import '../common/base.css';
import './docx.css';

interface State {
  zoom: number;
  scrollTop: number;
}

const ZOOM_STEPS = [0.25, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 3];

const app = document.getElementById('app')!;
let zoom = getState<State>().zoom ?? 1;

onHostMessage(async (message) => {
  if (message.type === 'error') return showMessage(app, message.message, 'error');
  if (message.type !== 'docx') return;
  await render(message.data);
});

showMessage(app, 'Loading document…');
send({ type: 'ready' });

async function render(data: Uint8Array): Promise<void> {
  const previousScroll = document.querySelector('.docx-host')?.scrollTop ?? getState<State>().scrollTop ?? 0;

  const toolbar = el('div', 'toolbar');
  const zoomOut = el('button', 'tool', '−');
  const zoomLabel = el('button', 'tool zoom-label');
  const zoomIn = el('button', 'tool', '+');
  const fitWidth = el('button', 'tool', 'Fit width');
  const pageLabel = el('span', 'label');
  zoomOut.title = 'Zoom out';
  zoomIn.title = 'Zoom in';
  zoomLabel.title = 'Reset zoom';
  toolbar.append(zoomOut, zoomLabel, zoomIn, fitWidth, el('span', 'spacer'), pageLabel);

  const host = el('div', 'docx-host');
  const content = el('div', 'docx-content');
  host.append(content);
  app.replaceChildren(toolbar, host);

  try {
    await renderAsync(data, content, undefined, {
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
    showMessage(app, `This document could not be displayed.\n${(e as Error).message}`, 'error');
    return;
  }

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

  applyZoom(zoom);
  host.scrollTop = previousScroll;
  updatePageLabel();
}
