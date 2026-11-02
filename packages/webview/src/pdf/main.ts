import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
import { el, getState, onHostMessage, patchState, send, showMessage } from '../common/host';
import 'pdfjs-dist/web/pdf_viewer.css';
import '../common/base.css';
import './pdf.css';

// pdf_viewer.mjs reads the core library from this global, so it must be set before that module loads.
(globalThis as { pdfjsLib?: unknown }).pdfjsLib = pdfjsLib;
const viewerModule = import('pdfjs-dist/web/pdf_viewer.mjs');

interface State {
  page: number;
  scale: string;
  outline: boolean;
}

interface OutlineNode {
  title: string;
  dest: string | unknown[] | null;
  url: string | null;
  items: OutlineNode[];
}

const ZOOM_PRESETS: [string, string][] = [
  ['auto', 'Automatic'],
  ['page-fit', 'Page fit'],
  ['page-width', 'Page width'],
  ['0.5', '50%'],
  ['0.75', '75%'],
  ['1', '100%'],
  ['1.25', '125%'],
  ['1.5', '150%'],
  ['2', '200%'],
  ['3', '300%'],
];

const app = document.getElementById('app')!;
const assets = document.body.dataset.assets ?? '.';
let current: { destroy(): void } | undefined;

onHostMessage(async (message) => {
  if (message.type === 'error') return showMessage(app, message.message, 'error');
  if (message.type !== 'pdf') return;
  await open(message.data);
});

showMessage(app, 'Loading PDF…');
void startWorker().then(() => send({ type: 'ready' }));

/** Loads the worker bundle through a blob URL, since webviews block cross-origin worker scripts. */
async function startWorker(): Promise<void> {
  const source = await (await fetch(`${assets}/pdf.worker.js`)).blob();
  const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
  pdfjsLib.GlobalWorkerOptions.workerPort = new Worker(url);
}

async function open(data: Uint8Array): Promise<void> {
  const state = getState<State>();
  current?.destroy();
  current = undefined;

  let doc: PDFDocumentProxy;
  const loadingTask = pdfjsLib.getDocument({
    data,
    cMapUrl: `${assets}/pdfjs/cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${assets}/pdfjs/standard_fonts/`,
    wasmUrl: `${assets}/pdfjs/wasm/`,
    iccUrl: `${assets}/pdfjs/iccs/`,
    // Fetch data files on the main thread: relative to a blob worker they would not resolve.
    useWorkerFetch: false,
  });
  try {
    doc = await loadingTask.promise;
  } catch (e) {
    const name = (e as Error).name;
    const text = name === 'PasswordException' ? 'Password-protected PDFs are not supported yet.' : `This PDF could not be displayed.\n${(e as Error).message}`;
    return showMessage(app, text, 'error');
  }

  const { EventBus, PDFLinkService, PDFFindController, PDFViewer, LinkTarget } = await viewerModule;

  // Toolbar
  const toolbar = el('div', 'toolbar');
  const outlineToggle = el('button', 'tool', '☰');
  outlineToggle.title = 'Toggle outline';
  const prev = el('button', 'tool', '‹');
  prev.title = 'Previous page';
  const pageInput = el('input', 'page-input');
  pageInput.type = 'text';
  pageInput.inputMode = 'numeric';
  const pageCount = el('span', 'label', `/ ${doc.numPages}`);
  const next = el('button', 'tool', '›');
  next.title = 'Next page';
  const zoomOut = el('button', 'tool', '−');
  zoomOut.title = 'Zoom out';
  const zoomSelect = el('select', 'zoom-select');
  for (const [value, label] of ZOOM_PRESETS) zoomSelect.append(new Option(label, value));
  const customZoom = new Option('', 'custom');
  customZoom.hidden = true;
  zoomSelect.append(customZoom);
  const zoomIn = el('button', 'tool', '+');
  zoomIn.title = 'Zoom in';
  const search = el('input', 'search-input');
  search.type = 'search';
  search.placeholder = 'Find';
  const findPrev = el('button', 'tool', '↑');
  findPrev.title = 'Previous match';
  const findNext = el('button', 'tool', '↓');
  findNext.title = 'Next match';
  const findStatus = el('span', 'label find-status');
  toolbar.append(
    outlineToggle, el('span', 'sep'), prev, pageInput, pageCount, next, el('span', 'sep'), zoomOut, zoomSelect, zoomIn,
    el('span', 'spacer'), search, findPrev, findNext, findStatus,
  );

  // Body: outline sidebar + viewer
  const body = el('div', 'pdf-body');
  const sidebar = el('nav', 'outline');
  const wrap = el('div', 'viewer-wrap');
  const container = el('div', 'viewer-container');
  container.tabIndex = 0;
  const viewerEl = el('div', 'pdfViewer');
  container.append(viewerEl);
  wrap.append(container);
  body.append(sidebar, wrap);
  app.replaceChildren(toolbar, body);

  const eventBus = new EventBus();
  const linkService = new PDFLinkService({ eventBus, externalLinkTarget: LinkTarget.BLANK });
  const findController = new PDFFindController({ eventBus, linkService });
  const viewer = new PDFViewer({ container, viewer: viewerEl, eventBus, linkService, findController });
  linkService.setViewer(viewer);

  eventBus.on('pagesinit', () => {
    viewer.currentScaleValue = state.scale ?? 'auto';
    if (state.page && state.page <= doc.numPages) viewer.currentPageNumber = state.page;
  });
  eventBus.on('pagechanging', ({ pageNumber }: { pageNumber: number }) => {
    pageInput.value = String(pageNumber);
    prev.disabled = pageNumber <= 1;
    next.disabled = pageNumber >= doc.numPages;
    patchState<State>({ page: pageNumber });
  });
  eventBus.on('scalechanging', ({ scale, presetValue }: { scale: number; presetValue?: string }) => {
    if (presetValue && ZOOM_PRESETS.some(([v]) => v === presetValue)) {
      zoomSelect.value = presetValue;
    } else {
      customZoom.textContent = `${Math.round(scale * 100)}%`;
      zoomSelect.value = 'custom';
    }
    patchState<State>({ scale: presetValue ?? String(scale) });
  });
  eventBus.on('updatefindmatchescount', ({ matchesCount }: { matchesCount: { current: number; total: number } }) => {
    findStatus.textContent = matchesCount.total ? `${matchesCount.current} / ${matchesCount.total}` : '';
  });
  eventBus.on('updatefindcontrolstate', ({ state: findState, matchesCount }: { state: number; matchesCount: { current: number; total: number } }) => {
    // FindState.NOT_FOUND === 1
    findStatus.textContent = findState === 1 ? 'No matches' : matchesCount.total ? `${matchesCount.current} / ${matchesCount.total}` : '';
  });

  // Preset zooms (auto, page-width, …) depend on the container size, which changes when the
  // outline is toggled or the editor is resized; the bare PDFViewer does not refit on its own.
  new ResizeObserver(() => {
    const value = viewer.currentScaleValue;
    if (doc.numPages && Number.isNaN(Number(value))) viewer.currentScaleValue = value;
  }).observe(container);

  viewer.setDocument(doc);
  linkService.setDocument(doc, null);
  pageInput.value = '1';

  // Controls
  prev.onclick = () => viewer.previousPage();
  next.onclick = () => viewer.nextPage();
  pageInput.onchange = () => {
    const n = Number(pageInput.value);
    if (Number.isInteger(n) && n >= 1 && n <= doc.numPages) viewer.currentPageNumber = n;
    else pageInput.value = String(viewer.currentPageNumber);
  };
  zoomSelect.onchange = () => {
    if (zoomSelect.value !== 'custom') viewer.currentScaleValue = zoomSelect.value;
  };
  zoomIn.onclick = () => viewer.increaseScale();
  zoomOut.onclick = () => viewer.decreaseScale();
  container.addEventListener(
    'wheel',
    (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      if (e.deltaY < 0) viewer.increaseScale();
      else viewer.decreaseScale();
    },
    { passive: false },
  );

  const find = (type: '' | 'again', findPrevious = false) =>
    eventBus.dispatch('find', {
      source: null,
      type,
      query: search.value,
      caseSensitive: false,
      entireWord: false,
      highlightAll: true,
      findPrevious,
      matchDiacritics: false,
    });
  search.oninput = () => find('');
  search.onkeydown = (e) => {
    if (e.key === 'Enter') find('again', e.shiftKey);
    if (e.key === 'Escape') {
      search.value = '';
      find('');
      container.focus();
    }
  };
  findNext.onclick = () => find('again');
  findPrev.onclick = () => find('again', true);
  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
      e.preventDefault();
      search.focus();
      search.select();
    }
  });

  // Outline
  const outline = ((await doc.getOutline()) ?? []) as OutlineNode[];
  const showOutline = (visible: boolean) => {
    body.classList.toggle('with-outline', visible);
    patchState<State>({ outline: visible });
  };
  if (outline.length) {
    sidebar.append(renderOutline(outline, (dest) => void linkService.goToDestination(dest)));
    showOutline(state.outline ?? true);
    outlineToggle.onclick = () => showOutline(!body.classList.contains('with-outline'));
  } else {
    outlineToggle.disabled = true;
    outlineToggle.title = 'This PDF has no outline';
  }

  current = {
    destroy: () => {
      viewer.setDocument(null as unknown as PDFDocumentProxy);
      void loadingTask.destroy();
    },
  };
}

function renderOutline(items: OutlineNode[], go: (dest: string | unknown[]) => void): HTMLElement {
  const list = el('ul');
  for (const item of items) {
    const li = el('li');
    const row = el('div', 'outline-item');
    if (item.items.length) {
      const twisty = el('span', 'twisty', '▾');
      twisty.onclick = (e) => {
        e.stopPropagation();
        li.classList.toggle('collapsed');
      };
      row.append(twisty);
    }
    row.append(el('span', 'outline-title', item.title));
    if (item.dest) {
      row.classList.add('link');
      row.onclick = () => go(item.dest!);
    } else if (item.url) {
      row.title = item.url;
    }
    li.append(row);
    if (item.items.length) li.append(renderOutline(item.items, go));
    list.append(li);
  }
  return list;
}
