import type { SheetData, WorkbookInfo } from '@office-workbench/core';
import { columnName } from '@office-workbench/core/address';
import { el, getState, onHostMessage, patchState, send, showMessage } from '../common/host';
import { SheetGrid, type Selection } from './grid';
import '../common/base.css';
import './xlsx.css';

interface State {
  sheet: number;
  views: Record<number, { top: number; left: number; selection?: Selection }>;
}

const app = document.getElementById('app')!;
let info: WorkbookInfo | undefined;
let activeSheet = getState<State>().sheet ?? -1;
let grid: SheetGrid | undefined;
let selection: Selection | undefined;
const sheetCache = new Map<number, SheetData>();

const formulaBar = el('div', 'formula-bar');
const addressBox = el('div', 'address');
const valueBox = el('div', 'value');
formulaBar.append(addressBox, valueBox);
const gridHost = el('div', 'grid-host');
const footer = el('div', 'sheet-footer');
const tabs = el('div', 'sheet-tabs');
const notice = el('div', 'sheet-notice');
footer.append(tabs, notice);

onHostMessage((message) => {
  switch (message.type) {
    case 'error':
      showMessage(app, message.message, 'error');
      break;
    case 'workbook': {
      info = message.info;
      sheetCache.clear();
      const visible = info.sheets.filter((s) => s.state === 'visible');
      if (!visible.length) return showMessage(app, 'This workbook has no visible worksheets.', 'error');
      if (!visible.some((s) => s.index === activeSheet)) activeSheet = visible[0].index;
      renderTabs();
      app.replaceChildren(formulaBar, gridHost, footer);
      openSheet(activeSheet);
      break;
    }
    case 'sheet':
      sheetCache.set(message.sheet.index, message.sheet);
      if (message.sheet.index === activeSheet) showSheet(message.sheet);
      break;
  }
});

showMessage(app, 'Loading workbook…');
send({ type: 'ready' });

function renderTabs(): void {
  tabs.replaceChildren(
    ...(info?.sheets ?? [])
      .filter((s) => s.state === 'visible')
      .map((s) => {
        const tab = el('button', 'sheet-tab', s.name);
        tab.classList.toggle('active', s.index === activeSheet);
        tab.onclick = () => openSheet(s.index);
        return tab;
      }),
  );
}

function openSheet(index: number): void {
  saveView();
  activeSheet = index;
  patchState<State>({ sheet: index });
  renderTabs();
  const cached = sheetCache.get(index);
  if (cached) return showSheet(cached);
  grid = undefined;
  showMessage(gridHost, 'Loading sheet…');
  send({ type: 'loadSheet', index });
}

function showSheet(sheet: SheetData): void {
  grid = new SheetGrid(sheet, {
    onSelect(cell, selected) {
      selection = selected;
      addressBox.textContent = `${columnName(selected.c)}${selected.r + 1}`;
      valueBox.textContent = cell?.f ? `=${cell.f}` : (cell?.v ?? '');
      saveView();
    },
  });
  gridHost.replaceChildren(grid.element);
  const view = getState<State>().views?.[sheet.index];
  grid.restore(view, view?.selection);
  grid.element.addEventListener('scroll', saveView, true);
  notice.textContent = sheet.unsupported.length ? `Not shown: ${sheet.unsupported.join(', ')}` : '';
  notice.title = notice.textContent;
}

function saveView(): void {
  if (!grid || activeSheet < 0) return;
  const views = { ...(getState<State>().views ?? {}) };
  views[activeSheet] = { ...grid.scroll, selection };
  patchState<State>({ views });
}
