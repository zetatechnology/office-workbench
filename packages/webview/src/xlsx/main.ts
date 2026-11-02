import type { EditOp, SheetData, WorkbookInfo } from '@office-workbench/core';
import { columnName } from '@office-workbench/core/address';
import { el, getState, onHostMessage, patchState, send, showMessage } from '../common/host';
import { MAX_COLS, MAX_ROWS, SheetGrid, type Range, type Selection } from './grid';
import { showMenu, type MenuItem } from './menu';
import '../common/base.css';
import './xlsx.css';

interface State {
  sheet: number;
  views: Record<number, { top: number; left: number; selection?: Selection }>;
}

const app = document.getElementById('app')!;
let info: WorkbookInfo | undefined;
let editable = false;
let activeSheet = getState<State>().sheet ?? -1;
let grid: SheetGrid | undefined;
let selection: Selection | undefined;
const sheetCache = new Map<number, SheetData>();

const formulaBar = el('div', 'formula-bar');
const addressBox = el('div', 'address');
const valueBox = el('input', 'value');
valueBox.spellcheck = false;
valueBox.setAttribute('aria-label', 'Formula bar');
formulaBar.append(addressBox, valueBox);
const gridHost = el('div', 'grid-host');
const footer = el('div', 'sheet-footer');
const tabs = el('div', 'sheet-tabs');
const addTab = el('button', 'sheet-add', '+');
addTab.title = 'New sheet';
const notice = el('div', 'sheet-notice');
footer.append(tabs, addTab, notice);
const toast = el('div', 'toast');
toast.hidden = true;

onHostMessage((message) => {
  switch (message.type) {
    case 'error':
      showMessage(app, message.message, 'error');
      break;
    case 'notice':
      showToast(message.message);
      break;
    case 'workbook': {
      info = message.info;
      editable = message.editable;
      sheetCache.clear();
      const visible = info.sheets.filter((s) => s.state === 'visible');
      if (!visible.length) return showMessage(app, 'This workbook has no visible worksheets.', 'error');
      if (message.activate !== undefined) activeSheet = message.activate;
      if (!visible.some((s) => s.index === activeSheet)) {
        activeSheet = (visible.find((s) => s.index > activeSheet) ?? visible[visible.length - 1]).index;
      }
      addTab.hidden = !editable;
      valueBox.readOnly = !editable;
      if (!app.contains(gridHost)) app.replaceChildren(formulaBar, gridHost, footer, toast);
      renderTabs();
      openSheet(activeSheet, false);
      break;
    }
    case 'sheet':
      sheetCache.set(message.sheet.index, message.sheet);
      if (message.sheet.index === activeSheet) showSheet(message.sheet);
      break;
    case 'cells':
      // The grid keeps its own copy of the cells; drop the cached sheet so it is reloaded fresh.
      sheetCache.delete(message.sheet);
      if (message.sheet === activeSheet) grid?.applyPatch(message);
      break;
  }
});

showMessage(app, 'Loading workbook…');
send({ type: 'ready' });

function edit(op: EditOp): void {
  if (editable) send({ type: 'edit', op });
}

// --- Sheet tabs ------------------------------------------------------------------------------

function renderTabs(): void {
  tabs.replaceChildren(
    ...(info?.sheets ?? [])
      .filter((s) => s.state === 'visible')
      .map((s) => {
        const tab = el('button', 'sheet-tab', s.name);
        tab.classList.toggle('active', s.index === activeSheet);
        tab.onclick = () => openSheet(s.index);
        if (editable) {
          tab.ondblclick = () => renameTab(tab, s.index, s.name);
          tab.oncontextmenu = (e) => {
            e.preventDefault();
            showMenu(e.clientX, e.clientY, [
              { label: 'Insert Sheet', run: () => edit({ type: 'addSheet' }) },
              { label: 'Rename', run: () => renameTab(tab, s.index, s.name) },
              { label: 'Delete', run: () => edit({ type: 'deleteSheet', sheet: s.index }), disabled: visibleCount() < 2 },
            ]);
          };
        }
        return tab;
      }),
  );
}

function visibleCount(): number {
  return info?.sheets.filter((s) => s.state === 'visible').length ?? 0;
}

function renameTab(tab: HTMLElement, index: number, name: string): void {
  const input = el('input', 'sheet-rename');
  input.value = name;
  input.maxLength = 31;
  input.spellcheck = false;
  tab.replaceWith(input);
  input.focus();
  input.select();
  let done = false;
  const finish = (commit: boolean) => {
    if (done) return;
    done = true;
    const value = input.value.trim();
    input.replaceWith(tab);
    if (commit && value && value !== name) edit({ type: 'renameSheet', sheet: index, name: value });
    grid?.focus();
  };
  input.onkeydown = (e) => {
    if (e.key === 'Enter') finish(true);
    if (e.key === 'Escape') finish(false);
  };
  input.onblur = () => finish(true);
}

addTab.onclick = () => edit({ type: 'addSheet' });

function openSheet(index: number, keepView = true): void {
  if (keepView) saveView();
  activeSheet = index;
  patchState<State>({ sheet: index });
  renderTabs();
  const cached = sheetCache.get(index);
  if (cached) return showSheet(cached);
  if (!grid || !keepView) showMessage(gridHost, 'Loading sheet…');
  grid = undefined;
  send({ type: 'loadSheet', index });
}

// --- Grid ------------------------------------------------------------------------------------

function showSheet(sheet: SheetData): void {
  const hadFocus = gridHost.contains(document.activeElement);
  grid = new SheetGrid(sheet, {
    editable,
    onSelect(cell, active, range, single) {
      selection = active;
      addressBox.textContent = single ? `${columnName(active.c)}${active.r + 1}` : rangeLabel(range);
      if (document.activeElement !== valueBox) valueBox.value = cell?.input ?? cell?.v ?? '';
      saveView();
    },
    onEditorInput(value) {
      if (value !== undefined) valueBox.value = value;
    },
    onEdit(cells) {
      edit({ type: 'setCells', sheet: sheet.index, cells });
    },
    onClear(range) {
      edit({ type: 'clearRange', sheet: sheet.index, range });
    },
    onContextMenu(e, kind, range) {
      showMenu(e.clientX, e.clientY, contextItems(sheet.index, kind, range));
    },
  });
  gridHost.replaceChildren(grid.element);
  const view = getState<State>().views?.[sheet.index];
  grid.restore(view, view?.selection);
  grid.element.addEventListener('scroll', saveView, true);
  notice.textContent = sheet.unsupported.length ? `Not shown: ${sheet.unsupported.join(', ')}` : '';
  notice.title = notice.textContent;
  if (hadFocus || document.activeElement === document.body) grid.focus();
}

function rangeLabel(range: Range): string {
  if (range.c1 === 0 && range.c2 === MAX_COLS - 1) return `${range.r1 + 1}:${range.r2 + 1}`;
  if (range.r1 === 0 && range.r2 === MAX_ROWS - 1) return `${columnName(range.c1)}:${columnName(range.c2)}`;
  return `${range.r2 - range.r1 + 1}R × ${range.c2 - range.c1 + 1}C`;
}

function contextItems(sheet: number, kind: 'cell' | 'row' | 'col', range: Range): (MenuItem | null)[] {
  const rows = range.r2 - range.r1 + 1, cols = range.c2 - range.c1 + 1;
  const copy = (cut: boolean) => () => {
    grid?.focus();
    document.execCommand(cut ? 'cut' : 'copy');
  };
  const items: (MenuItem | null)[] = [
    { label: 'Cut', run: copy(true), disabled: !editable },
    { label: 'Copy', run: copy(false) },
    null,
  ];
  if (!editable) return items.slice(0, 2);
  const insertRows: MenuItem = { label: rows > 1 ? `Insert ${rows} Rows Above` : 'Insert Row Above', run: () => edit({ type: 'insertRows', sheet, at: range.r1, count: rows }) };
  const deleteRows: MenuItem = { label: rows > 1 ? `Delete ${rows} Rows` : 'Delete Row', run: () => edit({ type: 'deleteRows', sheet, at: range.r1, count: rows }) };
  const insertCols: MenuItem = { label: cols > 1 ? `Insert ${cols} Columns Left` : 'Insert Column Left', run: () => edit({ type: 'insertCols', sheet, at: range.c1, count: cols }) };
  const deleteCols: MenuItem = { label: cols > 1 ? `Delete ${cols} Columns` : 'Delete Column', run: () => edit({ type: 'deleteCols', sheet, at: range.c1, count: cols }) };
  if (kind === 'row') items.push(insertRows, deleteRows);
  else if (kind === 'col') items.push(insertCols, deleteCols);
  else items.push(insertRows, insertCols, null, deleteRows, deleteCols);
  items.push(null, { label: 'Clear Contents', run: () => edit({ type: 'clearRange', sheet, range }) });
  return items;
}

// --- Formula bar -----------------------------------------------------------------------------

let formulaOriginal = '';
valueBox.onfocus = () => {
  formulaOriginal = valueBox.value;
};
valueBox.onkeydown = (e) => {
  if (e.isComposing || e.keyCode === 229) return;
  if (e.key === 'Enter' || e.key === 'Tab') {
    e.preventDefault();
    if (selection && valueBox.value !== formulaOriginal) {
      edit({ type: 'setCells', sheet: activeSheet, cells: [{ ...selection, input: valueBox.value }] });
    }
    formulaOriginal = valueBox.value;
    grid?.focus();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    valueBox.value = formulaOriginal;
    grid?.focus();
  }
};

// --- Misc ------------------------------------------------------------------------------------

let toastTimer: ReturnType<typeof setTimeout> | undefined;
function showToast(text: string): void {
  toast.textContent = text;
  toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (toast.hidden = true), 5000);
}

function saveView(): void {
  if (!grid || activeSheet < 0) return;
  const views = { ...(getState<State>().views ?? {}) };
  views[activeSheet] = { ...grid.scroll, selection };
  patchState<State>({ views });
}
