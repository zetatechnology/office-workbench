// A short message in the corner (refused edits and similar), hidden after a few seconds.
import { el } from './host';
import './popup.css';

const toast = el('div', 'toast');
toast.hidden = true;
let timer: ReturnType<typeof setTimeout> | undefined;

export function showToast(text: string): void {
  if (!toast.isConnected) document.body.append(toast);
  toast.textContent = text;
  toast.hidden = false;
  clearTimeout(timer);
  timer = setTimeout(() => (toast.hidden = true), 5000);
}
