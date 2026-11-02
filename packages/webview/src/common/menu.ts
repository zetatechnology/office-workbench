// A minimal context menu styled after VS Code's, positioned at the pointer.
import { el } from './host';
import './popup.css';

export interface MenuItem {
  label: string;
  run(): void;
  disabled?: boolean;
}

let open: HTMLElement | undefined;

export function closeMenu(): void {
  open?.remove();
  open = undefined;
}

/** `items` entries of `null` render as separators. */
export function showMenu(x: number, y: number, items: (MenuItem | null)[]): void {
  closeMenu();
  const menu = el('div', 'context-menu');
  menu.setAttribute('role', 'menu');
  for (const item of items) {
    if (!item) {
      menu.append(el('div', 'menu-separator'));
      continue;
    }
    const button = el('button', 'menu-item', item.label);
    button.setAttribute('role', 'menuitem');
    button.disabled = !!item.disabled;
    button.onmousedown = (e) => e.preventDefault();
    button.onclick = () => {
      closeMenu();
      item.run();
    };
    menu.append(button);
  }
  document.body.append(menu);
  const { innerWidth, innerHeight } = window;
  const rect = menu.getBoundingClientRect();
  menu.style.left = `${Math.max(0, Math.min(x, innerWidth - rect.width - 4))}px`;
  menu.style.top = `${Math.max(0, Math.min(y, innerHeight - rect.height - 4))}px`;
  open = menu;
}

window.addEventListener('mousedown', (e) => {
  if (open && !open.contains(e.target as Node)) closeMenu();
}, true);
window.addEventListener('keydown', (e) => {
  if (open && e.key === 'Escape') closeMenu();
}, true);
window.addEventListener('blur', closeMenu);
