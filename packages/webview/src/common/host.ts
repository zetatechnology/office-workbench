// Bridge to the extension host plus small DOM helpers shared by the viewers.
import type { HostMessage, WebviewMessage } from '@office-workbench/core';

interface VsCodeApi {
  postMessage(message: WebviewMessage): void;
  getState(): unknown;
  setState(state: unknown): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

const api = acquireVsCodeApi();

export function send(message: WebviewMessage): void {
  api.postMessage(message);
}

export function onHostMessage(handler: (message: HostMessage) => void): void {
  window.addEventListener('message', (event: MessageEvent<HostMessage>) => handler(event.data));
}

/** Per-editor UI state that survives the webview being hidden and recreated. */
export function getState<T extends object>(): Partial<T> {
  return (api.getState() as Partial<T> | undefined) ?? {};
}

export function patchState<T extends object>(patch: Partial<T>): void {
  api.setState({ ...getState<T>(), ...patch });
}

export function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  textContent?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (textContent !== undefined) node.textContent = textContent;
  return node;
}

export function showMessage(root: HTMLElement, text: string, kind: 'loading' | 'error' = 'loading'): void {
  root.replaceChildren(el('div', `message message-${kind}`, text));
}

export function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (ch) => `&#${ch.charCodeAt(0)};`);
}
