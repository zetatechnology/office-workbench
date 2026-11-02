// Namespace-agnostic helpers over @xmldom/xmldom. OOXML parts use stable prefixes in practice,
// but matching on localName keeps us robust against files written with unusual prefixes.
import { DOMParser, type Document, type Element, type Node } from '@xmldom/xmldom';

export type { Document, Element };

const ELEMENT_NODE = 1;

export function parseXml(text: string): Document {
  // xmldom never resolves external entities, so XXE is not a concern here.
  return new DOMParser({ onError: () => undefined }).parseFromString(text, 'application/xml');
}

function isElement(node: Node): node is Element {
  return node.nodeType === ELEMENT_NODE;
}

export function children(el: Element | null | undefined, localName?: string): Element[] {
  const result: Element[] = [];
  if (!el) return result;
  for (let node = el.firstChild; node; node = node.nextSibling) {
    if (isElement(node) && (localName === undefined || node.localName === localName)) result.push(node);
  }
  return result;
}

export function child(el: Element | null | undefined, localName: string): Element | undefined {
  if (!el) return undefined;
  for (let node = el.firstChild; node; node = node.nextSibling) {
    if (isElement(node) && node.localName === localName) return node;
  }
  return undefined;
}

/** Walks down a chain of child elements, e.g. `path(sp, 'spPr', 'xfrm', 'off')`. */
export function path(el: Element | null | undefined, ...names: string[]): Element | undefined {
  let current: Element | undefined = el ?? undefined;
  for (const name of names) {
    current = child(current, name);
    if (!current) return undefined;
  }
  return current;
}

export function descendants(el: Element | Document | null | undefined, localName: string): Element[] {
  if (!el) return [];
  return Array.from(el.getElementsByTagNameNS('*', localName));
}

/** Reads an attribute by local name, ignoring its prefix (`r:id` matches `id`). */
export function attr(el: Element | null | undefined, localName: string): string | undefined {
  if (!el) return undefined;
  const direct = el.getAttribute(localName);
  if (direct !== null && direct !== '') return direct;
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (a && a.localName === localName) return a.value;
  }
  return direct ?? undefined;
}

const RELATIONSHIPS_NS = /officeDocument\/2006\/relationships$|officeDocument\/relationships$/;

/** Reads a relationship reference such as `r:id` or `r:embed` (distinct from an unprefixed `id`). */
export function relAttr(el: Element | null | undefined, localName: string): string | undefined {
  if (!el) return undefined;
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (a && a.localName === localName && a.namespaceURI && RELATIONSHIPS_NS.test(a.namespaceURI)) return a.value;
  }
  return undefined;
}

export function numAttr(el: Element | null | undefined, localName: string): number | undefined {
  const value = attr(el, localName);
  if (value === undefined || value === '') return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
}

/** OOXML booleans: absent → fallback, "0"/"false"/"off" → false, anything else → true. */
export function boolAttr(el: Element | null | undefined, localName: string, fallback = false): boolean {
  const value = attr(el, localName);
  if (value === undefined) return fallback;
  return !(value === '0' || value === 'false' || value === 'off');
}

/** A toggle element such as `<b/>` or `<b val="0"/>`. */
export function toggle(el: Element | undefined): boolean | undefined {
  if (!el) return undefined;
  return boolAttr(el, 'val', true);
}

export function text(el: Element | null | undefined): string {
  return el?.textContent ?? '';
}
