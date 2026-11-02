// Renders the resolved slide model into absolutely positioned HTML/SVG at native slide size.
import type {
  Box, Fill, Line, Paragraph, PictureElement, Presentation, ShapeElement, Slide, SlideElement, TableElement, TextBody,
} from '@office-workbench/core';
import { el } from '../common/host';

const SVG_NS = 'http://www.w3.org/2000/svg';
const FALLBACK_FONTS = '"Yu Gothic", "Meiryo", "Segoe UI", sans-serif';
/** PowerPoint's single line spacing is roughly 1.2× the font size. */
const SINGLE_LINE = 1.2;

export type MediaUrls = Map<string, string>;

export function renderSlide(presentation: Presentation, slide: Slide, media: MediaUrls): HTMLElement {
  const root = el('div', 'slide');
  root.style.width = `${presentation.width}px`;
  root.style.height = `${presentation.height}px`;
  applyFill(root, slide.background, media);
  for (const element of slide.elements) root.append(renderElement(element, media));
  return root;
}

function renderElement(element: SlideElement, media: MediaUrls): HTMLElement {
  switch (element.type) {
    case 'shape':
      return renderShape(element, media);
    case 'picture':
      return renderPicture(element, media);
    case 'table':
      return renderTable(element, media);
    case 'unsupported': {
      const node = positioned('unsupported', element.box);
      node.append(el('span', undefined, element.label));
      return node;
    }
  }
}

function positioned(className: string, box: Box): HTMLElement {
  const node = el('div', `el ${className}`);
  Object.assign(node.style, {
    left: `${box.x}px`,
    top: `${box.y}px`,
    width: `${Math.max(box.w, 0)}px`,
    height: `${Math.max(box.h, 0)}px`,
  });
  if (box.rot) node.style.transform = `rotate(${box.rot}deg)`;
  return node;
}

function flipTransform(box: Box): string | undefined {
  if (!box.flipH && !box.flipV) return undefined;
  return `scale(${box.flipH ? -1 : 1}, ${box.flipV ? -1 : 1})`;
}

function renderShape(shape: ShapeElement, media: MediaUrls): HTMLElement {
  const node = positioned('shape', shape.box);
  const { w, h } = shape.box;
  const flip = flipTransform(shape.box);

  // Gradient and picture fills are painted with CSS, clipped to the geometry.
  if (shape.path && (shape.fill.type === 'gradient' || shape.fill.type === 'image')) {
    const layer = el('div', 'fill-layer');
    applyFill(layer, shape.fill, media);
    layer.style.clipPath = `path('${shape.path}')`;
    if (flip) layer.style.transform = flip;
    node.append(layer);
  }

  if (shape.path && (shape.fill.type === 'solid' || shape.line)) {
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('width', String(Math.max(w, 1)));
    svg.setAttribute('height', String(Math.max(h, 1)));
    svg.setAttribute('viewBox', `0 0 ${Math.max(w, 1)} ${Math.max(h, 1)}`);
    svg.classList.add('geometry');
    if (flip) svg.style.transform = flip;
    const pathEl = document.createElementNS(SVG_NS, 'path');
    pathEl.setAttribute('d', shape.path);
    pathEl.setAttribute('fill', shape.fill.type === 'solid' && shape.closed ? shape.fill.color : 'none');
    if (shape.line) applyStroke(svg, pathEl, shape.line);
    svg.append(pathEl);
    node.append(svg);
  } else if (!shape.path && shape.fill.type !== 'none') {
    applyFill(node, shape.fill, media);
  }

  if (shape.text) node.append(renderText(shape.text));
  return node;
}

let markerId = 0;

function applyStroke(svg: SVGSVGElement, pathEl: SVGPathElement, line: Line): void {
  pathEl.setAttribute('stroke', line.color);
  pathEl.setAttribute('stroke-width', String(line.width));
  pathEl.setAttribute('stroke-linejoin', 'round');
  if (line.dash) pathEl.setAttribute('stroke-dasharray', line.dash);
  const defs = document.createElementNS(SVG_NS, 'defs');
  for (const [end, attr] of [[line.headEnd, 'marker-start'], [line.tailEnd, 'marker-end']] as const) {
    if (!end) continue;
    const id = `m${++markerId}`;
    const marker = document.createElementNS(SVG_NS, 'marker');
    marker.setAttribute('id', id);
    marker.setAttribute('viewBox', '0 0 10 10');
    marker.setAttribute('refX', '5');
    marker.setAttribute('refY', '5');
    marker.setAttribute('markerWidth', '4');
    marker.setAttribute('markerHeight', '4');
    marker.setAttribute('orient', 'auto-start-reverse');
    const shapeEl = document.createElementNS(SVG_NS, end === 'oval' ? 'circle' : 'path');
    if (end === 'oval') {
      shapeEl.setAttribute('cx', '5');
      shapeEl.setAttribute('cy', '5');
      shapeEl.setAttribute('r', '4');
    } else {
      shapeEl.setAttribute('d', end === 'diamond' ? 'M5 0 L10 5 L5 10 L0 5 Z' : end === 'arrow' ? 'M0 0 L10 5 L0 10' : 'M0 0 L10 5 L0 10 Z');
    }
    shapeEl.setAttribute('fill', end === 'arrow' ? 'none' : line.color);
    if (end === 'arrow') {
      shapeEl.setAttribute('stroke', line.color);
      shapeEl.setAttribute('stroke-width', '1.5');
    }
    marker.append(shapeEl);
    defs.append(marker);
    pathEl.setAttribute(attr, `url(#${id})`);
  }
  if (defs.childNodes.length) svg.append(defs);
}

function renderPicture(pic: PictureElement, media: MediaUrls): HTMLElement {
  const node = positioned('picture', pic.box);
  const url = pic.src ? media.get(pic.src) : undefined;
  if (!url) {
    node.classList.add('unsupported');
    node.append(el('span', undefined, 'Image'));
    return node;
  }
  const img = el('img');
  img.src = url;
  img.alt = '';
  img.draggable = false;
  const crop = pic.crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const iw = pic.box.w / Math.max(1 - crop.l - crop.r, 0.01);
  const ih = pic.box.h / Math.max(1 - crop.t - crop.b, 0.01);
  Object.assign(img.style, { width: `${iw}px`, height: `${ih}px`, left: `${-crop.l * iw}px`, top: `${-crop.t * ih}px` });
  const flip = flipTransform(pic.box);
  if (flip) img.style.transform = flip;
  node.append(img);
  if (pic.line) node.style.outline = `${pic.line.width}px solid ${pic.line.color}`;
  return node;
}

function renderTable(table: TableElement, media: MediaUrls): HTMLElement {
  const node = positioned('table', table.box);
  node.style.height = 'auto';
  const t = el('table');
  const colgroup = el('colgroup');
  for (const w of table.columns) {
    const col = el('col');
    col.style.width = `${w}px`;
    colgroup.append(col);
  }
  t.append(colgroup);
  t.style.width = `${table.columns.reduce((a, b) => a + b, 0)}px`;
  for (const row of table.rows) {
    const tr = el('tr');
    tr.style.height = `${row.height}px`;
    for (const cell of row.cells) {
      if (cell.merged) continue;
      const td = el('td');
      if (cell.colSpan > 1) td.colSpan = cell.colSpan;
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      applyFill(td, cell.fill, media);
      const border = (line: Line | undefined) => (line ? `${line.width}px solid ${line.color}` : 'none');
      td.style.borderLeft = border(cell.borders.l);
      td.style.borderRight = border(cell.borders.r);
      td.style.borderTop = border(cell.borders.t);
      td.style.borderBottom = border(cell.borders.b);
      td.style.verticalAlign = cell.text.anchor === 'ctr' ? 'middle' : cell.text.anchor === 'b' ? 'bottom' : 'top';
      const { l, t: top, r, b } = cell.text.insets;
      td.style.padding = `${top}px ${r}px ${b}px ${l}px`;
      td.append(...renderParagraphs(cell.text));
      tr.append(td);
    }
    t.append(tr);
  }
  node.append(t);
  return node;
}

function renderText(body: TextBody): HTMLElement {
  const node = el('div', 'text');
  node.style.padding = `${body.insets.t}px ${body.insets.r}px ${body.insets.b}px ${body.insets.l}px`;
  node.style.justifyContent = body.anchor === 'ctr' ? 'center' : body.anchor === 'b' ? 'flex-end' : 'flex-start';
  if (!body.wrap) node.classList.add('nowrap');
  if (body.vertical) {
    node.style.writingMode = 'vertical-rl';
    if (body.vertical === 'vert270') node.style.transform = 'rotate(180deg)';
  }
  node.append(...renderParagraphs(body));
  return node;
}

function renderParagraphs(body: TextBody): HTMLElement[] {
  return body.paragraphs.map((p) => renderParagraph(p));
}

function renderParagraph(p: Paragraph): HTMLElement {
  const node = el('div', 'para');
  node.style.textAlign = p.align === 'ctr' ? 'center' : p.align === 'r' ? 'right' : p.align === 'just' || p.align === 'dist' ? 'justify' : 'left';
  node.style.paddingLeft = `${p.marginLeft}px`;
  node.style.textIndent = `${p.indent}px`;
  if (p.spaceBefore) node.style.marginTop = `${p.spaceBefore}px`;
  if (p.spaceAfter) node.style.marginBottom = `${p.spaceAfter}px`;
  const firstSize = p.runs.find((r) => !r.br)?.size ?? p.endSize;
  node.style.fontSize = `${firstSize}px`;
  node.style.lineHeight = p.lineSpacing
    ? 'px' in p.lineSpacing
      ? `${p.lineSpacing.px}px`
      : String(SINGLE_LINE * p.lineSpacing.percent)
    : String(SINGLE_LINE);

  if (p.bullet) {
    const first = p.runs.find((r) => r.text);
    const bullet = el('span', 'bullet', p.bullet);
    if (first) {
      bullet.style.color = first.color ?? '';
      bullet.style.fontSize = `${first.size}px`;
    }
    // A hanging indent: the bullet sits in the negative indent before the text.
    bullet.style.display = 'inline-block';
    bullet.style.minWidth = `${Math.max(-p.indent, 0)}px`;
    bullet.style.textIndent = '0';
    node.append(bullet);
  }

  if (!p.runs.length) {
    node.append(document.createTextNode('​'));
    node.style.fontSize = `${p.endSize}px`;
  }
  for (const run of p.runs) {
    if (run.br) {
      node.append(el('br'));
      continue;
    }
    const span = el('span', undefined, run.text);
    span.style.fontSize = `${run.size}px`;
    span.style.fontFamily = run.font ? `${run.font}, ${FALLBACK_FONTS}` : FALLBACK_FONTS;
    if (run.color) span.style.color = run.color;
    if (run.bold) span.style.fontWeight = 'bold';
    if (run.italic) span.style.fontStyle = 'italic';
    const deco = [run.underline && 'underline', run.strike && 'line-through'].filter(Boolean).join(' ');
    if (deco) span.style.textDecoration = deco;
    if (run.baseline) {
      span.style.verticalAlign = `${run.baseline}%`;
      span.style.fontSize = `${run.size * 0.7}px`;
    }
    node.append(span);
  }
  return node;
}

function applyFill(node: HTMLElement, fill: Fill, media: MediaUrls): void {
  switch (fill.type) {
    case 'solid':
      node.style.background = fill.color;
      break;
    case 'gradient':
      node.style.background = fill.css;
      break;
    case 'image': {
      const url = media.get(fill.src);
      if (url) {
        node.style.backgroundImage = `url("${url}")`;
        node.style.backgroundSize = '100% 100%';
      }
      break;
    }
  }
}
