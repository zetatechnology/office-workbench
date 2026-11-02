// PresentationML reader: resolves slides against their layouts, masters and theme into a flat,
// render-ready model (absolute boxes, resolved colors, inherited text styles).
import { applyDrawingModifiers, hexToRgb, parseTheme, presetColor, rgbToCss, type Rgb, type Theme } from '../color';
import { imageContentType, isBrowserImage, OoxmlPackage, type Relationship } from '../package';
import { attr, boolAttr, child, children, numAttr, path, relAttr, text, type Element } from '../xml';
import { customPath, presetPath, readAdjustments } from './geometry';
import type {
  Box, Fill, Line, MediaRef, Paragraph, PictureElement, Presentation, Run, ShapeElement, Slide, SlideElement,
  TableCell, TableElement, TextBody,
} from './types';

const EMU_PER_PX = 9525;
const PX_PER_PT = 4 / 3;
const px = (emu: number | undefined) => (emu ?? 0) / EMU_PER_PX;

export interface ParsedPresentation {
  presentation: Presentation;
  /** Image bytes keyed by part path (the `src` of pictures and image fills). */
  media: Map<string, Uint8Array>;
}

/** One part in the slide → layout → master chain. */
interface PartContext {
  path: string;
  root: Element;
  rels: Map<string, Relationship>;
}

interface SlideContext {
  pkg: OoxmlPackage;
  theme: Theme;
  clrMap: Record<string, string>;
  slide: PartContext;
  layout?: PartContext;
  master?: PartContext;
  masterTextStyles?: Element;
  defaultTextStyle?: Element;
  tableStyles?: Element;
  media: Map<string, MediaRef>;
  slideNumber: number;
}

/** Affine mapping applied to boxes inside groups. */
interface Transform {
  sx: number;
  sy: number;
  tx: number;
  ty: number;
}

const IDENTITY: Transform = { sx: 1, sy: 1, tx: 0, ty: 0 };

/** Inherited properties of a placeholder shape, from its layout and master counterparts. */
interface Inherited {
  /** Placeholder shapes in priority order: layout first, then master. */
  sources: Element[];
  category: 'title' | 'body' | 'other';
}

export async function parsePresentation(data: Uint8Array): Promise<ParsedPresentation> {
  return readPresentation(await OoxmlPackage.open(data));
}

/** Slide parts in presentation order, with their `<p:sldId>` entries. */
export async function slideParts(pkg: OoxmlPackage, presPath: string): Promise<{ sldId: Element; rId: string; path: string }[]> {
  const root = (await pkg.xml(presPath))?.documentElement;
  const rels = await pkg.rels(presPath);
  return children(child(root, 'sldIdLst'), 'sldId').flatMap((sldId) => {
    const rId = relAttr(sldId, 'id') ?? '';
    const rel = rels.get(rId);
    return rel && !rel.external ? [{ sldId, rId, path: rel.target }] : [];
  });
}

/**
 * Reads the presentation as it is in the package now (including unsaved edits). Without
 * `withMedia`, image bytes are not read (the media list is still complete).
 */
export async function readPresentation(pkg: OoxmlPackage, withMedia = true): Promise<ParsedPresentation> {
  const presPath = await pkg.mainPart();
  const root = (await pkg.xml(presPath))?.documentElement;
  if (!root || root.localName !== 'presentation') throw new Error('This file is not a PowerPoint presentation.');
  const rels = await pkg.rels(presPath);

  const size = child(root, 'sldSz');
  const width = px(numAttr(size, 'cx') ?? 9_144_000);
  const height = px(numAttr(size, 'cy') ?? 6_858_000);
  const defaultTextStyle = child(root, 'defaultTextStyle');
  const tableStylesPath = [...rels.values()].find((r) => r.type === 'tableStyles')?.target;
  const tableStyles = (tableStylesPath ? (await pkg.xml(tableStylesPath))?.documentElement : undefined) ?? undefined;

  const media = new Map<string, MediaRef>();
  const slides: Slide[] = [];
  for (const { path: slidePath, sldId } of await slideParts(pkg, presPath)) {
    const slide = await parseSlide(pkg, slidePath, slides.length, { defaultTextStyle, tableStyles, media });
    slides.push({ ...slide, id: attr(sldId, 'id') ?? slidePath });
  }

  const bytes = new Map<string, Uint8Array>();
  for (const ref of withMedia ? media.values() : []) {
    const b = await pkg.bytes(ref.path);
    if (b) bytes.set(ref.path, b);
  }
  return { presentation: { width, height, slides, media: [...media.values()] }, media: bytes };
}

async function loadPart(pkg: OoxmlPackage, partPath: string): Promise<PartContext | undefined> {
  const root = (await pkg.xml(partPath))?.documentElement;
  if (!root) return undefined;
  return { path: partPath, root, rels: await pkg.rels(partPath) };
}

async function parseSlide(
  pkg: OoxmlPackage,
  slidePath: string,
  index: number,
  shared: { defaultTextStyle?: Element; tableStyles?: Element; media: Map<string, MediaRef> },
): Promise<Omit<Slide, 'id'>> {
  const slide = await loadPart(pkg, slidePath);
  if (!slide) throw new Error(`Slide ${index + 1} could not be read.`);
  const layoutRel = [...slide.rels.values()].find((r) => r.type === 'slideLayout');
  const layout = layoutRel ? await loadPart(pkg, layoutRel.target) : undefined;
  const masterRel = layout && [...layout.rels.values()].find((r) => r.type === 'slideMaster');
  const master = masterRel ? await loadPart(pkg, masterRel.target) : undefined;
  const themeRel = master && [...master.rels.values()].find((r) => r.type === 'theme');
  const theme = parseTheme(themeRel ? await pkg.xml(themeRel.target) : undefined);

  const clrMap: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
  const mapEl = child(master?.root, 'clrMap');
  for (const key of ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) {
    const v = attr(mapEl, key);
    if (v) clrMap[key] = v;
  }

  const ctx: SlideContext = {
    pkg,
    theme,
    clrMap,
    slide,
    layout,
    master,
    masterTextStyles: child(master?.root, 'txStyles'),
    defaultTextStyle: shared.defaultTextStyle,
    tableStyles: shared.tableStyles,
    media: shared.media,
    slideNumber: index + 1,
  };

  const elements: SlideElement[] = [];
  const showMasterShapes = boolAttr(slide.root, 'showMasterSp', true);
  if (showMasterShapes && master && boolAttr(layout?.root, 'showMasterSp', true)) {
    walkTree(ctx, master, children(path(master.root, 'cSld', 'spTree')), IDENTITY, elements, true, false);
  }
  if (showMasterShapes && layout) walkTree(ctx, layout, children(path(layout.root, 'cSld', 'spTree')), IDENTITY, elements, true, false);
  walkTree(ctx, slide, children(path(slide.root, 'cSld', 'spTree')), IDENTITY, elements, false, true);

  return {
    index,
    hidden: !boolAttr(slide.root, 'show', true),
    background: resolveBackground(ctx),
    elements,
    notes: await readNotes(pkg, slide),
  };
}

async function readNotes(pkg: OoxmlPackage, slide: PartContext): Promise<string | undefined> {
  const rel = [...slide.rels.values()].find((r) => r.type === 'notesSlide');
  const notes = rel ? await loadPart(pkg, rel.target) : undefined;
  if (!notes) return undefined;
  for (const sp of children(path(notes.root, 'cSld', 'spTree'), 'sp')) {
    if (attr(path(sp, 'nvSpPr', 'nvPr', 'ph'), 'type') !== 'body') continue;
    const value = children(child(sp, 'txBody'), 'p')
      .map((p) => children(p).map((r) => (r.localName === 'br' ? '\n' : text(child(r, 't')))).join(''))
      .join('\n')
      .trim();
    return value || undefined;
  }
  return undefined;
}

// ---------------------------------------------------------------------------------------------
// Shape tree

function walkTree(
  ctx: SlideContext,
  part: PartContext,
  elements: Element[],
  transform: Transform,
  out: SlideElement[],
  skipPlaceholders: boolean,
  /** Shapes of the slide itself, outside alternate content: their text can be edited. */
  editable: boolean,
): void {
  for (const node of elements) {
    if (node.localName === 'AlternateContent') {
      // Prefer the fallback representation, which uses only baseline features. Editing it would
      // leave the other branch stale, so its shapes are not editable.
      const branch = child(node, 'Fallback') ?? child(node, 'Choice');
      walkTree(ctx, part, children(branch), transform, out, skipPlaceholders, false);
      continue;
    }
    if (skipPlaceholders && placeholderOf(node)) continue;
    switch (node.localName) {
      case 'sp':
      case 'cxnSp': {
        const shape = parseShape(ctx, part, node, transform, editable);
        if (shape) out.push(shape);
        break;
      }
      case 'pic': {
        const pic = parsePicture(ctx, part, node, transform);
        if (pic) out.push(pic);
        break;
      }
      case 'graphicFrame': {
        const frame = parseGraphicFrame(ctx, part, node, transform);
        if (frame?.type === 'table' && editable) frame.textId = attr(path(node, 'nvGraphicFramePr', 'cNvPr'), 'id');
        if (frame) out.push(frame);
        break;
      }
      case 'grpSp': {
        const xfrm = path(node, 'grpSpPr', 'xfrm');
        const off = child(xfrm, 'off'), ext = child(xfrm, 'ext');
        const chOff = child(xfrm, 'chOff'), chExt = child(xfrm, 'chExt');
        const cw = numAttr(chExt, 'cx') || numAttr(ext, 'cx') || 1;
        const ch = numAttr(chExt, 'cy') || numAttr(ext, 'cy') || 1;
        const sx = (numAttr(ext, 'cx') ?? cw) / cw;
        const sy = (numAttr(ext, 'cy') ?? ch) / ch;
        const local: Transform = {
          sx,
          sy,
          tx: px(numAttr(off, 'x')) - px(numAttr(chOff, 'x')) * sx,
          ty: px(numAttr(off, 'y')) - px(numAttr(chOff, 'y')) * sy,
        };
        const composed: Transform = {
          sx: transform.sx * local.sx,
          sy: transform.sy * local.sy,
          tx: transform.sx * local.tx + transform.tx,
          ty: transform.sy * local.ty + transform.ty,
        };
        walkTree(ctx, part, children(node), composed, out, skipPlaceholders, editable);
        break;
      }
    }
  }
}

function placeholderOf(el: Element): Element | undefined {
  const nv = child(el, 'nvSpPr') ?? child(el, 'nvPicPr') ?? child(el, 'nvGraphicFramePr') ?? child(el, 'nvCxnSpPr');
  return path(nv, 'nvPr', 'ph');
}

function findPlaceholder(part: PartContext | undefined, ph: Element, byTypeOnly = false): Element | undefined {
  if (!part) return undefined;
  const type = attr(ph, 'type') ?? 'body';
  const idx = attr(ph, 'idx');
  const shapes = children(path(part.root, 'cSld', 'spTree')).filter((s) => placeholderOf(s));
  const normalize = (t: string) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
  if (!byTypeOnly && idx !== undefined) {
    const byIdx = shapes.find((s) => attr(placeholderOf(s), 'idx') === idx);
    if (byIdx) return byIdx;
  }
  return (
    shapes.find((s) => (attr(placeholderOf(s), 'type') ?? 'body') === type) ??
    shapes.find((s) => normalize(attr(placeholderOf(s), 'type') ?? 'body') === normalize(type))
  );
}

function inheritance(ctx: SlideContext, part: PartContext, el: Element): Inherited {
  const ph = placeholderOf(el);
  if (!ph) return { sources: [], category: 'other' };
  const type = attr(ph, 'type') ?? 'body';
  const category = type === 'title' || type === 'ctrTitle' ? 'title' : ['body', 'subTitle', 'obj'].includes(type) ? 'body' : 'other';
  const sources: Element[] = [];
  if (part === ctx.slide) {
    const layoutPh = findPlaceholder(ctx.layout, ph);
    if (layoutPh) sources.push(layoutPh);
    const masterPh = findPlaceholder(ctx.master, layoutPh ? placeholderOf(layoutPh) ?? ph : ph, true);
    if (masterPh) sources.push(masterPh);
  } else if (part === ctx.layout) {
    const masterPh = findPlaceholder(ctx.master, ph, true);
    if (masterPh) sources.push(masterPh);
  }
  return { sources, category };
}

function readBox(xfrm: Element | undefined, transform: Transform): Box | undefined {
  const off = child(xfrm, 'off');
  const ext = child(xfrm, 'ext');
  if (!off || !ext) return undefined;
  const box: Box = {
    x: px(numAttr(off, 'x')) * transform.sx + transform.tx,
    y: px(numAttr(off, 'y')) * transform.sy + transform.ty,
    w: px(numAttr(ext, 'cx')) * transform.sx,
    h: px(numAttr(ext, 'cy')) * transform.sy,
  };
  const rot = numAttr(xfrm, 'rot');
  if (rot) box.rot = rot / 60_000;
  if (boolAttr(xfrm, 'flipH')) box.flipH = true;
  if (boolAttr(xfrm, 'flipV')) box.flipV = true;
  return box;
}

function shapeProps(el: Element): Element | undefined {
  return child(el, 'spPr') ?? child(el, 'grpSpPr');
}

/** Placeholders for pictures, charts, tables and the like, which hold no text. */
export const NON_TEXT_PLACEHOLDERS = new Set(['pic', 'chart', 'tbl', 'clipArt', 'dgm', 'media', 'sldImg']);

function parseShape(ctx: SlideContext, part: PartContext, el: Element, transform: Transform, editable: boolean): ShapeElement | undefined {
  const inherited = inheritance(ctx, part, el);
  const spPrChain = [shapeProps(el), ...inherited.sources.map(shapeProps)];
  const xfrm = spPrChain.map((s) => child(s, 'xfrm')).find((x) => x && child(x, 'off'));
  const box = readBox(xfrm, transform);
  if (!box) return undefined;

  const style = child(el, 'style') ?? inherited.sources.map((s) => child(s, 'style')).find(Boolean);
  let geometry: { d: string; closed: boolean } | undefined;
  const geomSource = spPrChain.find((s) => child(s, 'prstGeom') || child(s, 'custGeom'));
  const prstGeom = child(geomSource, 'prstGeom');
  const custGeom = child(geomSource, 'custGeom');
  if (custGeom) geometry = customPath(custGeom, box.w, box.h);
  else if (prstGeom) {
    const prst = attr(prstGeom, 'prst') ?? 'rect';
    geometry = presetPath(prst, box.w, box.h, readAdjustments(prstGeom)) ?? presetPath('rect', box.w, box.h);
  }

  let fill: Fill | undefined;
  for (const s of spPrChain) {
    fill = readFill(ctx, part, s);
    if (fill) break;
  }
  if (!fill) fill = styleFill(ctx, style);
  const lineEl = spPrChain.map((s) => child(s, 'ln')).find(Boolean);
  const line = readLine(ctx, lineEl, style);

  const txBody = child(el, 'txBody');
  const phType = attr(placeholderOf(el), 'type');
  const textId = editable && el.localName === 'sp' && !NON_TEXT_PLACEHOLDERS.has(phType ?? '') ? attr(path(el, 'nvSpPr', 'cNvPr'), 'id') : undefined;
  const textBody = txBody || textId ? readTextBody(ctx, el, txBody, inherited, style) : undefined;
  const isLine = el.localName === 'cxnSp' || geometry?.closed === false;
  const shape: ShapeElement = {
    type: 'shape',
    box,
    path: geometry?.d,
    closed: geometry?.closed ?? true,
    fill: isLine || !fill ? { type: 'none' } : fill,
    line,
    text: textBody && (textId || textBody.paragraphs.some((p) => p.runs.some((r) => r.text || r.br))) ? textBody : undefined,
  };
  if (textId) shape.textId = textId;
  return shape;
}

function parsePicture(ctx: SlideContext, part: PartContext, el: Element, transform: Transform): PictureElement | undefined {
  const inherited = inheritance(ctx, part, el);
  const xfrm = [shapeProps(el), ...inherited.sources.map(shapeProps)].map((s) => child(s, 'xfrm')).find((x) => x && child(x, 'off'));
  const box = readBox(xfrm, transform);
  if (!box) return undefined;
  const blipFill = child(el, 'blipFill');
  const src = blipSource(ctx, part, child(blipFill, 'blip'));
  const srcRect = child(blipFill, 'srcRect');
  const crop = srcRect
    ? {
        l: (numAttr(srcRect, 'l') ?? 0) / 100_000,
        t: (numAttr(srcRect, 't') ?? 0) / 100_000,
        r: (numAttr(srcRect, 'r') ?? 0) / 100_000,
        b: (numAttr(srcRect, 'b') ?? 0) / 100_000,
      }
    : undefined;
  const line = readLine(ctx, path(el, 'spPr', 'ln'), child(el, 'style'));
  return { type: 'picture', box, src, crop, line };
}

function blipSource(ctx: SlideContext, part: PartContext, blip: Element | undefined): string | undefined {
  if (!blip) return undefined;
  // Prefer an SVG rendition when the picture has one.
  const svg = children(child(blip, 'extLst'), 'ext')
    .map((ext) => child(ext, 'svgBlip'))
    .find(Boolean);
  for (const candidate of [svg, blip]) {
    const rel = candidate && part.rels.get(relAttr(candidate, 'embed') ?? '');
    if (!rel || rel.external) continue;
    const contentType = imageContentType(rel.target);
    if (!isBrowserImage(contentType)) continue;
    ctx.media.set(rel.target, { path: rel.target, contentType: contentType! });
    return rel.target;
  }
  return undefined;
}

function parseGraphicFrame(ctx: SlideContext, part: PartContext, el: Element, transform: Transform): SlideElement | undefined {
  const inherited = inheritance(ctx, part, el);
  const xfrm = [el, ...inherited.sources].map((s) => child(s, 'xfrm') ?? path(s, 'spPr', 'xfrm')).find((x) => x && child(x, 'off'));
  const box = readBox(xfrm, transform);
  if (!box) return undefined;
  const data = path(el, 'graphic', 'graphicData');
  const uri = attr(data, 'uri') ?? '';
  const tbl = child(data, 'tbl');
  if (tbl) return parseTable(ctx, part, tbl, box);
  const label = uri.endsWith('/chart') || uri.includes('chartex') ? 'Chart' : uri.endsWith('/diagram') ? 'SmartArt' : uri.endsWith('/ole') ? 'Embedded object' : 'Unsupported content';
  return { type: 'unsupported', box, label };
}

// ---------------------------------------------------------------------------------------------
// Tables

const TABLE_PARTS = ['wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'lastCol', 'firstCol', 'lastRow', 'firstRow'] as const;

function parseTable(ctx: SlideContext, part: PartContext, tbl: Element, box: Box): TableElement {
  const tblPr = child(tbl, 'tblPr');
  const styleId = text(child(tblPr, 'tableStyleId')) || attr(child(ctx.tableStyles, 'tblStyleLst'), 'def') || attr(ctx.tableStyles, 'def');
  const tableStyle = children(ctx.tableStyles, 'tblStyle').find((s) => attr(s, 'styleId') === styleId);
  const flags = {
    firstRow: boolAttr(tblPr, 'firstRow'),
    lastRow: boolAttr(tblPr, 'lastRow'),
    firstCol: boolAttr(tblPr, 'firstCol'),
    lastCol: boolAttr(tblPr, 'lastCol'),
    bandRow: boolAttr(tblPr, 'bandRow'),
    bandCol: boolAttr(tblPr, 'bandCol'),
  };
  const columns = children(child(tbl, 'tblGrid'), 'gridCol').map((c) => px(numAttr(c, 'w')));
  const rowEls = children(tbl, 'tr');
  const rows = rowEls.map((tr, r) => ({
    height: px(numAttr(tr, 'h')),
    cells: children(tr, 'tc').map((tc, c): TableCell => {
      const applicable = TABLE_PARTS.filter((name) => {
        switch (name) {
          case 'wholeTbl': return true;
          case 'firstRow': return flags.firstRow && r === 0;
          case 'lastRow': return flags.lastRow && r === rowEls.length - 1;
          case 'firstCol': return flags.firstCol && c === 0;
          case 'lastCol': return flags.lastCol && c === columns.length - 1;
          case 'band1H': return flags.bandRow && (r - (flags.firstRow ? 1 : 0)) % 2 === 0;
          case 'band2H': return flags.bandRow && (r - (flags.firstRow ? 1 : 0)) % 2 === 1;
          case 'band1V': return flags.bandCol && (c - (flags.firstCol ? 1 : 0)) % 2 === 0;
          case 'band2V': return flags.bandCol && (c - (flags.firstCol ? 1 : 0)) % 2 === 1;
        }
      }).map((name) => child(tableStyle, name)).filter((x): x is Element => !!x);

      const tcPr = child(tc, 'tcPr');
      let fill = readFill(ctx, part, tcPr);
      if (!fill) {
        for (const s of [...applicable].reverse()) {
          const tcStyle = child(s, 'tcStyle');
          fill = readFill(ctx, part, child(tcStyle, 'fill')) ?? styleFill(ctx, child(tcStyle, 'fillRef'), true);
          if (fill) break;
        }
      }
      const txStyle = [...applicable].reverse().map((s) => child(s, 'tcTxStyle'));
      const bold = txStyle.map((t) => attr(t, 'b')).find((v) => v !== undefined);
      const colorHost = txStyle.find((t) => t && colorChoice(t));
      const styleColor = colorHost ? colorToCss(ctx, colorHost) : undefined;

      const borders: TableCell['borders'] = {};
      const sides = { l: ['lnL', 'left', c === 0 ? 'left' : 'insideV'], r: ['lnR', 'right', c === columns.length - 1 ? 'right' : 'insideV'], t: ['lnT', 'top', r === 0 ? 'top' : 'insideH'], b: ['lnB', 'bottom', r === rowEls.length - 1 ? 'bottom' : 'insideH'] } as const;
      for (const [key, [own, , styleSide]] of Object.entries(sides) as [keyof typeof sides, readonly [string, string, string]][]) {
        let ln = child(tcPr, own);
        if (!ln) {
          for (const s of [...applicable].reverse()) {
            ln = path(s, 'tcStyle', 'tcBdr', styleSide, 'ln');
            if (ln) break;
          }
        }
        const line = readLine(ctx, ln, undefined);
        if (line) borders[key] = line;
      }

      const txBody = child(tc, 'txBody');
      const body = readTextBody(ctx, tc, txBody, { sources: [], category: 'other' }, undefined, {
        bold: bold === undefined ? undefined : bold === 'on',
        color: styleColor,
      });
      body.insets = {
        l: px(numAttr(tcPr, 'marL') ?? 91_440),
        r: px(numAttr(tcPr, 'marR') ?? 91_440),
        t: px(numAttr(tcPr, 'marT') ?? 45_720),
        b: px(numAttr(tcPr, 'marB') ?? 45_720),
      };
      const anchor = attr(tcPr, 'anchor');
      body.anchor = anchor === 'ctr' || anchor === 'b' ? anchor : 't';
      return {
        text: body,
        fill: fill ?? { type: 'none' },
        colSpan: numAttr(tc, 'gridSpan') ?? 1,
        rowSpan: numAttr(tc, 'rowSpan') ?? 1,
        merged: boolAttr(tc, 'hMerge') || boolAttr(tc, 'vMerge'),
        borders,
      };
    }),
  }));
  return { type: 'table', box, columns, rows };
}

// ---------------------------------------------------------------------------------------------
// Colors, fills and lines

const COLOR_TAGS = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'];

function colorChoice(el: Element | undefined): Element | undefined {
  return children(el).find((c) => COLOR_TAGS.includes(c.localName ?? ''));
}

function resolveColor(ctx: SlideContext, host: Element | undefined, placeholderColor?: Rgb): Rgb | undefined {
  const el = colorChoice(host);
  if (!el) return undefined;
  let base: Rgb | undefined;
  const val = attr(el, 'val') ?? '';
  switch (el.localName) {
    case 'srgbClr':
      base = hexToRgb(val);
      break;
    case 'sysClr':
      base = hexToRgb(attr(el, 'lastClr') ?? (val === 'window' ? 'FFFFFF' : '000000'));
      break;
    case 'prstClr':
      base = hexToRgb(presetColor(val));
      break;
    case 'scrgbClr':
      base = {
        r: ((numAttr(el, 'r') ?? 0) / 100_000) * 255,
        g: ((numAttr(el, 'g') ?? 0) / 100_000) * 255,
        b: ((numAttr(el, 'b') ?? 0) / 100_000) * 255,
        a: 1,
      };
      break;
    case 'schemeClr': {
      if (val === 'phClr') {
        base = placeholderColor;
        break;
      }
      const name = ctx.clrMap[val] ?? val;
      const hex = ctx.theme.colors[name];
      base = hex ? hexToRgb(hex) : undefined;
      break;
    }
  }
  return base ? applyDrawingModifiers(base, children(el)) : undefined;
}

function colorToCss(ctx: SlideContext, host: Element | undefined, placeholderColor?: Rgb): string | undefined {
  const rgb = resolveColor(ctx, host, placeholderColor);
  return rgb ? rgbToCss(rgb) : undefined;
}

/** Reads an explicit fill from a properties element; undefined means "not specified here". */
function readFill(ctx: SlideContext, part: PartContext, props: Element | undefined, placeholderColor?: Rgb): Fill | undefined {
  if (!props) return undefined;
  if (child(props, 'noFill')) return { type: 'none' };
  const solid = child(props, 'solidFill');
  if (solid) {
    const color = colorToCss(ctx, solid, placeholderColor);
    return color ? { type: 'solid', color } : undefined;
  }
  const grad = child(props, 'gradFill');
  if (grad) {
    const stops = children(child(grad, 'gsLst'), 'gs')
      .map((gs) => ({ pos: (numAttr(gs, 'pos') ?? 0) / 1000, color: colorToCss(ctx, gs, placeholderColor) }))
      .filter((s) => s.color)
      .sort((a, b) => a.pos - b.pos);
    if (!stops.length) return undefined;
    if (child(grad, 'path')) {
      return { type: 'gradient', css: `radial-gradient(circle, ${stops.map((s) => `${s.color} ${s.pos}%`).join(', ')})` };
    }
    const angle = (numAttr(child(grad, 'lin'), 'ang') ?? 0) / 60_000 + 90;
    return { type: 'gradient', css: `linear-gradient(${angle}deg, ${stops.map((s) => `${s.color} ${s.pos}%`).join(', ')})` };
  }
  const blip = path(props, 'blipFill', 'blip');
  if (blip) {
    const src = blipSource(ctx, part, blip);
    return src ? { type: 'image', src } : undefined;
  }
  const patt = child(props, 'pattFill');
  if (patt) {
    const color = colorToCss(ctx, child(patt, 'fgClr'));
    return color ? { type: 'solid', color } : undefined;
  }
  return undefined;
}

/** Fill from a shape style reference (`<p:style><a:fillRef idx="1">`). */
function styleFill(ctx: SlideContext, styleOrRef: Element | undefined, isRef = false): Fill | undefined {
  const ref = isRef ? styleOrRef : child(styleOrRef, 'fillRef');
  if (!ref) return undefined;
  if ((numAttr(ref, 'idx') ?? 0) === 0) return { type: 'none' };
  const color = colorToCss(ctx, ref);
  return color ? { type: 'solid', color } : undefined;
}

const THEME_LINE_WIDTHS = [0, 6350, 12700, 19050];
const DASHES: Record<string, number[]> = {
  dash: [4, 3], sysDash: [3, 1], dot: [1, 1], sysDot: [1, 1], lgDash: [8, 3],
  dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

function readLine(ctx: SlideContext, ln: Element | undefined, style: Element | undefined): Line | undefined {
  const lnRef = child(style, 'lnRef');
  const refIdx = numAttr(lnRef, 'idx') ?? 0;
  if (child(ln, 'noFill')) return undefined;
  let color = colorToCss(ctx, child(ln, 'solidFill'));
  if (!color && child(ln, 'gradFill')) color = colorToCss(ctx, path(ln, 'gradFill', 'gsLst', 'gs'));
  if (!color && refIdx > 0) color = colorToCss(ctx, lnRef);
  if (!color) return undefined;
  const widthEmu = numAttr(ln, 'w') ?? THEME_LINE_WIDTHS[Math.min(refIdx, 3)] ?? 9525;
  const width = Math.max(px(widthEmu), 0.75);
  const dashName = attr(child(ln, 'prstDash'), 'val');
  const dash = dashName && DASHES[dashName] ? DASHES[dashName].map((d) => d * width).join(' ') : undefined;
  const end = (name: string) => {
    const t = attr(child(ln, name), 'type');
    return t && t !== 'none' ? t : undefined;
  };
  return { color, width, dash, headEnd: end('headEnd'), tailEnd: end('tailEnd') };
}

function resolveBackground(ctx: SlideContext): Fill {
  for (const part of [ctx.slide, ctx.layout, ctx.master]) {
    const bg = path(part?.root, 'cSld', 'bg');
    if (!part || !bg) continue;
    const bgPr = child(bg, 'bgPr');
    if (bgPr) {
      const fill = readFill(ctx, part, bgPr);
      if (fill) return fill;
    }
    const bgRef = child(bg, 'bgRef');
    if (bgRef) {
      const color = colorToCss(ctx, bgRef);
      if (color) return { type: 'solid', color };
    }
  }
  return { type: 'solid', color: '#ffffff' };
}

// ---------------------------------------------------------------------------------------------
// Text

interface RunDefaults {
  bold?: boolean;
  color?: string;
}

function readTextBody(
  ctx: SlideContext,
  shape: Element,
  txBody: Element | undefined,
  inherited: Inherited,
  style: Element | undefined,
  overrides: RunDefaults = {},
): TextBody {
  // bodyPr attributes inherit from the placeholder chain (own first).
  const bodyPrs = [child(txBody, 'bodyPr'), ...inherited.sources.map((s) => path(s, 'txBody', 'bodyPr'))];
  const bodyAttr = (name: string) => bodyPrs.map((b) => attr(b, name)).find((v) => v !== undefined);
  const bodyNum = (name: string, def: number) => {
    const v = bodyAttr(name);
    return v === undefined ? def : Number(v);
  };
  const autofit = bodyPrs.map((b) => child(b, 'normAutofit')).find(Boolean);
  const fontScale = (numAttr(autofit, 'fontScale') ?? 100_000) / 100_000;
  const spacingReduction = (numAttr(autofit, 'lnSpcReduction') ?? 0) / 100_000;

  const masterStyle =
    inherited.category === 'title'
      ? child(ctx.masterTextStyles, 'titleStyle')
      : inherited.category === 'body'
        ? child(ctx.masterTextStyles, 'bodyStyle')
        : placeholderOf(shape)
          ? child(ctx.masterTextStyles, 'otherStyle')
          : ctx.defaultTextStyle;
  const listStyles = [child(txBody, 'lstStyle'), ...inherited.sources.map((s) => path(s, 'txBody', 'lstStyle')), masterStyle];

  const fontRef = child(style, 'fontRef');
  const fontRefColor = colorToCss(ctx, fontRef);
  const fontRefFace = attr(fontRef, 'idx') === 'major' ? '+mj-lt' : undefined;

  const defaultTextColor = '#' + (ctx.theme.colors[ctx.clrMap.tx1 ?? 'dk1'] ?? '000000');
  const counters = new Map<number, number>();
  const paragraphs: Paragraph[] = [];
  // A body without paragraphs (or no body at all) reads as one empty paragraph, sized for new text.
  const paragraphEls = children(txBody, 'p');
  for (const p of paragraphEls.length ? paragraphEls : [undefined]) {
    const pPr = child(p, 'pPr');
    const level = numAttr(pPr, 'lvl') ?? 0;
    const levelProps = [pPr, ...listStyles.map((ls) => child(ls, `lvl${level + 1}pPr`))];
    const pAttr = (name: string) => levelProps.map((e) => attr(e, name)).find((v) => v !== undefined);
    const pChild = (name: string) => levelProps.map((e) => child(e, name)).find(Boolean);
    const defRPrs = levelProps.map((e) => child(e, 'defRPr'));

    const runProps = (rPr: Element | undefined): Run => {
      const chain = [rPr, ...defRPrs];
      const a = (name: string) => chain.map((e) => attr(e, name)).find((v) => v !== undefined);
      const sz = Number(a('sz') ?? 1800) / 100;
      const run: Run = { text: '', size: sz * PX_PER_PT * fontScale };
      // Overrides (table styles) rank below the run's own properties but above inherited list styles.
      const b = attr(rPr, 'b') ?? (overrides.bold !== undefined ? String(overrides.bold) : a('b'));
      if (b === '1' || b === 'true') run.bold = true;
      if (['1', 'true'].includes(a('i') ?? '')) run.italic = true;
      const u = a('u');
      if (u && u !== 'none') run.underline = true;
      const strike = a('strike');
      if (strike && strike !== 'noStrike') run.strike = true;
      const baseline = Number(a('baseline') ?? 0);
      if (baseline) run.baseline = baseline / 1000;
      const fillOf = (e: Element | undefined) => (e && child(e, 'solidFill') ? colorToCss(ctx, child(e, 'solidFill')) : undefined);
      // The shape style's font color (fontRef) beats only the master / presentation-wide defaults (last entry).
      const specific = defRPrs.slice(0, -1).find((e) => child(e, 'solidFill'));
      run.color = fillOf(rPr) ?? overrides.color ?? fillOf(specific) ?? fontRefColor ?? fillOf(defRPrs[defRPrs.length - 1]) ?? defaultTextColor;
      const latin = chain.map((e) => attr(child(e, 'latin'), 'typeface')).find(Boolean) ?? fontRefFace ?? '+mn-lt';
      const ea = chain.map((e) => attr(child(e, 'ea'), 'typeface')).find(Boolean) ?? (latin.startsWith('+mj') ? '+mj-ea' : '+mn-ea');
      run.font = fontStack(ctx.theme, latin, ea);
      return run;
    };

    const runs: Run[] = [];
    for (const r of children(p)) {
      if (r.localName === 'r' || r.localName === 'fld') {
        const run = runProps(child(r, 'rPr'));
        run.text = r.localName === 'fld' && attr(r, 'type') === 'slidenum' ? String(ctx.slideNumber) : text(child(r, 't'));
        runs.push(run);
      } else if (r.localName === 'br') {
        runs.push({ ...runProps(child(r, 'rPr')), br: true });
      }
    }
    const endProps = runProps(child(p, 'endParaRPr'));

    let bullet: string | undefined;
    const hasText = runs.some((r) => r.text.trim());
    const bulletEl = levelProps.map((e) => children(e).find((c) => ['buNone', 'buChar', 'buAutoNum'].includes(c.localName ?? ''))).find(Boolean);
    if (hasText && bulletEl?.localName === 'buChar') {
      bullet = attr(bulletEl, 'char') ?? '•';
      counters.delete(level);
    } else if (hasText && bulletEl?.localName === 'buAutoNum') {
      const n = (counters.get(level) ?? (numAttr(bulletEl, 'startAt') ?? 1) - 1) + 1;
      counters.set(level, n);
      bullet = autoNumber(attr(bulletEl, 'type') ?? 'arabicPeriod', n);
    } else {
      counters.delete(level);
    }
    for (const key of [...counters.keys()]) if (key > level) counters.delete(key);

    const spacing = (name: string): number => {
      const el = pChild(name);
      const pts = numAttr(child(el, 'spcPts'), 'val');
      if (pts !== undefined) return (pts / 100) * PX_PER_PT;
      const pct = numAttr(child(el, 'spcPct'), 'val');
      return pct !== undefined ? (pct / 100_000) * endProps.size : 0;
    };
    const lnSpc = pChild('lnSpc');
    const lnPct = numAttr(child(lnSpc, 'spcPct'), 'val');
    const lnPts = numAttr(child(lnSpc, 'spcPts'), 'val');
    const algn = pAttr('algn');
    paragraphs.push({
      align: algn === 'ctr' || algn === 'r' || algn === 'just' || algn === 'dist' ? algn : 'l',
      level,
      runs,
      bullet,
      marginLeft: px(Number(pAttr('marL') ?? 0)),
      indent: px(Number(pAttr('indent') ?? 0)),
      spaceBefore: spacing('spcBef'),
      spaceAfter: spacing('spcAft'),
      lineSpacing:
        lnPct !== undefined
          ? { percent: Math.max(lnPct / 100_000 - spacingReduction, 0.5) }
          : lnPts !== undefined
            ? { px: (lnPts / 100) * PX_PER_PT }
            : undefined,
      endSize: endProps.size,
    });
  }

  const anchor = bodyAttr('anchor');
  const vert = bodyAttr('vert');
  return {
    paragraphs,
    anchor: anchor === 'ctr' || anchor === 'b' ? anchor : 't',
    insets: {
      l: px(bodyNum('lIns', 91_440)),
      t: px(bodyNum('tIns', 45_720)),
      r: px(bodyNum('rIns', 91_440)),
      b: px(bodyNum('bIns', 45_720)),
    },
    wrap: bodyAttr('wrap') !== 'none',
    vertical: vert === 'vert' || vert === 'vert270' || vert === 'eaVert' ? vert : undefined,
  };
}

function fontStack(theme: Theme, latin: string, ea: string): string {
  const resolve = (face: string) => {
    if (face === '+mj-lt') return theme.majorFont.latin;
    if (face === '+mn-lt') return theme.minorFont.latin;
    if (face === '+mj-ea') return theme.majorFont.ea;
    if (face === '+mn-ea') return theme.minorFont.ea;
    return face;
  };
  const faces = [resolve(latin), resolve(ea)].filter((f): f is string => !!f);
  return [...new Set(faces)].map((f) => `"${f.replace(/"/g, '')}"`).join(', ');
}

function autoNumber(type: string, n: number): string {
  const roman = (num: number) => {
    const table: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let out = '';
    for (const [v, s] of table) while (num >= v) { out += s; num -= v; }
    return out;
  };
  const alpha = (num: number) => String.fromCharCode(96 + (((num - 1) % 26) + 1));
  let core: string;
  if (type.startsWith('alphaLc')) core = alpha(n);
  else if (type.startsWith('alphaUc')) core = alpha(n).toUpperCase();
  else if (type.startsWith('romanLc')) core = roman(n);
  else if (type.startsWith('romanUc')) core = roman(n).toUpperCase();
  else if (type.startsWith('circleNum')) core = n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
  else core = String(n);
  if (type.endsWith('ParenBoth')) return `(${core})`;
  if (type.endsWith('ParenR')) return `${core})`;
  if (type.endsWith('Period')) return `${core}.`;
  return core;
}
