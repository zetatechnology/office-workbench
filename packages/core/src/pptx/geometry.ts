// SVG paths for DrawingML preset geometries (a common subset) and custom geometries.
import { attr, child, children, numAttr, type Element } from '../xml';

export interface GeometryPath {
  d: string;
  closed: boolean;
}

const EMU_PER_PX = 9525;

function f(n: number): string {
  return String(Math.round(n * 100) / 100);
}

function polygon(points: [number, number][]): GeometryPath {
  return { d: points.map(([x, y], i) => `${i ? 'L' : 'M'}${f(x)} ${f(y)}`).join(' ') + ' Z', closed: true };
}

function rect(w: number, h: number): GeometryPath {
  return polygon([[0, 0], [w, 0], [w, h], [0, h]]);
}

function roundRect(w: number, h: number, r: number): GeometryPath {
  r = Math.min(r, w / 2, h / 2);
  return {
    d: `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)} V${f(h - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)} H${f(r)} A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)} V${f(r)} A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0 Z`,
    closed: true,
  };
}

function ellipse(w: number, h: number): GeometryPath {
  const rx = w / 2, ry = h / 2;
  return { d: `M0 ${f(ry)} A${f(rx)} ${f(ry)} 0 1 1 ${f(w)} ${f(ry)} A${f(rx)} ${f(ry)} 0 1 1 0 ${f(ry)} Z`, closed: true };
}

function regular(w: number, h: number, sides: number, rotation = -Math.PI / 2): GeometryPath {
  const pts: [number, number][] = [];
  for (let i = 0; i < sides; i++) {
    const a = rotation + (i * 2 * Math.PI) / sides;
    pts.push([w / 2 + (w / 2) * Math.cos(a), h / 2 + (h / 2) * Math.sin(a)]);
  }
  return polygon(pts);
}

function star(w: number, h: number, points: number, inner: number): GeometryPath {
  const pts: [number, number][] = [];
  for (let i = 0; i < points * 2; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / points;
    const r = i % 2 ? inner : 1;
    pts.push([w / 2 + (w / 2) * r * Math.cos(a), h / 2 + (h / 2) * r * Math.sin(a)]);
  }
  return polygon(pts);
}

/** Reads `<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>` into a map. */
export function readAdjustments(prstGeom: Element | undefined): Map<string, number> {
  const map = new Map<string, number>();
  for (const gd of children(child(prstGeom, 'avLst'), 'gd')) {
    const m = /^val\s+(-?\d+)/.exec(attr(gd, 'fmla') ?? '');
    if (m) map.set(attr(gd, 'name') ?? '', Number(m[1]));
  }
  return map;
}

/** Path for a preset geometry, or undefined when the preset is unknown (callers fall back to a rectangle). */
export function presetPath(prst: string, w: number, h: number, adj: Map<string, number> = new Map()): GeometryPath | undefined {
  const ss = Math.min(w, h);
  const a = (name: string, def: number) => (adj.get(name) ?? def) / 100_000;
  switch (prst) {
    case 'rect':
    case 'flowChartProcess':
    case 'flowChartAlternateProcess':
      return prst === 'flowChartAlternateProcess' ? roundRect(w, h, ss * 0.1667) : rect(w, h);
    case 'roundRect':
      return roundRect(w, h, ss * a('adj', 16667));
    case 'flowChartTerminator':
      return roundRect(w, h, h / 2);
    case 'ellipse':
    case 'flowChartConnector':
      return ellipse(w, h);
    case 'triangle':
    case 'flowChartExtract': {
      const x = w * a('adj', 50000);
      return polygon([[x, 0], [w, h], [0, h]]);
    }
    case 'rtTriangle':
      return polygon([[0, 0], [w, h], [0, h]]);
    case 'diamond':
    case 'flowChartDecision':
      return polygon([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram':
    case 'flowChartInputOutput': {
      const x = ss * a('adj', 25000);
      return polygon([[x, 0], [w, 0], [w - x, h], [0, h]]);
    }
    case 'trapezoid': {
      const x = ss * a('adj', 25000);
      return polygon([[x, 0], [w - x, 0], [w, h], [0, h]]);
    }
    case 'pentagon':
      return regular(w, h, 5);
    case 'hexagon': {
      const x = ss * a('adj', 25000);
      return polygon([[x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h], [0, h / 2]]);
    }
    case 'octagon': {
      const x = ss * a('adj', 29289);
      return polygon([[x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]]);
    }
    case 'homePlate': {
      const x = w - ss * a('adj', 50000);
      return polygon([[0, 0], [x, 0], [w, h / 2], [x, h], [0, h]]);
    }
    case 'chevron': {
      const x = ss * a('adj', 50000);
      return polygon([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]]);
    }
    case 'rightArrow':
    case 'leftArrow': {
      const shaft = h * a('adj1', 50000);
      const head = ss * a('adj2', 50000);
      const y1 = (h - shaft) / 2, y2 = y1 + shaft;
      const pts: [number, number][] = [[0, y1], [w - head, y1], [w - head, 0], [w, h / 2], [w - head, h], [w - head, y2], [0, y2]];
      return polygon(prst === 'leftArrow' ? pts.map(([x, y]) => [w - x, y]) : pts);
    }
    case 'upArrow':
    case 'downArrow': {
      const shaft = w * a('adj1', 50000);
      const head = ss * a('adj2', 50000);
      const x1 = (w - shaft) / 2, x2 = x1 + shaft;
      const pts: [number, number][] = [[x1, h], [x1, head], [0, head], [w / 2, 0], [w, head], [x2, head], [x2, h]];
      return polygon(prst === 'downArrow' ? pts.map(([x, y]) => [x, h - y]) : pts);
    }
    case 'leftRightArrow': {
      const shaft = h * a('adj1', 50000);
      const head = ss * a('adj2', 50000);
      const y1 = (h - shaft) / 2, y2 = y1 + shaft;
      return polygon([[0, h / 2], [head, 0], [head, y1], [w - head, y1], [w - head, 0], [w, h / 2], [w - head, h], [w - head, y2], [head, y2], [head, h]]);
    }
    case 'plus':
    case 'mathPlus': {
      const x = ss * a('adj', 25000);
      return polygon([[x, 0], [w - x, 0], [w - x, x], [w, x], [w, h - x], [w - x, h - x], [w - x, h], [x, h], [x, h - x], [0, h - x], [0, x], [x, x]]);
    }
    case 'star4':
      return star(w, h, 4, 0.25);
    case 'star5':
      return star(w, h, 5, 0.382);
    case 'star6':
      return star(w, h, 6, 0.577);
    case 'line':
    case 'straightConnector1':
    case 'bentConnector2':
    case 'bentConnector3':
    case 'bentConnector4':
    case 'curvedConnector3':
      return prst.startsWith('bentConnector3')
        ? { d: `M0 0 H${f(w * a('adj1', 50000))} V${f(h)} H${f(w)}`, closed: false }
        : { d: `M0 0 L${f(w)} ${f(h)}`, closed: false };
    case 'snip1Rect': {
      const x = ss * a('adj', 16667);
      return polygon([[0, 0], [w - x, 0], [w, x], [w, h], [0, h]]);
    }
    case 'flowChartDocument':
      return {
        d: `M0 0 H${f(w)} V${f(h * 0.83)} C${f(w * 0.75)} ${f(h * 0.66)} ${f(w * 0.25)} ${f(h * 1.05)} 0 ${f(h * 0.88)} Z`,
        closed: true,
      };
    case 'can': {
      const ry = ss * a('adj', 25000) / 2;
      return {
        d: `M0 ${f(ry)} A${f(w / 2)} ${f(ry)} 0 0 1 ${f(w)} ${f(ry)} V${f(h - ry)} A${f(w / 2)} ${f(ry)} 0 0 1 0 ${f(h - ry)} Z M0 ${f(ry)} A${f(w / 2)} ${f(ry)} 0 0 0 ${f(w)} ${f(ry)}`,
        closed: true,
      };
    }
    default:
      return undefined;
  }
}

/** Converts `<a:custGeom>` into SVG path data scaled to the shape size. */
export function customPath(custGeom: Element, w: number, h: number): GeometryPath | undefined {
  const parts: string[] = [];
  let closed = false;
  for (const p of children(child(custGeom, 'pathLst'), 'path')) {
    const pw = numAttr(p, 'w') ?? 0;
    const ph = numAttr(p, 'h') ?? 0;
    // Path coordinates are in their own space (w/h), or in EMU of the shape when w/h are absent.
    const sx = pw ? w / pw : 1 / EMU_PER_PX;
    const sy = ph ? h / ph : 1 / EMU_PER_PX;
    const pt = (el: Element | undefined): [number, number] => [(numAttr(el, 'x') ?? 0) * sx, (numAttr(el, 'y') ?? 0) * sy];
    let cur: [number, number] = [0, 0];
    for (const cmd of children(p)) {
      const pts = children(cmd, 'pt').map(pt);
      switch (cmd.localName) {
        case 'moveTo':
          cur = pts[0] ?? cur;
          parts.push(`M${f(cur[0])} ${f(cur[1])}`);
          break;
        case 'lnTo':
          cur = pts[0] ?? cur;
          parts.push(`L${f(cur[0])} ${f(cur[1])}`);
          break;
        case 'cubicBezTo':
          if (pts.length === 3) {
            parts.push(`C${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(' ')}`);
            cur = pts[2];
          }
          break;
        case 'quadBezTo':
          if (pts.length === 2) {
            parts.push(`Q${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(' ')}`);
            cur = pts[1];
          }
          break;
        case 'arcTo': {
          const wR = (numAttr(cmd, 'wR') ?? 0) * sx;
          const hR = (numAttr(cmd, 'hR') ?? 0) * sy;
          const st = ((numAttr(cmd, 'stAng') ?? 0) / 60_000) * (Math.PI / 180);
          const sw = ((numAttr(cmd, 'swAng') ?? 0) / 60_000) * (Math.PI / 180);
          const cx = cur[0] - wR * Math.cos(st);
          const cy = cur[1] - hR * Math.sin(st);
          const end: [number, number] = [cx + wR * Math.cos(st + sw), cy + hR * Math.sin(st + sw)];
          parts.push(`A${f(wR)} ${f(hR)} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${f(end[0])} ${f(end[1])}`);
          cur = end;
          break;
        }
        case 'close':
          parts.push('Z');
          closed = true;
          break;
      }
    }
  }
  return parts.length ? { d: parts.join(' '), closed } : undefined;
}

