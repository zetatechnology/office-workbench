// Read access to an OOXML package (a ZIP of XML parts plus relationship files).
import JSZip from 'jszip';
import { attr, children, parseXml, type Document } from './xml';


export interface Relationship {
  id: string;
  /** Last segment of the relationship type URI, e.g. "worksheet" or "slideLayout". */
  type: string;
  /** Absolute part path inside the package (no leading slash), or the raw URL when external. */
  target: string;
  external: boolean;
}

const MAX_ENTRIES = 20_000;
const MAX_UNCOMPRESSED_BYTES = 1024 * 1024 * 1024;

export class PackageError extends Error {}

export class OoxmlPackage {
  private readonly xmlCache = new Map<string, Promise<Document | undefined>>();
  private readonly relsCache = new Map<string, Promise<Map<string, Relationship>>>();

  private constructor(private readonly zip: JSZip) {}

  static async open(data: Uint8Array): Promise<OoxmlPackage> {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(data);
    } catch (e) {
      throw new PackageError(`Not a valid Office Open XML file: ${(e as Error).message}`);
    }
    const files = Object.values(zip.files);
    if (files.length > MAX_ENTRIES) throw new PackageError('The file contains too many parts.');
    // Guard against ZIP bombs using the sizes declared in the central directory.
    const declared = files.reduce(
      (sum, f) => sum + ((f as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0),
      0,
    );
    if (declared > MAX_UNCOMPRESSED_BYTES) throw new PackageError('The file is too large to open.');
    if (!zip.file('[Content_Types].xml')) throw new PackageError('Missing [Content_Types].xml; not an OOXML file.');
    return new OoxmlPackage(zip);
  }

  has(partPath: string): boolean {
    return this.zip.file(partPath) !== null;
  }

  async text(partPath: string): Promise<string | undefined> {
    return this.zip.file(partPath)?.async('string');
  }

  async bytes(partPath: string): Promise<Uint8Array | undefined> {
    return this.zip.file(partPath)?.async('uint8array');
  }

  xml(partPath: string): Promise<Document | undefined> {
    let cached = this.xmlCache.get(partPath);
    if (!cached) {
      cached = this.text(partPath).then((t) => (t === undefined ? undefined : parseXml(t)));
      this.xmlCache.set(partPath, cached);
    }
    return cached;
  }

  /** Relationships of a part; pass "" for the package-level relationships. */
  rels(partPath: string): Promise<Map<string, Relationship>> {
    let cached = this.relsCache.get(partPath);
    if (!cached) {
      cached = this.loadRels(partPath);
      this.relsCache.set(partPath, cached);
    }
    return cached;
  }

  async relsByType(partPath: string, type: string): Promise<Relationship[]> {
    return [...(await this.rels(partPath)).values()].filter((r) => r.type === type);
  }

  /** The main document part (word/document.xml, xl/workbook.xml, ppt/presentation.xml). */
  async mainPart(): Promise<string> {
    const rels = await this.rels('');
    for (const rel of rels.values()) {
      if (rel.type === 'officeDocument') return rel.target;
    }
    throw new PackageError('The package has no main document part.');
  }

  private async loadRels(partPath: string): Promise<Map<string, Relationship>> {
    const dir = dirname(partPath);
    const base = partPath.slice(dir.length).replace(/^\//, '');
    const relsPath = (dir ? `${dir}/` : '') + `_rels/${base}.rels`;
    const doc = await this.xml(relsPath);
    const map = new Map<string, Relationship>();
    if (!doc?.documentElement) return map;
    for (const el of children(doc.documentElement, 'Relationship')) {
      const id = attr(el, 'Id') ?? '';
      const rawType = attr(el, 'Type') ?? '';
      const target = attr(el, 'Target') ?? '';
      const external = attr(el, 'TargetMode') === 'External';
      map.set(id, {
        id,
        type: rawType.slice(rawType.lastIndexOf('/') + 1),
        target: external ? target : resolvePartPath(dir, target),
        external,
      });
    }
    return map;
  }
}

export function dirname(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  return i < 0 ? '' : partPath.slice(0, i);
}

/** Resolves a relationship target relative to the directory of its source part. */
export function resolvePartPath(baseDir: string, target: string): string {
  const raw = target.startsWith('/') ? target.slice(1) : (baseDir ? `${baseDir}/` : '') + target;
  const out: string[] = [];
  for (const seg of decodeURI(raw).split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return out.join('/');
}

const IMAGE_TYPES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

export function imageContentType(partPath: string): string | undefined {
  return IMAGE_TYPES[partPath.slice(partPath.lastIndexOf('.') + 1).toLowerCase()];
}

/** Image formats a browser can display directly. */
export function isBrowserImage(contentType: string | undefined): boolean {
  return !!contentType && !/x-emf|x-wmf|tiff/.test(contentType);
}
