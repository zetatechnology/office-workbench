// Read access to an OOXML package (a ZIP of XML parts plus relationship files).
import JSZip from 'jszip';
import { XMLSerializer } from '@xmldom/xmldom';
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

/** Prior contents of the parts touched while journaling; `null` means the part did not exist. */
export type PartSnapshot = Map<string, string | Uint8Array | null>;

export class OoxmlPackage {
  private readonly xmlCache = new Map<string, Promise<Document | undefined>>();
  private readonly relsCache = new Map<string, Promise<Map<string, Relationship>>>();
  /** Parts whose parsed DOM was modified and must be serialized on save. */
  private readonly dirty = new Set<string>();
  private journal?: PartSnapshot;

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
    // JSZip recompresses every entry whose method differs from the one passed to generateAsync,
    // so entries stored uncompressed (Word stores images this way) are pinned to STORE to keep
    // their original bytes.
    for (const f of files) {
      const internal = f as unknown as { _data?: { compression?: { magic?: string } }; options: { compression?: string } };
      if (internal._data?.compression?.magic === '\x00\x00') internal.options.compression = 'STORE';
    }
    return new OoxmlPackage(zip);
  }

  has(partPath: string): boolean {
    return this.zip.file(partPath) !== null;
  }

  /** The part as text, including unsaved edits to its DOM. */
  async text(partPath: string): Promise<string | undefined> {
    if (this.dirty.has(partPath)) {
      const doc = await this.xml(partPath);
      return doc ? new XMLSerializer().serializeToString(doc) : undefined;
    }
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

  /** Part paths in the package, excluding directory entries. */
  parts(): string[] {
    return Object.values(this.zip.files)
      .filter((f) => !f.dir)
      .map((f) => f.name);
  }

  // --- Writing -------------------------------------------------------------------------------
  // Edits mutate the cached DOM of a part; only those parts are re-serialized on save, so every
  // other part keeps its original compressed bytes.

  /** The DOM of a part, marked as modified. Mutate it only after awaiting this call. */
  async edit(partPath: string): Promise<Document> {
    await this.record(partPath);
    const doc = await this.xml(partPath);
    if (!doc) throw new PackageError(`Missing part ${partPath}.`);
    this.dirty.add(partPath);
    if (partPath.endsWith('.rels')) this.relsCache.clear();
    return doc;
  }

  async setText(partPath: string, text: string): Promise<void> {
    await this.record(partPath);
    this.zip.file(partPath, text, { createFolders: false });
    this.forget(partPath);
  }

  async remove(partPath: string): Promise<void> {
    if (!this.has(partPath)) return;
    await this.record(partPath);
    this.zip.remove(partPath);
    this.forget(partPath);
  }

  /** Starts recording the prior state of every part modified until `commit()`. */
  begin(): void {
    this.journal = new Map();
  }

  commit(): PartSnapshot {
    const journal = this.journal ?? new Map();
    this.journal = undefined;
    return journal;
  }

  /** Puts parts back the way a snapshot recorded them. */
  restore(snapshot: PartSnapshot): void {
    for (const [partPath, content] of snapshot) {
      if (content === null) this.zip.remove(partPath);
      else this.zip.file(partPath, content, { createFolders: false });
      this.forget(partPath);
    }
  }

  async save(): Promise<Uint8Array> {
    const serializer = new XMLSerializer();
    for (const partPath of this.dirty) {
      const doc = await this.xml(partPath);
      if (doc) this.zip.file(partPath, serializer.serializeToString(doc), { createFolders: false });
    }
    this.dirty.clear();
    return this.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  private async record(partPath: string): Promise<void> {
    if (!this.journal || this.journal.has(partPath)) return;
    let prior: string | Uint8Array | null = null;
    if (this.dirty.has(partPath)) {
      const doc = await this.xml(partPath);
      prior = doc ? new XMLSerializer().serializeToString(doc) : null;
    } else if (this.has(partPath)) {
      prior = (await this.bytes(partPath)) ?? null;
    }
    this.journal.set(partPath, prior);
  }

  private forget(partPath: string): void {
    this.xmlCache.delete(partPath);
    this.dirty.delete(partPath);
    if (partPath.endsWith('.rels')) this.relsCache.clear();
  }

  private async loadRels(partPath: string): Promise<Map<string, Relationship>> {
    const dir = dirname(partPath);
    const doc = await this.xml(relsPathOf(partPath));
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

/** "xl/workbook.xml" → "xl/_rels/workbook.xml.rels"; "" → "_rels/.rels". */
export function relsPathOf(partPath: string): string {
  const dir = dirname(partPath);
  const base = partPath.slice(dir.length).replace(/^\//, '');
  return (dir ? `${dir}/` : '') + `_rels/${base}.rels`;
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

/** "ppt/slides/_rels/slide1.xml.rels" → "ppt/slides/slide1.xml"; "_rels/.rels" → "". */
export function relsSourceOf(relsPath: string): string | undefined {
  const match = /^(.*?)_rels\/([^/]*)\.rels$/.exec(relsPath);
  return match ? match[1] + match[2] : undefined;
}

/** A relationship target for `toPart`, relative to the directory of `fromPart`. */
export function relativeTarget(fromPart: string, toPart: string): string {
  const from = dirname(fromPart).split('/').filter(Boolean);
  const to = toPart.split('/');
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++;
  return [...Array<string>(from.length - common).fill('..'), ...to.slice(common)].join('/');
}
