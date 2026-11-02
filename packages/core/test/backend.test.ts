import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { ViewerBackend, type ViewerKind } from '../src';

const load = (name: string) => () => Promise.resolve(new Uint8Array(readFileSync(new URL(`../../../test-fixtures/${name}`, import.meta.url))));
const backend = (kind: ViewerKind) => new ViewerBackend(kind, `sample.${kind}`, load(`sample.${kind}`));

describe('ViewerBackend', () => {
  it('sends Word documents as raw bytes', async () => {
    const [message] = await backend('docx').handle({ type: 'ready' });
    expect(message.type).toBe('docx');
    if (message.type === 'docx') expect(message.data.byteLength).toBeGreaterThan(1000);
  });

  it('sends PDFs as raw bytes', async () => {
    const [message] = await backend('pdf').handle({ type: 'ready' });
    expect(message.type).toBe('pdf');
    if (message.type === 'pdf') expect(new TextDecoder().decode(message.data.subarray(0, 5))).toBe('%PDF-');
  });

  it('loads workbook structure first and sheets on request', async () => {
    const b = backend('xlsx');
    const [info] = await b.handle({ type: 'ready' });
    expect(info).toMatchObject({ type: 'workbook', fileName: 'sample.xlsx' });
    const [sheet] = await b.handle({ type: 'loadSheet', index: 1 });
    expect(sheet.type === 'sheet' && sheet.sheet.name).toBe('Data');
  });

  it('sends presentations with their media', async () => {
    const [message] = await backend('pptx').handle({ type: 'ready' });
    expect(message.type).toBe('presentation');
    if (message.type === 'presentation') {
      expect(message.presentation.slides).toHaveLength(5);
      expect(message.media.map((m) => m.path)).toEqual(['ppt/media/image1.png']);
    }
  });

  it('reports errors as messages instead of throwing', async () => {
    const broken = new ViewerBackend('xlsx', 'broken.xlsx', () => Promise.resolve(new Uint8Array([0, 1, 2])));
    const [message] = await broken.handle({ type: 'ready' });
    expect(message.type).toBe('error');
    const [missing] = await backend('xlsx').handle({ type: 'loadSheet', index: 99 });
    expect(missing).toMatchObject({ type: 'error', message: 'Sheet 99 does not exist.' });
  });
});
