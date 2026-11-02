// Writes edited copies of the Word and PowerPoint fixtures (every kind of text and slide edit the
// editors make), for opening in Office to check that the results open without repair.
// Usage: node scripts/make-edited-samples.mjs [--out <folder>]
// Then, on Windows with Office: powershell -ExecutionPolicy Bypass -File scripts/check-office.ps1
import * as esbuild from 'esbuild';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const outIndex = args.indexOf('--out');
const outDir = path.resolve(outIndex >= 0 ? args[outIndex + 1] : path.join(os.tmpdir(), 'office-workbench-edited'));

const cacheDir = path.join(root, 'node_modules/.cache/office-workbench');
await mkdir(cacheDir, { recursive: true });
const coreBundle = path.join(cacheDir, 'core.mjs');
await esbuild.build({
  entryPoints: [path.join(root, 'packages/core/src/index.ts')],
  outfile: coreBundle,
  bundle: true,
  platform: 'node',
  format: 'esm',
  logLevel: 'warning',
});
const { WordDocument, PresentationDocument } = await import(pathToFileURL(coreBundle).href);
const fixture = async (name) => new Uint8Array(await readFile(path.join(root, 'test-fixtures', name)));
await mkdir(outDir, { recursive: true });

async function save(name, doc) {
  const target = path.join(outDir, name);
  await writeFile(target, await doc.save());
  console.log(`Wrote ${target}`);
}

// Word: rewrite, split, insert and delete paragraphs, in the body and in a table cell.
{
  const doc = await WordDocument.open(await fixture('sample.docx'));
  await doc.apply({ type: 'setParagraphText', id: '0', text: '見出し（編集済み）' });
  await doc.apply({ type: 'setParagraphText', id: '1', text: '本文\tタブ\n改行の後 ' });
  await doc.apply({ type: 'insertParagraph', after: '5', text: '四つ目の項目' });
  await doc.apply({ type: 'insertParagraph', after: '1', text: '後半', update: '前半' });
  await doc.apply({ type: 'deleteParagraph', id: '8' }); // "表:"
  await doc.apply({ type: 'setParagraphText', id: '8/2/2/1', text: '空だったセル' });
  await doc.apply({ type: 'setParagraphText', id: '8/0/1/1', text: '数量\n（個）' });
  const blocks = await doc.blocks();
  await doc.apply({ type: 'insertParagraph', after: blocks[blocks.length - 1].id, text: '最後に追加した段落' });
  await save('edited.docx', doc);
}

// PowerPoint: shape text (with an empty placeholder and a shape without text), table cells, and
// duplicating, deleting and moving slides.
{
  const doc = await PresentationDocument.open(await fixture('sample.pptx'));
  await doc.apply({ type: 'setShapeText', slide: 0, shape: '2', paragraphs: ['タイトル（編集済み）', '二行目'] });
  await doc.apply({ type: 'setShapeText', slide: 1, shape: '3', paragraphs: ['第一レベルの項目', '追加した項目', '第二レベルの項目', '改行\nあり'] });
  await doc.apply({ type: 'setShapeText', slide: 2, shape: '6', paragraphs: ['楕円に文字'] });
  await doc.apply({ type: 'setShapeText', slide: 3, shape: '4', paragraphs: ['空のプレースホルダーに入力'] });
  await doc.apply({ type: 'setTableCellText', slide: 3, shape: '5', row: 2, col: 2, paragraphs: ['99%', '（速報）'] });
  await doc.apply({ type: 'duplicateSlide', slide: 0 });
  await doc.apply({ type: 'duplicateSlide', slide: 3 });
  await doc.apply({ type: 'setShapeText', slide: 4, shape: '2', paragraphs: ['複製したスライド'] });
  await doc.apply({ type: 'deleteSlide', slide: 2 });
  await doc.apply({ type: 'moveSlide', slide: 5, to: 0 });
  await save('edited.pptx', doc);
}
