import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
  EditError, parsePresentation, PresentationDocument, PresentationSession, toMarkdown, type HostMessage, type Presentation,
  type ShapeElement, type Slide, type TableElement,
} from '../src';
import { changedEntries, fixture, part } from './helpers';

const sample = () => fixture('sample.pptx');

const shapes = (slide: Slide) => slide.elements.filter((e): e is ShapeElement => e.type === 'shape');
const textOf = (shape: ShapeElement | undefined) => shape?.text?.paragraphs.map((p) => p.runs.map((r) => (r.br ? '\n' : r.text)).join('')) ?? [];
const shapeById = (slide: Slide, id: string) => shapes(slide).find((s) => s.textId === id);

async function presentation(data: Uint8Array): Promise<Presentation> {
  return (await parsePresentation(data)).presentation;
}

/** The sample with one slide's XML transformed. */
async function withSlide(n: number, edit: (xml: string) => string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(sample());
  const name = `ppt/slides/slide${n}.xml`;
  zip.file(name, edit(await zip.file(name)!.async('string')), { createFolders: false });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('editable shapes in the slide model', () => {
  it('marks the slide’s own text shapes and tables, including empty placeholders', async () => {
    const { slides } = await presentation(sample());
    expect(shapes(slides[0]).map((s) => s.textId)).toEqual(['2', '3']);
    // Slide 4: title, an empty chart placeholder (no text), an empty text placeholder, the table.
    const slide4 = slides[3];
    expect(slide4.elements.map((e) => (e.type === 'shape' || e.type === 'table' ? (e.textId ?? '-') : e.type))).toEqual(['2', '-', '4', '5']);
    expect(shapeById(slide4, '4')!.text!.paragraphs).toHaveLength(1);
    // The connector line and the chart placeholder hold no text.
    expect(shapes(slides[2]).filter((s) => !s.textId)).toHaveLength(2);
  });
});

describe('editing slide text', () => {
  it('rewrites changed paragraphs only, keeping paragraph and run properties', async () => {
    const original = sample();
    const doc = await PresentationDocument.open(original);
    await doc.apply({
      type: 'setShapeText',
      slide: 1,
      shape: '3',
      paragraphs: ['第一レベルの項目', '新しい項目', '第二レベル（変更）', '第一レベル（太字）', 'Last\nitem'],
    });
    const saved = await doc.save();
    expect(changedEntries(original, saved)).toEqual(['ppt/slides/slide2.xml']);

    const xml = (await part(saved, 'ppt/slides/slide2.xml'))!;
    const body = xml.slice(xml.indexOf('name="テキスト プレースホルダー 2"'));
    const paragraphs = body.match(/<a:p>[\s\S]*?<\/a:p>/g)!;
    expect(paragraphs).toEqual([
      '<a:p><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US"/><a:t>第一レベルの項目</a:t></a:r></a:p>',
      // Inserted before the second-level item: copies the properties of the paragraph before it.
      '<a:p><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US"/><a:t>新しい項目</a:t></a:r></a:p>',
      '<a:p><a:pPr lvl="1"/><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US"/><a:t>第二レベル（変更）</a:t></a:r></a:p>',
      '<a:p><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US" b="1"/><a:t>第一レベル（太字）</a:t></a:r></a:p>',
      '<a:p><a:r><a:rPr kumimoji="1" lang="en-US" altLang="ja-JP"/><a:t>Last</a:t></a:r><a:br><a:rPr kumimoji="1" lang="en-US" altLang="ja-JP"/></a:br>' +
        '<a:r><a:rPr kumimoji="1" lang="en-US" altLang="ja-JP"/><a:t>item</a:t></a:r><a:endParaRPr kumimoji="1" lang="ja-JP" altLang="en-US"/></a:p>',
    ]);
    const shape = shapeById((await presentation(saved)).slides[1], '3')!;
    expect(textOf(shape)).toEqual(['第一レベルの項目', '新しい項目', '第二レベル（変更）', '第一レベル（太字）', 'Last\nitem']);
    expect(shape.text!.paragraphs.map((p) => p.level)).toEqual([0, 0, 1, 0, 0]);
  });

  it('keeps the other paragraphs untouched when one is removed', async () => {
    const doc = await PresentationDocument.open(sample());
    await doc.apply({ type: 'setShapeText', slide: 1, shape: '3', paragraphs: ['第一レベルの項目', '第一レベル（太字）', 'Last item'] });
    const xml = (await part(await doc.save(), 'ppt/slides/slide2.xml'))!;
    const original = (await part(sample(), 'ppt/slides/slide2.xml'))!;
    const third = '<a:p><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US" b="1"/><a:t>第一レベル（太字）</a:t></a:r></a:p>';
    expect(original).toContain(third);
    expect(xml).toContain(third);
    expect(xml).not.toContain('第二レベル');
  });

  it('keeps a slide number field as it is while other paragraphs of the shape change', async () => {
    // The stored field text is a placeholder; the slide model shows the real number.
    const withField = await withSlide(2, (xml) =>
      xml.replace('<a:t>Last item</a:t></a:r>', '<a:t>Last item</a:t></a:r><a:fld id="{1}" type="slidenum"><a:t>‹#›</a:t></a:fld>'),
    );
    const { presentation: pres } = await parsePresentation(withField);
    const current = textOf(shapeById(pres.slides[1], '3'));
    expect(current[3]).toBe('Last item2');
    const doc = await PresentationDocument.open(withField);
    await doc.apply({ type: 'setShapeText', slide: 1, shape: '3', paragraphs: ['changed', ...current.slice(1)] });
    const xml = (await part(await doc.save(), 'ppt/slides/slide2.xml'))!;
    expect(xml).toContain('<a:fld id="{1}" type="slidenum"><a:t>‹#›</a:t></a:fld>');
    expect(xml).toContain('<a:t>changed</a:t>');
  });

  it('finds the slide being edited by its id after slides moved', async () => {
    const doc = await PresentationDocument.open(sample());
    const { slides } = (await doc.read(false)).presentation;
    // Typing started on slide 3 (id 258); meanwhile another edit moved it to the front.
    await doc.apply({ type: 'moveSlide', slide: 2, to: 0 });
    await doc.apply({ type: 'setShapeText', slide: 2, slideId: slides[2].id, shape: '2', paragraphs: ['図形（編集）'] });
    const after = (await doc.read(false)).presentation.slides;
    expect(textOf(shapes(after[0])[0])).toEqual(['図形（編集）']);
    expect(textOf(shapes(after[2])[0])).toEqual(['箇条書き']);
  });

  it('writes into empty placeholders and shapes without a text body', async () => {
    const original = await withSlide(1, (xml) => xml.replace(/(<p:cNvPr id="3"[\s\S]*?<p:spPr\/>)<p:txBody>[\s\S]*?<\/p:txBody>/, '$1'));
    expect(textOf(shapeById((await presentation(original)).slides[0], '3'))).toEqual(['']);
    const doc = await PresentationDocument.open(original);
    await doc.apply({ type: 'setShapeText', slide: 0, shape: '3', paragraphs: ['サブタイトル'] });
    await doc.apply({ type: 'setShapeText', slide: 3, shape: '4', paragraphs: ['注記', ''] });
    const saved = await doc.save();
    const slide1 = (await part(saved, 'ppt/slides/slide1.xml'))!;
    expect(slide1).toContain('<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP"/><a:t>サブタイトル</a:t></a:r></a:p></p:txBody></p:sp>');
    const slide4 = (await part(saved, 'ppt/slides/slide4.xml'))!;
    expect(slide4).toContain(
      '<a:p><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US"/><a:t>注記</a:t></a:r><a:endParaRPr kumimoji="1" lang="ja-JP" altLang="en-US"/></a:p>' +
        '<a:p><a:endParaRPr kumimoji="1" lang="ja-JP" altLang="en-US"/></a:p></p:txBody>',
    );
    const { slides } = await presentation(saved);
    expect(textOf(shapeById(slides[0], '3'))).toEqual(['サブタイトル']);
    expect(slides[0].elements.find((e) => e.type === 'shape' && e.textId === '3')?.box.w).toBeGreaterThan(500);
  });

  it('edits table cells', async () => {
    const original = sample();
    const doc = await PresentationDocument.open(original);
    await doc.apply({ type: 'setTableCellText', slide: 3, shape: '5', row: 1, col: 1, paragraphs: ['1,350', '(見込み)'] });
    const saved = await doc.save();
    expect(changedEntries(original, saved)).toEqual(['ppt/slides/slide4.xml']);
    const table = (await presentation(saved)).slides[3].elements.find((e): e is TableElement => e.type === 'table')!;
    expect(table.textId).toBe('5');
    expect(table.rows[1].cells[1].text.paragraphs.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['1,350', '(見込み)']);
    expect(await toMarkdown('pptx', saved)).toContain('| 東日本 | 1,350<br>(見込み) | 105% |');
  });

  it('refuses text it cannot write safely', async () => {
    const withField = await withSlide(2, (xml) =>
      xml.replace('<a:t>Last item</a:t></a:r>', '<a:t>Last item</a:t></a:r><a:fld id="{1}" type="slidenum"><a:t>2</a:t></a:fld>'),
    );
    const doc = await PresentationDocument.open(withField);
    await expect(doc.apply({ type: 'setShapeText', slide: 1, shape: '3', paragraphs: ['a', 'b', 'c', 'd'] })).rejects.toThrow(/field/);
    // Paragraphs with fields may stay as they are while others change.
    const { slides } = await presentation(withField);
    const current = textOf(shapeById(slides[1], '3'));
    await doc.apply({ type: 'setShapeText', slide: 1, shape: '3', paragraphs: ['changed', ...current.slice(1)] });

    await expect(doc.apply({ type: 'setShapeText', slide: 3, shape: '3', paragraphs: ['x'] })).rejects.toThrow(/chart/);
    await expect(doc.apply({ type: 'setShapeText', slide: 0, slideId: '999', shape: '2', paragraphs: ['x'] })).rejects.toThrow('The slide being edited was deleted.');
    await expect(doc.apply({ type: 'setShapeText', slide: 2, shape: '99', paragraphs: ['x'] })).rejects.toThrow('That shape no longer exists on the slide.');
    await expect(doc.apply({ type: 'setShapeText', slide: 9, shape: '2', paragraphs: ['x'] })).rejects.toBeInstanceOf(EditError);
    await expect(doc.apply({ type: 'setTableCellText', slide: 3, shape: '5', row: 9, col: 0, paragraphs: ['x'] })).rejects.toThrow('That table cell no longer exists.');
    await expect(doc.apply({ type: 'setTableCellText', slide: 3, shape: '2', row: 0, col: 0, paragraphs: ['x'] })).rejects.toThrow('That shape is not a table.');
  });
});

describe('slide operations', () => {
  it('duplicates a slide with its notes', async () => {
    const original = sample();
    const doc = await PresentationDocument.open(original);
    const applied = await doc.apply({ type: 'duplicateSlide', slide: 0 });
    expect(applied.activate).toBe(1);
    const saved = await doc.save();
    expect(changedEntries(original, saved)).toEqual([
      '[Content_Types].xml', 'ppt/_rels/presentation.xml.rels', 'ppt/notesSlides/_rels/notesSlide2.xml.rels', 'ppt/notesSlides/notesSlide2.xml',
      'ppt/presentation.xml', 'ppt/slides/_rels/slide6.xml.rels', 'ppt/slides/slide6.xml',
    ]);
    const { slides } = await presentation(saved);
    expect(slides).toHaveLength(6);
    expect(textOf(shapes(slides[1])[0])).toEqual(['Office Workbench']);
    expect(slides[1].notes).toBe('タイトルスライドのノートです。');
    expect(await part(saved, 'ppt/slides/_rels/slide6.xml.rels')).toContain('Target="../notesSlides/notesSlide2.xml"');
    expect(await part(saved, 'ppt/notesSlides/_rels/notesSlide2.xml.rels')).toContain('Target="../slides/slide6.xml"');
    expect(await part(saved, '[Content_Types].xml')).toContain('<Override PartName="/ppt/slides/slide6.xml"');
    expect(await part(saved, 'ppt/presentation.xml')).toContain('<p:sldId id="256" r:id="rId2"/><p:sldId id="261" r:id="rId');

    // The copy is independent of the original.
    const copy = await PresentationDocument.open(saved);
    await copy.apply({ type: 'setShapeText', slide: 1, shape: '2', paragraphs: ['Copy'] });
    const edited = await presentation(await copy.save());
    expect([textOf(shapes(edited.slides[0])[0]), textOf(shapes(edited.slides[1])[0])]).toEqual([['Office Workbench'], ['Copy']]);
  });

  it('deletes a slide with its notes and moves slides', async () => {
    const original = sample();
    const doc = await PresentationDocument.open(original);
    await doc.apply({ type: 'deleteSlide', slide: 0 });
    let saved = await doc.save();
    let { slides } = await presentation(saved);
    expect(slides.map((s) => textOf(shapes(s)[0])[0])).toEqual(['箇条書き', '図形', '表とグラフ', '非表示スライド']);
    const zip = await JSZip.loadAsync(saved);
    expect(zip.file('ppt/slides/slide1.xml')).toBeNull();
    expect(zip.file('ppt/notesSlides/notesSlide1.xml')).toBeNull();
    expect(await part(saved, '[Content_Types].xml')).not.toMatch(/slide1\.xml|notesSlide1/);
    expect(await part(saved, 'ppt/_rels/presentation.xml.rels')).not.toContain('slides/slide1.xml');

    const moved = await doc.apply({ type: 'moveSlide', slide: 3, to: 0 });
    expect(moved.activate).toBe(0);
    saved = await doc.save();
    ({ slides } = await presentation(saved));
    expect(slides.map((s) => textOf(shapes(s)[0])[0])).toEqual(['非表示スライド', '箇条書き', '図形', '表とグラフ']);
    await expect(doc.apply({ type: 'moveSlide', slide: 0, to: 4 })).rejects.toThrow(/outside/);
  });

  it('refuses to delete the last slide or slides other slides link to', async () => {
    const linked = await JSZip.loadAsync(sample());
    const rels = 'ppt/slides/_rels/slide3.xml.rels';
    const xml = (await linked.file(rels)!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide2.xml"/></Relationships>',
    );
    linked.file(rels, xml, { createFolders: false });
    const doc = await PresentationDocument.open(await linked.generateAsync({ type: 'uint8array' }));
    await expect(doc.apply({ type: 'deleteSlide', slide: 1 })).rejects.toThrow(/links to this slide/);
    for (const slide of [4, 3, 2, 1]) await doc.apply({ type: 'deleteSlide', slide });
    await expect(doc.apply({ type: 'deleteSlide', slide: 0 })).rejects.toThrow('A presentation must keep at least one slide.');
  });
});

describe('PresentationSession', () => {
  it('sends updates without media and undoes text and slide edits', async () => {
    const original = sample();
    const session = new PresentationSession('sample.pptx', original);
    type Message = Extract<HostMessage, { type: 'presentation' }>;
    const [first] = (await session.handle({ type: 'ready' })) as Message[];
    expect(first.media).toHaveLength(1);
    expect(first.update).toBe(false);

    const text = await session.edit({ type: 'setShapeText', slide: 2, shape: '2', paragraphs: ['図形の例'] });
    const [updated] = text.messages as Message[];
    expect(updated).toMatchObject({ update: true, activate: 2, media: [] });
    expect(updated.presentation.media).toHaveLength(1);
    const duplicate = await session.edit({ type: 'duplicateSlide', slide: 2 });
    expect((duplicate.messages[0] as Message).presentation.slides).toHaveLength(6);

    const [afterUndo] = (await duplicate.undo()) as Message[];
    expect(afterUndo.presentation.slides).toHaveLength(5);
    const [textUndone] = (await text.undo()) as Message[];
    expect(textOf(shapeById(textUndone.presentation.slides[2], '2'))).toEqual(['図形']);
    expect(await toMarkdown('pptx', await session.save())).toBe(await toMarkdown('pptx', original));

    await text.redo();
    await duplicate.redo();
    const { slides } = await presentation(await session.save());
    expect(slides.map((s) => textOf(shapes(s)[0])[0])).toEqual(['Office Workbench', '箇条書き', '図形の例', '図形の例', '表とグラフ', '非表示スライド']);
  });
});
