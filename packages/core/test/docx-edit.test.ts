import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { EditError, WordDocument, WordSession, toMarkdown, type DocxBlock, type DocxParagraph, type HostMessage } from '../src';
import { changedEntries, fixture, part } from './helpers';

const sample = () => fixture('sample.docx');
const DOC = 'word/document.xml';

function paragraphs(blocks: DocxBlock[]): DocxParagraph[] {
  return blocks.flatMap((b) =>
    b.type === 'paragraph' ? [b] : b.type === 'table' ? b.rows.flatMap((r) => r.cells.flatMap((c) => paragraphs(c.blocks))) : [],
  );
}

const byText = (blocks: DocxBlock[], text: string) => paragraphs(blocks).find((p) => p.text === text)!;

/** The sample document with its body replaced. */
async function withBody(bodyXml: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(sample());
  const xml = (await zip.file(DOC)!.async('string')).replace(/<w:body>[\s\S]*<w:sectPr/, `<w:body>${bodyXml}<w:sectPr`);
  zip.file(DOC, xml, { createFolders: false });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** The XML of one body paragraph in a saved file. */
async function paragraphXml(data: Uint8Array, index: number): Promise<string> {
  const xml = (await part(data, DOC))!;
  const body = xml.slice(xml.indexOf('<w:body>') + 8);
  const items = body.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p [^>]*\/>|<w:tbl>[\s\S]*?<\/w:tbl>/g)!;
  return items[index];
}

describe('Word outline', () => {
  it('lists paragraphs with styles, lists and tables, and locks what cannot be rewritten', async () => {
    const blocks = await (await WordDocument.open(sample())).blocks();
    expect(blocks[0]).toMatchObject({ type: 'paragraph', id: '0', text: 'Office Workbench サンプル文書', heading: 1 });
    expect(byText(blocks, '一つ目の項目')).toMatchObject({ list: { ordered: false, level: 0 } });
    const table = blocks.find((b) => b.type === 'table')!;
    expect(table.type === 'table' && table.rows.length).toBe(3);
    expect(byText(blocks, '数量').id).toBe(`${table.id}/0/1/1`);
    const locked = paragraphs(blocks).filter((p) => p.locked);
    expect(locked.map((p) => p.locked)).toEqual([
      'Paragraphs with pictures or shapes cannot be edited here yet.',
      'Paragraphs with page or column breaks cannot be edited here yet.',
    ]);
  });

  it('locks fields, hyperlinks, tracked changes, comments, content controls and inner bookmarks', async () => {
    const r = (t: string) => `<w:r><w:t>${t}</w:t></w:r>`;
    const doc = await WordDocument.open(
      await withBody(
        [
          `<w:p>${r('a')}<w:fldSimple w:instr="PAGE">${r('1')}</w:fldSimple></w:p>`,
          `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
          `<w:p><w:hyperlink r:id="rId99">${r('link')}</w:hyperlink></w:p>`,
          `<w:p><w:ins w:id="1" w:author="x">${r('new')}</w:ins></w:p>`,
          `<w:p><w:commentRangeStart w:id="0"/>${r('c')}<w:commentRangeEnd w:id="0"/></w:p>`,
          `<w:p><w:sdt><w:sdtContent>${r('cc')}</w:sdtContent></w:sdt></w:p>`,
          `<w:p>${r('x')}<w:bookmarkStart w:id="5" w:name="b"/>${r('y')}<w:bookmarkEnd w:id="5"/></w:p>`,
          `<w:p><w:bookmarkStart w:id="6" w:name="_Toc1"/>${r('edge')}<w:bookmarkEnd w:id="6"/><w:proofErr w:type="spellStart"/></w:p>`,
          `<w:sdt><w:sdtContent><w:p>${r('block control')}</w:p></w:sdtContent></w:sdt>`,
          `<w:p>${r('last')}</w:p>`,
        ].join(''),
      ),
    );
    const blocks = await doc.blocks();
    const reasons = paragraphs(blocks).map((p) => p.locked?.replace(/ cannot.*| so it.*/, ''));
    expect(reasons).toEqual([
      'Paragraphs with fields (page numbers, references, ...)',
      'Paragraphs with fields (page numbers, references, ...)',
      'Paragraphs with hyperlinks',
      'Paragraphs with tracked changes',
      'Paragraphs with comments',
      'Paragraphs with content controls',
      'A bookmark starts or ends inside this paragraph,',
      undefined,
      undefined,
    ]);
    expect(blocks[8]).toMatchObject({ type: 'other', label: 'Content control', text: 'block control' });
    for (const id of ['0', '2', '3', '6']) {
      await expect(doc.apply({ type: 'setParagraphText', id, text: 'z' })).rejects.toBeInstanceOf(EditError);
    }
    await expect(doc.apply({ type: 'deleteParagraph', id: '7' })).rejects.toThrow(/bookmarks/);

    // A bookmark around the whole text stays around the new text; spelling marks are dropped.
    await doc.apply({ type: 'setParagraphText', id: '7', text: 'renamed' });
    expect(await paragraphXml(await doc.save(), 7)).toBe(
      '<w:p><w:bookmarkStart w:id="6" w:name="_Toc1"/><w:r><w:t>renamed</w:t></w:r><w:bookmarkEnd w:id="6"/></w:p>',
    );
  });
});

describe('editing Word paragraphs', () => {
  it('rewrites one paragraph, keeping its properties and first-run formatting, and leaves other parts byte-identical', async () => {
    const original = sample();
    const doc = await WordDocument.open(original);
    await doc.apply({ type: 'setParagraphText', id: '0', text: '新しい見出し' });
    await doc.apply({ type: 'setParagraphText', id: '1', text: '  Tab\tand\nbreak ' });
    const table = (await doc.blocks()).find((b) => b.type === 'table')!;
    await doc.apply({ type: 'setParagraphText', id: `${table.id}/0/1/1`, text: '個数' });
    const saved = await doc.save();

    expect(changedEntries(original, saved)).toEqual([DOC]);
    expect(await paragraphXml(saved, 0)).toMatch(/^<w:p [^>]*w14:paraId="72BE9AEF"[^>]*><w:pPr><w:pStyle w:val="1"\/><\/w:pPr><w:r><w:t>新しい見出し<\/w:t><\/w:r><\/w:p>$/);
    expect(await paragraphXml(saved, 1)).toContain(
      '<w:r><w:t xml:space="preserve">  Tab</w:t><w:tab/><w:t>and</w:t><w:br/><w:t xml:space="preserve">break </w:t></w:r></w:p>',
    );
    expect(await paragraphXml(saved, 7)).toContain(
      '<w:r w:rsidRPr="000D0BD4"><w:rPr><w:rFonts w:hint="eastAsia"/><w:b/></w:rPr><w:t>項目</w:t></w:r>',
    );
    expect(await paragraphXml(saved, 7)).toContain('<w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:b/></w:rPr><w:t>個数</w:t></w:r>');

    const reopened = await (await WordDocument.open(saved)).blocks();
    expect(reopened[0]).toMatchObject({ text: '新しい見出し', heading: 1 });
    expect(reopened[1]).toMatchObject({ text: '  Tab\tand\nbreak ' });
    expect(await toMarkdown('docx', saved)).toContain('| 項目 | 個数 | 備考 |');
  });

  it('inserts, splits and deletes paragraphs', async () => {
    const doc = await WordDocument.open(sample());
    // A new list item after the first one copies its paragraph properties (style and indent).
    await doc.apply({ type: 'insertParagraph', after: '3', text: '一・五' });
    let blocks = await doc.blocks();
    expect(paragraphs(blocks).slice(3, 6).map((p) => p.text)).toEqual(['一つ目の項目', '一・五', '二つ目の項目']);
    expect(blocks[4]).toMatchObject({ id: '4', list: { ordered: false, level: 0 } });
    expect(await paragraphXml(await doc.save(), 4)).toBe(
      '<w:p><w:pPr><w:pStyle w:val="a"/><w:ind w:left="440" w:hanging="440"/></w:pPr><w:r><w:t>一・五</w:t></w:r></w:p>',
    );

    // Enter in the middle of a paragraph: the text before stays, the rest moves to a new paragraph.
    await doc.apply({ type: 'insertParagraph', after: '1', text: '後半', update: '前半' });
    blocks = await doc.blocks();
    expect([blocks[1], blocks[2]].map((b) => b.type === 'paragraph' && b.text)).toEqual(['前半', '後半']);

    await doc.apply({ type: 'deleteParagraph', id: '2' });
    blocks = await doc.blocks();
    expect(blocks[2]).toMatchObject({ text: '箇条書き' });
  });

  it('refuses to delete the only paragraph between two tables', async () => {
    const table = (text: string) => `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const doc = await WordDocument.open(await withBody(`${table('A')}<w:p/>${table('B')}<w:p><w:r><w:t>end</w:t></w:r></w:p>`));
    await expect(doc.apply({ type: 'deleteParagraph', id: '1' })).rejects.toThrow(/separates two tables/);
  });

  it('refuses edits that would break the document', async () => {
    const doc = await WordDocument.open(sample());
    const blocks = await doc.blocks();
    const table = blocks.find((b) => b.type === 'table')!;
    await expect(doc.apply({ type: 'deleteParagraph', id: `${table.id}/0/0/1` })).rejects.toThrow('The last paragraph of a table cell cannot be deleted.');
    const last = paragraphs(blocks).at(-1)!;
    await expect(doc.apply({ type: 'deleteParagraph', id: last.id })).rejects.toThrow('The last paragraph of the document cannot be deleted.');
    const picture = paragraphs(blocks).find((p) => p.locked)!;
    await expect(doc.apply({ type: 'setParagraphText', id: picture.id, text: 'x' })).rejects.toThrow(/pictures/);
    await expect(doc.apply({ type: 'deleteParagraph', id: picture.id })).rejects.toThrow(/pictures/);
    await expect(doc.apply({ type: 'setParagraphText', id: table.id, text: 'x' })).rejects.toThrow('That paragraph no longer exists.');
    await expect(doc.apply({ type: 'setParagraphText', id: '99', text: 'x' })).rejects.toThrow('That paragraph no longer exists.');
    await expect(doc.apply({ type: 'setParagraphText', id: '1/a', text: 'x' })).rejects.toThrow('is not a paragraph id');
  });

  it('undoes and redoes through the session, and ignores stale undo after a reset', async () => {
    const original = sample();
    const session = new WordSession('sample.docx', original);
    const [initial] = await session.handle({ type: 'ready' });
    const blocksOf = (messages: HostMessage[]) => {
      const [message] = messages;
      if (message.type !== 'docx' && message.type !== 'docxOutline') throw new Error(`unexpected ${message.type}`);
      return message.blocks!;
    };
    expect(initial.type === 'docx' && initial.data).toBe(original);

    const edits = [
      await session.edit({ type: 'setParagraphText', id: '6', text: '表の例:' }),
      await session.edit({ type: 'insertParagraph', after: '6', text: '追加' }),
      await session.edit({ type: 'deleteParagraph', id: '5' }),
    ];
    expect(edits.map((e) => e.label)).toEqual(['Edit Paragraph', 'Insert Paragraph', 'Delete Paragraph']);
    // After an edit only the outline is sent; the page data comes when a view asks for it.
    expect(edits[0].messages.map((m) => m.type)).toEqual(['docxOutline']);
    const [preview] = await session.handle({ type: 'preview' });
    expect(preview.type === 'docx' && (await toMarkdown('docx', preview.data))).toContain('表の例:');
    expect(await session.blocks()).toEqual(blocksOf(edits[2].messages));
    const after = paragraphs(blocksOf(edits[2].messages)).map((p) => p.text);
    expect(after.slice(4, 7)).toEqual(['二つ目の項目', '表の例:', '追加']);

    for (const edit of [...edits].reverse()) await edit.undo();
    const undone = paragraphs(blocksOf(await session.handle({ type: 'ready' })));
    expect(undone).toEqual(paragraphs(await (await WordDocument.open(original)).blocks()));
    // The undone document is the original, apart from serialization of the edited part.
    expect(await toMarkdown('docx', await session.save())).toBe(await toMarkdown('docx', original));

    const redone = await edits[0].redo();
    expect(byText(blocksOf(redone), '表の例:')).toBeDefined();

    session.reset(original);
    expect(await edits[1].redo()).toEqual([]);
    expect(byText(blocksOf(await session.handle({ type: 'ready' })), '表:')).toBeDefined();
  });
});
