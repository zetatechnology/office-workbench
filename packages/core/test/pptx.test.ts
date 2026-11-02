import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parsePresentation, type ShapeElement, type Slide, type TableElement } from '../src';

const fixture = () => readFileSync(new URL('../../../test-fixtures/sample.pptx', import.meta.url));

const texts = (slide: Slide) =>
  slide.elements.flatMap((e) => (e.type === 'shape' && e.text ? [e.text.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n')] : []));

const shapeWithText = (slide: Slide, text: string) =>
  slide.elements.find((e): e is ShapeElement => e.type === 'shape' && !!e.text?.paragraphs.some((p) => p.runs.some((r) => r.text.includes(text))));

describe('parsePresentation (PowerPoint-generated fixture)', () => {
  it('reads slide size, order, visibility and notes', async () => {
    const { presentation } = await parsePresentation(fixture());
    expect(presentation.width).toBe(1280);
    expect(presentation.height).toBe(720);
    expect(presentation.slides).toHaveLength(5);
    expect(presentation.slides.map((s) => s.hidden)).toEqual([false, false, false, false, true]);
    expect(presentation.slides[0].notes).toBe('タイトルスライドのノートです。');
    expect(texts(presentation.slides[0])).toEqual(['Office Workbench', 'PowerPoint 表示テスト']);
  });

  it('inherits placeholder geometry and text styles from layout and master', async () => {
    const { presentation } = await parsePresentation(fixture());
    const title = shapeWithText(presentation.slides[0], 'Office Workbench')!;
    expect(title.box.w).toBeGreaterThan(900);
    expect(title.text!.paragraphs[0].align).toBe('ctr');
    expect(title.text!.paragraphs[0].runs[0].size).toBe(80); // 60 pt title from the master.

    const body = shapeWithText(presentation.slides[1], '第一レベル')!;
    const [first, second, third] = body.text!.paragraphs;
    expect(first.bullet).toBe('•');
    expect(second.level).toBe(1);
    expect(second.marginLeft).toBeGreaterThan(first.marginLeft);
    expect(third.runs[0].bold).toBe(true);
  });

  it('resolves shape fills, lines and style colors', async () => {
    const { presentation } = await parsePresentation(fixture());
    const slide = presentation.slides[2];
    const rect = shapeWithText(slide, '四角形')!;
    expect(rect.fill.type).toBe('solid');
    expect(rect.line?.color).toMatch(/^#/);
    expect(rect.text!.paragraphs[0].runs[0].color).toBe('#ffffff');
    expect(shapeWithText(slide, 'グラデーション')!.fill.type).toBe('gradient');
    const ovals = slide.elements.filter((e): e is ShapeElement => e.type === 'shape' && e.fill.type === 'solid' && e.fill.color === '#ff9933');
    expect(ovals).toHaveLength(1);
    const arrowLine = slide.elements.find((e): e is ShapeElement => e.type === 'shape' && e.line?.tailEnd === 'triangle');
    expect(arrowLine?.closed).toBe(false);
    const star = slide.elements.find((e) => e.box.rot === 20);
    expect(star).toBeDefined();
    expect(shapeWithText(slide, 'テキストボックス')!.text!.paragraphs[0]).toMatchObject({ align: 'ctr' });
  });

  it('collects pictures as media', async () => {
    const { presentation, media } = await parsePresentation(fixture());
    const pic = presentation.slides[2].elements.find((e) => e.type === 'picture');
    expect(pic && pic.type === 'picture' && pic.src).toBe('ppt/media/image1.png');
    expect(presentation.media).toEqual([{ path: 'ppt/media/image1.png', contentType: 'image/png' }]);
    expect(media.get('ppt/media/image1.png')?.byteLength).toBeGreaterThan(100);
  });

  it('applies table styles', async () => {
    const { presentation } = await parsePresentation(fixture());
    const table = presentation.slides[3].elements.find((e): e is TableElement => e.type === 'table')!;
    expect(table.columns).toHaveLength(3);
    const [header, band1, band2] = table.rows;
    expect(header.cells[0].text.paragraphs[0].runs[0]).toMatchObject({ text: '地域', bold: true, color: '#ffffff' });
    expect(header.cells[0].fill.type).toBe('solid');
    expect(band1.cells[0].fill).not.toEqual(band2.cells[0].fill);
  });

  it('marks charts as unsupported content', async () => {
    const zip = await JSZip.loadAsync(fixture());
    const slidePath = 'ppt/slides/slide4.xml';
    const xml = await zip.file(slidePath)!.async('string');
    const chartFrame =
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="99" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="6350000" y="1778000"/><a:ext cx="5080000" cy="3810000"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></p:graphicFrame>';
    zip.file(slidePath, xml.replace('</p:spTree>', `${chartFrame}</p:spTree>`));
    const { presentation } = await parsePresentation(await zip.generateAsync({ type: 'uint8array' }));
    const chart = presentation.slides[3].elements.find((e) => e.type === 'unsupported');
    expect(chart).toMatchObject({ type: 'unsupported', label: 'Chart', box: { x: 6350000 / 9525, w: 5080000 / 9525 } });
  });
});
