// Presentation → Markdown: each slide's text boxes and tables in shape order, then its notes.
import { parsePresentation } from '../pptx/presentation';
import type { TextBody } from '../pptx/types';
import { escapeCell, markdownTable } from './markdown';

export async function pptxToMarkdown(data: Uint8Array): Promise<string> {
  const { presentation } = await parsePresentation(data);
  const out: string[] = [];
  for (const slide of presentation.slides) {
    out.push(`## Slide ${slide.index + 1}${slide.hidden ? ' (hidden)' : ''}`);
    const blocks: string[] = [];
    for (const element of slide.elements) {
      if (element.type === 'shape' && element.text) {
        const text = bodyText(element.text);
        if (text) blocks.push(text);
      } else if (element.type === 'table') {
        blocks.push(
          markdownTable(element.rows.map((row) => row.cells.map((cell) => (cell.merged ? '' : escapeCell(bodyText(cell.text, false)))))),
        );
      } else if (element.type === 'picture') {
        blocks.push('_[picture]_');
      } else if (element.type === 'unsupported') {
        blocks.push(`_[${element.label}]_`);
      }
    }
    out.push(...(blocks.length ? blocks : ['_(no text)_']));
    if (slide.notes?.trim()) out.push(`> **Notes:** ${slide.notes.trim().replace(/\n+/g, '\n> ')}`);
  }
  return out.join('\n\n') + '\n';
}

/** Paragraphs of a text body; bulleted paragraphs become list items when `lists` is set. */
function bodyText(body: TextBody, lists = true): string {
  return body.paragraphs
    .map((p) => {
      const text = p.runs.map((r) => (r.br ? '\n' : r.text)).join('').trim();
      if (!text) return '';
      return lists && p.bullet ? `${'  '.repeat(p.level)}- ${text}` : text;
    })
    .filter(Boolean)
    .join('\n');
}
