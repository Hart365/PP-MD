import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { createRenderedDocumentationPdf } from './pdfExporter';

const SAMPLE_MARKDOWN = `# Sample Solution Documentation

## Overview

This solution contains **bold text**, _italic text_, and \`inline code\` alongside a [link](https://example.test).

## Tables & Relationships

| Name | Type | Required |
| --- | --- | --- |
| accountid | Uniqueidentifier | Yes |
| name | String | Yes |

## Notes

- First bullet point
- Second bullet point with a nested list:
  - Nested item one
  - Nested item two

> A blockquote with **emphasis** inside it.

\`\`\`json
{ "example": true }
\`\`\`

\`\`\`mermaid
erDiagram
  ACCOUNT ||--o{ CONTACT : has
\`\`\`
`;

describe('Markdown-to-PDF export', () => {
  it('produces a real, non-empty PDF blob from Markdown source (not a screenshot)', async () => {
    const blob = await createRenderedDocumentationPdf(SAMPLE_MARKDOWN, 'Sample Solution');

    expect(blob.type).toBe('application/pdf');
    expect(blob.size).toBeGreaterThan(1000);

    const header = new TextDecoder().decode(new Uint8Array(await blob.slice(0, 5).arrayBuffer()));
    expect(header).toBe('%PDF-');
  });

  it('embeds real vector fonts (not a rasterized screenshot) for the main content pages', async () => {
    const blob = await createRenderedDocumentationPdf(SAMPLE_MARKDOWN, 'Sample Solution');
    const raw = await blob.text();

    // Real fonts are referenced when text is drawn with jsPDF's text API; a
    // pure-image (screenshot) PDF would only ever reference an /Image XObject.
    expect(raw).toMatch(/\/BaseFont \/Helvetica/);
    expect(raw).not.toMatch(/\/Subtype \/Image/);
  });

  it('removes complete HTML tags when attributes contain greater-than signs', async () => {
    const markdown = '<table>\n<tr><td><span title="2 > 1">Clean</span></td></tr>\n</table>';
    const blob = await createRenderedDocumentationPdf(markdown, 'HTML table');
    const bytes = new Uint8Array(await blob.arrayBuffer());
    const raw = Array.from(bytes, (byte) => String.fromCharCode(byte)).join('');
    const compressedStreams = [...raw.matchAll(/stream\r?\n([\s\S]*?)\r?\nendstream/g)];
    const pageText = await Promise.all(compressedStreams.map(async (match) => {
      const compressed = Uint8Array.from(match[1], (character) => character.charCodeAt(0));
      const content = inflateSync(compressed);
      return Array.from(content, (byte) => String.fromCharCode(byte)).join('');
    })).then((streams) => streams.join('\n'));

    expect(pageText).toContain('(Clean)');
    expect(pageText).not.toContain('1">Clean');
  });

  it('falls back to rendering the raw diagram source when no pre-rendered image is supplied', async () => {
    const blob = await createRenderedDocumentationPdf(SAMPLE_MARKDOWN, 'Sample Solution', []);
    expect(blob.type).toBe('application/pdf');
    expect(blob.size).toBeGreaterThan(1000);
  });

  it('embeds a supplied diagram image without throwing', async () => {
    // 1x1 transparent PNG.
    const onePixelPng = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
    const blob = await createRenderedDocumentationPdf(SAMPLE_MARKDOWN, 'Sample Solution', [onePixelPng]);
    expect(blob.type).toBe('application/pdf');
    expect(blob.size).toBeGreaterThan(1000);
  });

  it('handles very large documents by paginating instead of producing a blank/oversized canvas', async () => {
    const bigSection = Array.from({ length: 400 }, (_, i) => `## Section ${i}\n\nParagraph text for section ${i} describing an entity in detail.\n`).join('\n');
    const blob = await createRenderedDocumentationPdf(`# Big Doc\n\n${bigSection}`, 'All Selected Solutions');
    expect(blob.type).toBe('application/pdf');
    expect(blob.size).toBeGreaterThan(5000);
  });
});
