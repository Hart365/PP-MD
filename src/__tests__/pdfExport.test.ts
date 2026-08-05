import {
  buildPdfDocumentHtml,
  optimizeRenderedHtmlForPdf,
  sanitizeHtmlForPdfExport,
  toPdfFileName,
} from '../utils/pdfExport';

describe('pdfExport utilities', () => {
  it('converts markdown and generic names to .pdf extension', () => {
    expect(toPdfFileName('Solution.md')).toBe('Solution.pdf');
    expect(toPdfFileName('Solution')).toBe('Solution.pdf');
    expect(toPdfFileName('Solution.PDF')).toBe('Solution.PDF');
    expect(toPdfFileName('   ')).toBe('PP-MD Document.pdf');
  });

  it('removes scriptable HTML from exported payload', () => {
    const html = [
      '<h2 onclick="alert(1)">Heading</h2>',
      '<script>alert(1)</script>',
      '<a href="javascript:alert(1)">bad link</a>',
      '<iframe src="https://example.com"></iframe>',
      '<p>Safe</p>',
    ].join('');

    const sanitized = sanitizeHtmlForPdfExport(html);

    expect(sanitized).not.toMatch(/script|onclick|javascript:|iframe/i);
    expect(sanitized).toContain('<p>Safe</p>');
  });

  it('builds a printable HTML shell with provided content', () => {
    const html = buildPdfDocumentHtml({
      title: 'Contoso Report',
      bodyHtml: '<h1>Contoso Report</h1><p>Body</p>',
      language: 'en',
    });

    expect(html).toContain('<html lang="en">');
    expect(html).toContain('<title>Contoso Report</title>');
    expect(html).toContain('<main aria-label="Contoso Report">');
    expect(html).toContain('<p>Body</p>');
    expect(html).toContain('button,');
    expect(html).toContain('display: none !important;');
  });

  it('lands wide tables to landscape and removes scroll wrappers during PDF optimization', () => {
    const longText = 'This is a very long cell value intended to exceed the wide table text threshold for PDF splitting behavior.';
    const html = [
      '<div role="region" aria-label="Table" style="overflow-x:auto">',
      '<table>',
      '<thead><tr><th>Display Name</th><th>Schema</th><th>C3</th><th>C4</th><th>C5</th><th>C6</th><th>C7</th><th>C8</th><th>C9</th></tr></thead>',
      `<tbody><tr><td>Project Name</td><td>new_name</td><td>${longText}</td><td>b</td><td>c</td><td>d</td><td>e</td><td>f</td><td>g</td></tr></tbody>`,
      '</table>',
      '</div>',
    ].join('');

    const optimized = optimizeRenderedHtmlForPdf(html);

    expect(optimized).not.toContain('overflow-x:auto');
    expect(optimized).toContain('pdf-landscape-page');
    expect((optimized.match(/<table/g) || []).length).toBeGreaterThanOrEqual(1);
  });

  it('still applies landscape wrapping on very large rendered HTML', () => {
    const filler = `<p>${'x'.repeat(1_200_000)}</p>`;
    const html = [
      filler,
      '<div role="region" aria-label="Table" style="overflow-x:auto">',
      '<table>',
      '<thead><tr><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th><th>G</th></tr></thead>',
      '<tbody><tr><td>1</td><td>2</td><td>3</td><td>4</td><td>5</td><td>6</td><td>7</td></tr></tbody>',
      '</table>',
      '</div>',
    ].join('');

    const optimized = optimizeRenderedHtmlForPdf(html);

    expect(optimized).not.toContain('overflow-x:auto');
    expect(optimized).toContain('pdf-landscape-page');
  });

  it('wraps mermaid figures without throwing hierarchy errors', () => {
    const html = [
      '<figure data-pdf-mermaid="true">',
      '<svg data-pdf-mermaid="true" viewBox="0 0 1600 900" xmlns="http://www.w3.org/2000/svg">',
      '<text x="20" y="20">Diagram label</text>',
      '<rect x="30" y="40" width="1200" height="700" />',
      '</svg>',
      '</figure>',
    ].join('');

    expect(() => optimizeRenderedHtmlForPdf(html)).not.toThrow();
    const optimized = optimizeRenderedHtmlForPdf(html);
    expect(optimized).toContain('pdf-diagram-page');
    expect(optimized).toContain('pdf-landscape-page');
    expect(optimized).toContain('data:image/svg+xml');
  });

  it('encodes Mermaid SVGs as base64 data URIs so Unicode labels remain reliable in PDF export', () => {
    const html = [
      '<figure data-pdf-mermaid="true">',
      '<svg data-pdf-mermaid="true" viewBox="0 0 1200 900" xmlns="http://www.w3.org/2000/svg">',
      '<text x="20" y="40">Café — Résumé</text>',
      '<rect x="30" y="60" width="1000" height="700" />',
      '</svg>',
      '</figure>',
    ].join('');

    const optimized = optimizeRenderedHtmlForPdf(html);

    expect(optimized).toContain('data:image/svg+xml;base64,');
    expect(optimized).not.toContain('data:image/svg+xml;charset=utf-8,');
  });

  it('keeps regular SVG figures intact instead of converting them to print-safe images', () => {
    const html = [
      '<figure>',
      '<svg viewBox="0 0 1200 4200" xmlns="http://www.w3.org/2000/svg">',
      '<text x="20" y="20">Regular diagram</text>',
      '<rect x="30" y="40" width="1000" height="4000" />',
      '</svg>',
      '</figure>',
    ].join('');

    const optimized = optimizeRenderedHtmlForPdf(html);

    expect(optimized).not.toContain('data:image/svg+xml');
    expect(optimized).not.toContain('pdf-diagram-slice');
  });

  it('converts sliced tall Mermaid diagrams to print-safe image payloads', () => {
    const html = [
      '<figure data-pdf-mermaid="true">',
      '<svg data-pdf-mermaid="true" viewBox="0 0 1200 4200" xmlns="http://www.w3.org/2000/svg">',
      '<text x="20" y="20">Diagram label</text>',
      '<rect x="30" y="40" width="1000" height="4000" />',
      '</svg>',
      '</figure>',
    ].join('');

    const optimized = optimizeRenderedHtmlForPdf(html);

    expect(optimized).toContain('data:image/svg+xml');
    expect(optimized).toContain('pdf-diagram-slice');
    expect((optimized.match(/Diagram part /g) || []).length).toBeGreaterThan(1);
  });

  it('enforces minimum 8pt text size for Mermaid diagrams and splits tall diagrams across pages', () => {
    const html = [
      '<figure data-pdf-mermaid="true">',
      '<svg data-pdf-mermaid="true" viewBox="0 0 1200 4200" xmlns="http://www.w3.org/2000/svg">',
      '<text x="20" y="20" font-size="6pt">Too small</text>',
      '<rect x="30" y="40" width="1000" height="4000" />',
      '</svg>',
      '</figure>',
    ].join('');

    const optimized = optimizeRenderedHtmlForPdf(html);

    expect(optimized).toContain('data:image/svg+xml;base64,');
    expect(optimized).toContain('pdf-diagram-slice');
    expect((optimized.match(/Diagram part /g) || []).length).toBeGreaterThan(1);
  });
});
