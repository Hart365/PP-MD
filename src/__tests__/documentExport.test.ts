import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { createDocumentExport, downloadDocumentBatch, safeExportName } from '../utils/documentExport';
import { renderMarkdownDocument, svgToPngDataUrl } from '../api/documentRendering';

vi.mock('file-saver', () => ({ saveAs: vi.fn() }));

describe('portable document exports', () => {
  it('falls back instead of hanging when canvas conversion throws', async () => {
    const image = { onload: null as (() => void) | null, onerror: null as (() => void) | null };
    Object.defineProperty(image, 'src', { set: () => queueMicrotask(() => image.onload?.()) });
    vi.stubGlobal('Image', class { constructor() { return image; } });
    const context = vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => { throw new Error('Canvas unavailable'); });
    try {
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('viewBox', '0 0 100 100');
      await expect(svgToPngDataUrl(svg)).resolves.toBeNull();
    } finally {
      context.mockRestore();
      vi.unstubAllGlobals();
    }
  });
  it('sanitizes imported HTML and produces a standalone offline document', async () => {
    const markdown = '# Example\n\n<script>alert(1)</script>\n\n<img src="x" onerror="alert(2)">\n\n[Unsafe](javascript:alert(3))\n\n| Name | Value |\n| --- | --- |\n| Item | Text |';
    const rendered = await renderMarkdownDocument(markdown);
    expect(rendered.bodyHtml).not.toContain('<script');
    expect(rendered.bodyHtml).not.toContain('onerror');
    expect(rendered.bodyHtml).not.toContain('href="javascript:');
    const exported = await createDocumentExport({ markdown, fileName: 'Example.md' }, 'html');
    const html = await new Promise<string>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.readAsText(exported);
    });
    expect(html).toContain('<html lang="en-GB">');
    expect(html).toContain('<table>');
    expect(html).not.toContain('<script');
  });

  it('avoids overwriting duplicate archive names and keeps category folders safe', async () => {
    await downloadDocumentBatch([
      { markdown: '# First', fileName: 'Same.md', folder: '../Example' },
      { markdown: '# Second', fileName: 'same.md', folder: '../Example' },
    ], 'md');
    const archive = vi.mocked(saveAs).mock.calls.at(-1)?.[0] as Blob;
    const zip = await JSZip.loadAsync(archive);
    const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
    expect(names).toHaveLength(2);
    expect(names.every((name) => !name.startsWith('../'))).toBe(true);
    expect(names.some((name) => name.includes('(2)'))).toBe(true);
    expect(safeExportName('CON')).toBe('Document');
    expect(safeExportName('bad/name')).toBe('bad_name');
  });

  it('does not download an archive after cancellation', async () => {
    vi.mocked(saveAs).mockClear();
    const controller = new AbortController();
    controller.abort();
    await expect(downloadDocumentBatch([{ markdown: '# Example', fileName: 'Example.md' }], 'md', undefined, undefined, controller.signal)).rejects.toMatchObject({ name: 'AbortError' });
    expect(saveAs).not.toHaveBeenCalled();
  });
});