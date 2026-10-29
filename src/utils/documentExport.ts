import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { createSolutionWorkbook } from '../api/excelExporter';
import { createRenderedDocumentationPdf } from '../api/pdfExporter';
import { buildStandaloneHtmlDocument, renderMarkdownDocument, svgMarkupToPngDataUrl } from '../api/documentRendering';
import type { DocumentContext, DocumentationSettings } from '../generator/markdownGenerator';
import type { ParsedSolution } from '../types/solution';

export type ExportFormat = 'md' | 'html' | 'pdf' | 'xlsx';

export interface ExportDocument {
  markdown: string;
  fileName: string;
  solution?: ParsedSolution;
  isReport?: boolean;
  folder?: string;
}

export function safeExportName(name: string): string {
  const cleaned = Array.from(name).filter((character) => (character.codePointAt(0) ?? 0) >= 32)
    .join('').replace(/[<>:"/\\|?*]/g, '_').replace(/[. ]+$/g, '').trim();
  return cleaned && !/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(cleaned) ? cleaned.slice(0, 180) : 'Document';
}

export async function createDocumentExport(
  document: ExportDocument,
  format: ExportFormat,
  context?: DocumentContext,
  settings?: DocumentationSettings,
  signal?: AbortSignal,
  onProgress?: (message: string) => void,
): Promise<Blob> {
  signal?.throwIfAborted();
  if (format === 'md') return new Blob([document.markdown], { type: 'text/markdown;charset=utf-8' });
  if (format === 'xlsx') {
    if (!document.solution) throw new Error('A parsed solution is required for structured Excel export.');
    return createSolutionWorkbook(document.solution, context, { generateMissingDependencies: settings?.generateMissingDependencies, documentationSettings: settings });
  }
  const title = document.fileName.replace(/\.md$/i, '');
  onProgress?.('Rendering documentation and diagrams...');
  const rendered = await renderMarkdownDocument(document.markdown, settings?.mermaidColourTheme);
  signal?.throwIfAborted();
  if (format === 'html') return new Blob([buildStandaloneHtmlDocument(title, rendered.bodyHtml)], { type: 'text/html;charset=utf-8' });
  const images: Array<string | null> = [];
  onProgress?.('Preparing diagram images for PDF...');
  for (const svg of rendered.diagramSvgs) {
    signal?.throwIfAborted();
    images.push(svg ? await svgMarkupToPngDataUrl(svg) : null);
  }
  onProgress?.('Drawing searchable PDF pages...');
  return createRenderedDocumentationPdf(document.markdown, title, images);
}

export async function downloadDocument(document: ExportDocument, format: ExportFormat, context?: DocumentContext, settings?: DocumentationSettings, signal?: AbortSignal, onProgress?: (message: string) => void): Promise<void> {
  const blob = await createDocumentExport(document, format, context, settings, signal, onProgress);
  signal?.throwIfAborted();
  saveAs(blob, `${safeExportName(document.fileName.replace(/\.md$/i, ''))}.${format}`);
}

export async function downloadDocumentBatch(
  documents: ExportDocument[], format: ExportFormat, context?: DocumentContext, settings?: DocumentationSettings,
  signal?: AbortSignal, onProgress?: (completed: number, total: number) => void,
): Promise<void> {
  const zip = new JSZip();
  const usedNames = new Set<string>();
  for (const [index, document] of documents.entries()) {
    signal?.throwIfAborted();
    const actualFormat = format === 'xlsx' && document.isReport ? 'md' : format;
    const folder = document.folder ? document.folder.split('/').map(safeExportName).join('/') + '/' : '';
    const baseName = folder + safeExportName(document.fileName.replace(/\.md$/i, ''));
    let name = `${baseName}.${actualFormat}`;
    let suffix = 2;
    while (usedNames.has(name.toLowerCase())) name = `${baseName} (${suffix++}).${actualFormat}`;
    usedNames.add(name.toLowerCase());
    const blob = await createDocumentExport(document, actualFormat, context, settings, signal);
    zip.file(name, blob);
    onProgress?.(index + 1, documents.length);
  }
  const archive = await zip.generateAsync({ type: 'blob' });
  signal?.throwIfAborted();
  saveAs(archive, `pp-md-${format}-documents.zip`);
}