/**
 * @file pdfExport.ts
 * @description Renderer-side helpers for exporting rendered markdown content
 * to a tagged PDF through Electron IPC.
 */

import DOMPurify from 'dompurify';

export interface PdfExportRequest {
  title: string;
  fileName: string;
  renderedHtml: string;
  language?: string;
}

export interface PdfExportResult {
  cancelled: boolean;
  filePath?: string;
  error?: string;
}

const WIDE_TABLE_COLUMN_THRESHOLD = 8;
const PORTRAIT_CONTENT_WIDTH_PX = 680;
const LANDSCAPE_CONTENT_WIDTH_PX = 990;
const MIN_DIAGRAM_FONT_PT = 8;
const LANDSCAPE_PAGE_CONTENT_RATIO = 1.45;
const DIAGRAM_SLICE_OVERLAP_UNITS = 120;
const MAX_DIAGRAM_SLICES = 24;
const MIN_DIAGRAM_SLICE_HEIGHT_UNITS = 520;

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

const HTML_ENTITY_DECODE_MAP: Readonly<Record<string, string>> = {
  '&nbsp;': ' ',
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  "&#39;": "'",
};
const HTML_ENTITY_PATTERN = /&(?:nbsp|amp|lt|gt|quot|#39);/gi;

/** Decodes known HTML entities in a single pass to avoid double-unescaping. */
function decodeHtmlEntities(value: string): string {
  return value.replace(HTML_ENTITY_PATTERN, (match) => HTML_ENTITY_DECODE_MAP[match.toLowerCase()] ?? match);
}

function stripHtmlTags(value: string): string {
  return decodeHtmlEntities(
    value
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, ''),
  ).trim();
}

function estimateTextWidthPx(value: string): number {
  const longest = value
    .split('\n')
    .reduce((max, line) => Math.max(max, line.length), 0);
  return Math.min(340, Math.max(40, longest * 6.5));
}

function estimateTableWidthPx(headers: string[], rows: string[][]): number {
  if (headers.length === 0) return 0;
  let width = 0;
  headers.forEach((header, idx) => {
    const maxCell = rows.reduce((max, row) => {
      const value = stripHtmlTags(row[idx] || '');
      return Math.max(max, estimateTextWidthPx(value));
    }, estimateTextWidthPx(header));
    width += Math.min(220, Math.max(78, maxCell + 20));
  });
  return width;
}

function shouldUseLandscapeTable(headers: string[], rows: string[][]): boolean {
  const estimatedWidth = estimateTableWidthPx(headers, rows);
  if (estimatedWidth > PORTRAIT_CONTENT_WIDTH_PX) return true;
  if (headers.length > 6) return true;
  return false;
}

function ptToPx(pt: number): number {
  return pt * (96 / 72);
}

function enforceMinDiagramFont(svg: SVGElement): void {
  svg.querySelectorAll('text').forEach((node) => {
    const existing = node.getAttribute('font-size')?.trim();
    if (!existing) {
      node.setAttribute('font-size', `${MIN_DIAGRAM_FONT_PT}pt`);
      return;
    }

    const numeric = Number.parseFloat(existing);
    if (!Number.isFinite(numeric)) {
      node.setAttribute('font-size', `${MIN_DIAGRAM_FONT_PT}pt`);
      return;
    }

    const lower = existing.toLowerCase();
    const minPx = ptToPx(MIN_DIAGRAM_FONT_PT);
    const valuePx = lower.endsWith('pt')
      ? ptToPx(numeric)
      : lower.endsWith('px')
        ? numeric
        : numeric;

    if (!Number.isFinite(valuePx) || valuePx < minPx) {
      node.setAttribute('font-size', `${MIN_DIAGRAM_FONT_PT}pt`);
    }
  });
}

function isMermaidDiagramSvg(svg: SVGElement): boolean {
  return svg.getAttribute('data-pdf-mermaid') === 'true'
    || svg.closest('figure')?.getAttribute('data-pdf-mermaid') === 'true'
    || svg.closest('[data-pdf-mermaid]')?.getAttribute('data-pdf-mermaid') === 'true';
}

function convertSvgToPrintImage(svg: SVGElement, doc: Document): HTMLImageElement | null {
  if (!isMermaidDiagramSvg(svg)) {
    return null;
  }

  const clone = svg.cloneNode(true) as SVGElement;
  clone.removeAttribute('width');
  clone.removeAttribute('height');

  const viewBox = clone.getAttribute('viewBox')?.trim().split(/\s+/) ?? [];
  if (viewBox.length === 4) {
    const [, , width, height] = viewBox.map((value) => Number(value));
    if (Number.isFinite(width) && width > 0 && Number.isFinite(height) && height > 0) {
      clone.setAttribute('width', String(width));
      clone.setAttribute('height', String(height));
    }
  }

  if (!clone.getAttribute('xmlns')) {
    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  }

  const serialized = new XMLSerializer().serializeToString(clone);
  const encoded = typeof btoa === 'function'
    ? (() => {
        const utf8Bytes = typeof TextEncoder !== 'undefined'
          ? new TextEncoder().encode(serialized)
          : Uint8Array.from(serialized, (char) => char.charCodeAt(0));

        let binary = '';
        utf8Bytes.forEach((byte) => {
          binary += String.fromCharCode(byte);
        });
        return btoa(binary);
      })()
    : serialized;

  const image = doc.createElement('img');
  image.setAttribute('src', `data:image/svg+xml;base64,${encoded}`);
  image.setAttribute('alt', clone.getAttribute('aria-label') || 'Diagram');
  image.setAttribute('role', 'img');
  image.style.display = 'block';
  image.style.width = '100%';
  image.style.height = 'auto';
  image.style.maxWidth = '100%';
  image.style.objectFit = 'contain';
  return image;
}

function splitWideTable(table: HTMLTableElement, doc: Document): HTMLElement[] {
  const headerRow = table.querySelector('thead tr') ?? table.querySelector('tr');
  if (!headerRow) return [table];

  const headerCells = Array.from(headerRow.querySelectorAll('th,td')).map((cell) => stripHtmlTags(cell.innerHTML));
  const bodyRows = Array.from(table.querySelectorAll('tbody tr'));
  const bodyCells = bodyRows.map((row) => Array.from(row.querySelectorAll('td,th')).map((cell) => cell.innerHTML || ''));
  const forcedLandscape = table.getAttribute('data-pdf-overflow') === '1';
  const tableNeedsLandscape = forcedLandscape || shouldUseLandscapeTable(headerCells, bodyCells);

  if (headerCells.length <= WIDE_TABLE_COLUMN_THRESHOLD) {
    if (!tableNeedsLandscape) return [table];

    const container = doc.createElement('section');
    container.className = 'pdf-landscape-page pdf-wide-table';
    container.appendChild(table.cloneNode(true));
    return [container];
  }

  const estimatedWidth = estimateTableWidthPx(headerCells, bodyCells);
  const shouldSplit = estimatedWidth > LANDSCAPE_CONTENT_WIDTH_PX;
  if (!shouldSplit) {
    const container = doc.createElement('section');
    container.className = 'pdf-landscape-page pdf-wide-table';
    container.appendChild(table.cloneNode(true));
    return [container];
  }

  const keyHeaderIndexes = headerCells
    .map((header, index) => ({ header, index }))
    .filter((entry) => /(display name|name|entity|table|role|schema|logical)/i.test(entry.header))
    .map((entry) => entry.index);

  const keyIndexes = (keyHeaderIndexes.length > 0 ? keyHeaderIndexes : [0]).slice(0, 2);
  const nonKeyIndexes = headerCells
    .map((_, index) => index)
    .filter((index) => !keyIndexes.includes(index));

  const columnsPerChunk = Math.max(1, WIDE_TABLE_COLUMN_THRESHOLD - keyIndexes.length);
  const chunkedIndexes: number[][] = [];
  for (let i = 0; i < nonKeyIndexes.length; i += columnsPerChunk) {
    chunkedIndexes.push([...keyIndexes, ...nonKeyIndexes.slice(i, i + columnsPerChunk)]);
  }

  const chunks = chunkedIndexes.map((indexes, chunkIndex) => {
    const tableClone = table.cloneNode(false) as HTMLTableElement;
    const thead = doc.createElement('thead');
    const tr = doc.createElement('tr');

    indexes.forEach((idx) => {
      const th = doc.createElement('th');
      th.textContent = headerCells[idx] || `Column ${idx + 1}`;
      tr.appendChild(th);
    });
    thead.appendChild(tr);
    tableClone.appendChild(thead);

    const tbody = doc.createElement('tbody');
    bodyRows.forEach((sourceRow) => {
      const sourceCells = Array.from(sourceRow.querySelectorAll('td,th'));
      const row = doc.createElement('tr');
      indexes.forEach((idx) => {
        const td = doc.createElement('td');
        td.innerHTML = sourceCells[idx]?.innerHTML || '';
        row.appendChild(td);
      });
      tbody.appendChild(row);
    });
    tableClone.appendChild(tbody);

    const container = doc.createElement('section');
    container.className = 'pdf-landscape-page pdf-wide-table';
    if (chunkedIndexes.length > 1) {
      const label = doc.createElement('p');
      label.className = 'pdf-split-note';
      label.textContent = `Table split part ${chunkIndex + 1} of ${chunkedIndexes.length}`;
      container.appendChild(label);
    }
    container.appendChild(tableClone);
    return container;
  });

  return chunks;
}

function splitDiagramFigure(figure: HTMLElement, doc: Document): HTMLElement[] {
  const figureClone = figure.cloneNode(true) as HTMLElement;
  const svg = figureClone.querySelector('svg');
  if (!svg) return [figure];

  const isMermaidDiagram = isMermaidDiagramSvg(svg as SVGElement);
  if (!isMermaidDiagram) {
    const wrapped = doc.createElement('section');
    wrapped.className = 'pdf-diagram-page';
    wrapped.appendChild(figureClone);
    return [wrapped];
  }

  enforceMinDiagramFont(svg as SVGElement);
  const printSafeFigure = figureClone.cloneNode(true) as HTMLElement;
  const printSafeSvg = printSafeFigure.querySelector('svg');
  if (printSafeSvg) {
    const printImage = convertSvgToPrintImage(printSafeSvg as SVGElement, doc);
    if (printImage) {
      printSafeSvg.replaceWith(printImage);
    }
  }

  const viewBox = svg.getAttribute('viewBox')?.trim().split(/\s+/).map((part) => Number(part)) ?? [];
  const viewBoxWidth = viewBox.length === 4 && Number.isFinite(viewBox[2]) ? viewBox[2] : NaN;
  const viewBoxHeight = viewBox.length === 4 && Number.isFinite(viewBox[3]) ? viewBox[3] : NaN;
  const widthAttr = Number.parseFloat(svg.getAttribute('width') || '');
  const heightAttr = Number.parseFloat(svg.getAttribute('height') || '');
  const inlineWidth = Number.parseFloat((svg as SVGElement).style.width || '');
  const inlineHeight = Number.parseFloat((svg as SVGElement).style.height || '');

  const width = Number.isFinite(viewBoxWidth)
    ? viewBoxWidth
    : Number.isFinite(inlineWidth) ? inlineWidth : widthAttr;
  const height = Number.isFinite(viewBoxHeight)
    ? viewBoxHeight
    : Number.isFinite(inlineHeight) ? inlineHeight : heightAttr;
  const ratio = Number.isFinite(width) && Number.isFinite(height) && height > 0
    ? width / height
    : Number.NaN;

  // Prefer a single landscape figure wrapper over slicing, which can create
  // clipped labels and excessive whitespace in Chromium PDF output.
  const shouldLandscape = Number.isFinite(width) && Number.isFinite(height)
    ? width > PORTRAIT_CONTENT_WIDTH_PX || ratio >= 1.05
    : true;

  svg.setAttribute('preserveAspectRatio', 'xMidYMin meet');
  svg.removeAttribute('width');
  svg.removeAttribute('height');

  const finalViewBox = svg.getAttribute('viewBox')?.trim().split(/\s+/).map((part) => Number(part)) ?? [];
  const canSliceByViewBox = finalViewBox.length === 4
    && Number.isFinite(finalViewBox[2])
    && Number.isFinite(finalViewBox[3])
    && finalViewBox[2] > 0
    && finalViewBox[3] > 0;

  if (canSliceByViewBox) {
    const [vbX, vbY, vbWidth, vbHeight] = finalViewBox;
    const maxSliceHeight = Math.max(MIN_DIAGRAM_SLICE_HEIGHT_UNITS, vbWidth / LANDSCAPE_PAGE_CONTENT_RATIO);

    if (vbHeight > maxSliceHeight) {
      const slices: HTMLElement[] = [];
      const step = Math.max(220, maxSliceHeight - DIAGRAM_SLICE_OVERLAP_UNITS);
      const totalSlices = Math.min(MAX_DIAGRAM_SLICES, Math.ceil((vbHeight - DIAGRAM_SLICE_OVERLAP_UNITS) / step));

      for (let i = 0; i < totalSlices; i += 1) {
        const yOffset = vbY + (i * step);
        const sliceStart = i === 0 ? vbY : Math.max(vbY, yOffset - DIAGRAM_SLICE_OVERLAP_UNITS);
        const remaining = (vbY + vbHeight) - sliceStart;
        if (remaining <= 0) break;

        const thisSliceHeight = Math.min(maxSliceHeight + DIAGRAM_SLICE_OVERLAP_UNITS, remaining);
        const sliceFigure = figureClone.cloneNode(true) as HTMLElement;
        const sliceSvg = sliceFigure.querySelector('svg');
        if (!sliceSvg) continue;

        enforceMinDiagramFont(sliceSvg as SVGElement);
        sliceSvg.setAttribute('viewBox', `${vbX} ${sliceStart} ${vbWidth} ${thisSliceHeight}`);
        sliceSvg.setAttribute('preserveAspectRatio', 'xMidYMin meet');
        sliceSvg.removeAttribute('width');
        sliceSvg.removeAttribute('height');

        const printImage = convertSvgToPrintImage(sliceSvg as SVGElement, doc);
        if (printImage) {
          sliceSvg.replaceWith(printImage);
        }

        const wrappedSlice = doc.createElement('section');
        wrappedSlice.className = 'pdf-diagram-page pdf-diagram-slice';
        if (i === 0) {
          wrappedSlice.classList.add('pdf-diagram-slice-first');
        }

        if (totalSlices > 1) {
          const label = doc.createElement('p');
          label.className = 'pdf-split-note';
          label.textContent = `Diagram part ${i + 1} of ${totalSlices}`;
          wrappedSlice.appendChild(label);
        }

        wrappedSlice.appendChild(sliceFigure);
        slices.push(wrappedSlice);
      }

      if (slices.length > 1) {
        return slices;
      }
    }
  }

  const wrapped = doc.createElement('section');
  wrapped.className = shouldLandscape ? 'pdf-landscape-page pdf-diagram-page' : 'pdf-diagram-page';
  wrapped.appendChild(printSafeFigure);
  return [wrapped];
}

export function optimizeRenderedHtmlForPdf(bodyHtml: string): string {
  if (typeof DOMParser === 'undefined' || typeof document === 'undefined') {
    return bodyHtml;
  }

  const parser = new DOMParser();
  const parsed = parser.parseFromString(`<div id="pdf-export-root">${bodyHtml}</div>`, 'text/html');
  const root = parsed.querySelector('#pdf-export-root') as HTMLElement | null;
  if (!root) return bodyHtml;

  const tables = root.querySelectorAll<HTMLTableElement>('table');
  const figures = root.querySelectorAll<HTMLElement>('figure');

  // Remove controls that should never appear in print output.
  root.querySelectorAll('button, details, summary, [role="toolbar"]').forEach((node) => node.remove());

  // Unwrap table scroll containers and clear overflow styles to avoid printed scrollbars.
  root.querySelectorAll<HTMLElement>('div[role="region"][aria-label="Table"]').forEach((container) => {
    const nestedTable = container.querySelector('table');
    if (nestedTable) {
      nestedTable.setAttribute('data-pdf-overflow', '1');
    }

    const parent = container.parentElement;
    if (!parent) return;
    while (container.firstChild) {
      parent.insertBefore(container.firstChild, container);
    }
    parent.removeChild(container);
  });

  root.querySelectorAll<HTMLElement>('[style*="overflow"]').forEach((node) => {
    node.style.removeProperty('overflow');
    node.style.removeProperty('overflow-x');
    node.style.removeProperty('overflow-y');
  });

  // Split oversized diagrams into landscape slices.
  figures.forEach((figure) => {
    if (!figure.querySelector('svg')) return;
    const replacement = splitDiagramFigure(figure, parsed);
    const parent = figure.parentElement;
    if (!parent || replacement.length === 0) return;
    replacement.forEach((node) => parent.insertBefore(node, figure));
    parent.removeChild(figure);
  });

  // Split extra-wide tables and retain key columns in every chunk.
  tables.forEach((table) => {
    const replacement = splitWideTable(table, parsed);
    if (replacement.length === 1 && replacement[0] === table) return;
    const parent = table.parentElement;
    if (!parent) return;
    replacement.forEach((node) => parent.insertBefore(node, table));
    parent.removeChild(table);
  });

  return root.innerHTML;
}

export function sanitizeHtmlForPdfExport(html: string): string {
  // Uses DOMPurify (a vetted DOM-based sanitizer) instead of hand-rolled
  // regular expressions, which cannot reliably parse arbitrary HTML.
  return DOMPurify.sanitize(html, {
    FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'style'],
    FORBID_ATTR: ['srcdoc'],
  });
}

export function toPdfFileName(fileName: string): string {
  const trimmed = fileName.trim();
  if (!trimmed) return 'PP-MD Document.pdf';
  if (/\.pdf$/i.test(trimmed)) return trimmed;
  if (/\.[a-z0-9]{1,8}$/i.test(trimmed)) {
    return trimmed.replace(/\.[a-z0-9]{1,8}$/i, '.pdf');
  }
  return `${trimmed}.pdf`;
}

export function buildPdfDocumentHtml(options: {
  title: string;
  bodyHtml: string;
  language: string;
}): string {
  const safeTitle = escapeHtml(options.title.trim() || 'PP-MD Documentation');
  return `<!doctype html>
<html lang="${escapeHtml(options.language || 'en')}">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${safeTitle}</title>
    <style>
      @page {
        size: A4 landscape;
        margin: 8mm;
      }

      :root {
        color-scheme: light;
      }

      html,
      body {
        margin: 0;
        padding: 0;
        background: #ffffff;
        color: #111827;
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 11pt;
        line-height: 1.45;
      }

      main {
        display: block;
      }

      h1, h2, h3, h4, h5, h6 {
        color: #0f172a;
        break-after: avoid-page;
      }

      h1 {
        margin: 0 0 8mm;
        font-size: 20pt;
      }

      h2 {
        margin-top: 7mm;
        font-size: 15pt;
      }

      h3 {
        margin-top: 5mm;
        font-size: 12.5pt;
      }

      p,
      li,
      td,
      th {
        font-size: 10.5pt;
      }

      a {
        color: #0f4fb8;
      }

      table {
        width: 100%;
        border-collapse: collapse;
        margin: 4mm 0;
      }

      th,
      td {
        border: 1px solid #475569;
        padding: 2.5mm;
        text-align: left;
        vertical-align: top;
      }

      th {
        background: #f1f5f9;
      }

      code {
        font-family: "Cascadia Code", Consolas, monospace;
        font-size: 9.5pt;
      }

      pre {
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        padding: 3mm;
        overflow: visible;
        white-space: pre-wrap;
      }

      figure {
        margin: 4mm 0;
        break-inside: avoid;
      }

      figure svg {
        width: 100%;
        height: auto;
        max-width: 100%;
        max-height: 192mm;
        display: block;
        margin: 0 auto;
        overflow: visible;
      }

      figcaption {
        font-size: 9.5pt;
        color: #334155;
      }

      .pdf-landscape-page {
        break-before: auto;
        break-inside: auto;
      }

      .pdf-wide-table table {
        width: 100%;
        table-layout: auto;
      }

      .pdf-wide-table th,
      .pdf-wide-table td {
        word-break: break-word;
        overflow-wrap: anywhere;
        white-space: pre-wrap;
      }

      .pdf-diagram-page figure {
        margin: 0;
        break-inside: avoid;
      }

      .pdf-diagram-page {
        display: block;
        break-inside: avoid;
      }

      .pdf-diagram-page figure svg {
        max-height: none;
        height: auto;
        width: 100%;
      }

      .pdf-diagram-slice {
        break-before: page;
        page-break-before: always;
      }

      .pdf-diagram-slice:first-of-type {
        break-before: auto;
        page-break-before: auto;
      }

      .pdf-diagram-slice-first {
        break-before: auto;
        page-break-before: auto;
      }

      .pdf-split-note {
        margin: 0 0 3mm;
        font-size: 9pt;
        color: #334155;
      }

      svg text:not([font-size]) {
        font-size: 8pt;
      }

      ::-webkit-scrollbar {
        width: 0;
        height: 0;
      }

      button,
      details,
      summary,
      [role="toolbar"] {
        display: none !important;
      }
    </style>
  </head>
  <body>
    <main aria-label="${safeTitle}">
      ${options.bodyHtml}
    </main>
  </body>
</html>`;
}

export async function exportRenderedMarkdownAsPdf(request: PdfExportRequest): Promise<PdfExportResult> {
  const runtimeWindow = window as Window & {
    electron?: {
      invoke: (channel: string, ...args: unknown[]) => Promise<PdfExportResult>;
    };
  };
  if (!runtimeWindow.electron?.invoke) {
    return {
      cancelled: true,
      error: 'PDF export is available in the desktop app only.',
    };
  }

  try {
    const sanitizedHtml = sanitizeHtmlForPdfExport(request.renderedHtml);

    let optimizedHtml = sanitizedHtml;
    try {
      optimizedHtml = optimizeRenderedHtmlForPdf(sanitizedHtml);
    } catch {
      // Fall back to sanitized HTML if optimization fails on edge-case DOM content.
      optimizedHtml = sanitizedHtml;
    }

    const documentHtml = buildPdfDocumentHtml({
      title: request.title,
      language: request.language || 'en',
      bodyHtml: optimizedHtml,
    });

    const payload = {
      title: request.title,
      defaultFileName: toPdfFileName(request.fileName),
      html: documentHtml,
    };

    return runtimeWindow.electron.invoke('export-markdown-pdf', payload) as Promise<PdfExportResult>;
  } catch (error) {
    return {
      cancelled: false,
      error: `PDF export failed: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
}
