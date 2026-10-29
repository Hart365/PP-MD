import { jsPDF } from 'jspdf';
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkGfm from 'remark-gfm';
import type {
  Blockquote,
  Code,
  Heading,
  Html,
  List,
  ListItem,
  Paragraph,
  PhrasingContent,
  Root,
  RootContent,
  Table,
  TableCell,
  TableRow,
} from 'mdast';

// jsPDF's built-in ("standard 14") fonts only support the WinAnsi/Latin-1
// character range. Anything outside it (emoji, arrows, checkmarks used
// throughout the generated Markdown) silently fails to render as glyphs, so
// meaningful symbols are swapped for readable ASCII text first and anything
// left outside the supported range is stripped rather than shown as tofu.
const SYMBOL_REPLACEMENTS: Array<[string, string]> = [
  ['✅', '[Yes]'],
  ['⛔', '[No]'],
  ['⚠️', '[Warning]'],
  ['⚠', '[Warning]'],
  ['🔑', '[Key]'],
  ['🔒', '[Locked]'],
  ['🔍', '[Audited]'],
  ['🔎', '[Advanced Find]'],
  ['✳️', '[Custom]'],
  ['🔀', '[Sync]'],
  ['☁️', '[Cloud]'],
  ['→', '->'],
  ['–', '-'],
  ['—', '-'],
  ['•', '-'],
];

function sanitizePdfText(text: string): string {
  let result = text;
  for (const [symbol, replacement] of SYMBOL_REPLACEMENTS) {
    result = result.split(symbol).join(replacement);
  }
  // Keep printable ASCII plus the Latin-1 supplement (accented characters);
  // drop anything else (remaining emoji/symbols) that the base14 fonts can't render.
  return result.replace(/[^\t\n\r\x20-\x7E\u00A0-\u00FF]/gu, '');
}

function stripTags(html: string): string {
  let text = '';
  let tagStart = -1;
  let quote = '';

  for (let index = 0; index < html.length; index += 1) {
    const character = html[index];
    if (tagStart < 0) {
      if (character === '<') {
        tagStart = index;
      } else {
        text += character;
      }
      continue;
    }

    if (quote) {
      if (character === quote) quote = '';
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === '>') {
      tagStart = -1;
    }
  }

  return tagStart < 0 ? text : text + html.slice(tagStart);
}

interface Style {
  bold?: boolean;
  italic?: boolean;
  code?: boolean;
  link?: string;
}

interface Token {
  text: string;
  style: Style;
  break?: boolean;
}

interface Word {
  text: string;
  style: Style;
  isSpace: boolean;
  hardBreak?: boolean;
}

interface Ctx {
  pdf: jsPDF;
  x0: number;
  xw: number;
  y: number;
  pageHeight: number;
  marginTop: number;
  marginBottom: number;
  diagramImages: ReadonlyArray<string | null>;
  diagramIndex: number;
}

function isExternalLink(url?: string): boolean {
  return !!url && /^(https?:|mailto:)/i.test(url);
}

function flattenInline(nodes: PhrasingContent[] | undefined, base: Style = {}): Token[] {
  if (!nodes) return [];
  const tokens: Token[] = [];
  for (const node of nodes) {
    switch (node.type) {
      case 'text':
        tokens.push({ text: sanitizePdfText(node.value), style: base });
        break;
      case 'strong':
        tokens.push(...flattenInline(node.children, { ...base, bold: true }));
        break;
      case 'emphasis':
        tokens.push(...flattenInline(node.children, { ...base, italic: true }));
        break;
      case 'delete':
        tokens.push(...flattenInline(node.children, base));
        break;
      case 'inlineCode':
        tokens.push({ text: sanitizePdfText(node.value), style: { ...base, code: true } });
        break;
      case 'link':
        tokens.push(...flattenInline(node.children, { ...base, link: node.url }));
        break;
      case 'image':
        tokens.push({ text: sanitizePdfText(`[Image: ${node.alt || 'image'}]`), style: { ...base, italic: true } });
        break;
      case 'break':
        tokens.push({ text: '', style: base, break: true });
        break;
      case 'html':
        if (/^<br\s*\/?>$/i.test(((node as Html).value ?? '').trim())) {
          tokens.push({ text: '', style: base, break: true });
          break;
        }
        tokens.push({ text: sanitizePdfText(stripTags((node as Html).value ?? '')), style: base });
        break;
      default: {
        const withChildren = node as unknown as { children?: PhrasingContent[]; value?: string };
        if (Array.isArray(withChildren.children)) {
          tokens.push(...flattenInline(withChildren.children, base));
        } else if (typeof withChildren.value === 'string') {
          tokens.push({ text: sanitizePdfText(withChildren.value), style: base });
        }
        break;
      }
    }
  }
  return tokens;
}

function toWords(tokens: Token[]): Word[] {
  const words: Word[] = [];
  for (const token of tokens) {
    if (token.break) {
      words.push({ text: '', style: token.style, isSpace: false, hardBreak: true });
      continue;
    }
    if (!token.text) continue;
    const parts = token.text.split(/(\s+)/).filter((part) => part.length > 0);
    for (const part of parts) {
      words.push({ text: part, style: token.style, isSpace: /^\s+$/.test(part) });
    }
  }
  return words;
}

function applyFont(pdf: jsPDF, style: Style) {
  if (style.code) {
    pdf.setFont('courier', style.bold ? 'bold' : 'normal');
    return;
  }
  const variant = style.bold && style.italic ? 'bolditalic' : style.bold ? 'bold' : style.italic ? 'italic' : 'normal';
  pdf.setFont('helvetica', variant);
}

/** Wraps a flat token stream into printable lines without drawing anything (used for measuring table/paragraph height up front). */
function buildLines(pdf: jsPDF, words: Word[], maxWidth: number, fontSize: number): Word[][] {
  const lines: Word[][] = [];
  let line: Word[] = [];
  let lineWidth = 0;

  const flush = () => {
    while (line.length && line[0].isSpace) line.shift();
    while (line.length && line[line.length - 1].isSpace) line.pop();
    lines.push(line);
    line = [];
    lineWidth = 0;
  };

  for (const word of words) {
    if (word.hardBreak) {
      flush();
      continue;
    }
    applyFont(pdf, word.style);
    pdf.setFontSize(fontSize);
    const width = pdf.getTextWidth(word.text);
    if (word.isSpace && line.length === 0) continue;
    if (lineWidth + width > maxWidth && line.length > 0) {
      flush();
      if (word.isSpace) continue;
    }
    line.push(word);
    lineWidth += width;
  }
  if (line.length) flush();
  if (lines.length === 0) lines.push([]);
  return lines;
}

function ensureSpace(ctx: Ctx, neededHeight: number) {
  if (ctx.y + neededHeight > ctx.pageHeight - ctx.marginBottom) {
    ctx.pdf.addPage();
    ctx.y = ctx.marginTop;
  }
}

/**
 * Hands control back to the browser for a frame. jsPDF's drawing calls are
 * all synchronous, so without periodic yields a large document blocks the
 * main thread for the entire export — the "Generating PDF..." notice never
 * gets a chance to paint (or update) and the tab appears to freeze until the
 * whole thing suddenly finishes.
 */
function yieldToBrowser(): Promise<void> {
  return new Promise((resolve) => {
    const timer = setTimeout(resolve, 16);
    if (typeof requestAnimationFrame === 'function') {
      requestAnimationFrame(() => {
        clearTimeout(timer);
        resolve();
      });
    }
  });
}

function drawWrappedWords(ctx: Ctx, words: Word[], x: number, maxWidth: number, fontSize: number, lineHeight: number) {
  const lines = buildLines(ctx.pdf, words, maxWidth, fontSize);
  for (const line of lines) {
    ensureSpace(ctx, lineHeight);
    let cx = x;
    for (const word of line) {
      applyFont(ctx.pdf, word.style);
      ctx.pdf.setFontSize(fontSize);
      const width = ctx.pdf.getTextWidth(word.text);
      if (word.style.code) {
        ctx.pdf.setFillColor(244, 244, 248);
        ctx.pdf.rect(cx - 1, ctx.y - fontSize * 0.78, width + 2, fontSize * 1.05, 'F');
        applyFont(ctx.pdf, word.style);
        ctx.pdf.setFontSize(fontSize);
      }
      if (word.style.link) {
        ctx.pdf.setTextColor(37, 99, 235);
        if (isExternalLink(word.style.link)) {
          ctx.pdf.textWithLink(word.text, cx, ctx.y, { url: word.style.link });
        } else {
          ctx.pdf.text(word.text, cx, ctx.y);
        }
        ctx.pdf.setDrawColor(37, 99, 235);
        ctx.pdf.line(cx, ctx.y + 1.5, cx + width, ctx.y + 1.5);
      } else {
        ctx.pdf.setTextColor(26, 26, 46);
        ctx.pdf.text(word.text, cx, ctx.y);
      }
      cx += width;
    }
    ctx.y += lineHeight;
  }
}

function drawHeading(ctx: Ctx, node: Heading) {
  const HEADING_SIZES = [19, 15.5, 13.5, 12, 11, 10.5];
  const depth = Math.min(Math.max(node.depth, 1), 6);
  const fontSize = HEADING_SIZES[depth - 1];
  ctx.y += depth === 1 ? 16 : 12;
  const words = toWords(flattenInline(node.children, { bold: true }));
  drawWrappedWords(ctx, words, ctx.x0, ctx.xw, fontSize, fontSize * 1.3);
  if (depth <= 2) {
    ctx.pdf.setDrawColor(200, 203, 214);
    ctx.pdf.line(ctx.x0, ctx.y - 2, ctx.x0 + ctx.xw, ctx.y - 2);
    ctx.y += 4;
  }
  ctx.y += 6;
}

function drawParagraph(ctx: Ctx, node: Paragraph, opts?: { x?: number; width?: number; fontSize?: number; italic?: boolean }) {
  const x = opts?.x ?? ctx.x0;
  const width = opts?.width ?? ctx.xw;
  const fontSize = opts?.fontSize ?? 10.25;
  const words = toWords(flattenInline(node.children, opts?.italic ? { italic: true } : {}));
  drawWrappedWords(ctx, words, x, width, fontSize, fontSize * 1.42);
  ctx.y += fontSize * 0.6;
}

function drawThematicBreak(ctx: Ctx) {
  ensureSpace(ctx, 20);
  ctx.y += 8;
  ctx.pdf.setDrawColor(200, 203, 214);
  ctx.pdf.line(ctx.x0, ctx.y, ctx.x0 + ctx.xw, ctx.y);
  ctx.y += 14;
}

async function drawList(ctx: Ctx, node: List, depth = 0) {
  let counter = node.start ?? 1;
  let itemCount = 0;
  for (const item of node.children as ListItem[]) {
    const indent = ctx.x0 + 16 + depth * 16;
    let marker = node.ordered ? `${counter}.` : '-';
    counter += 1;
    if (item.checked === true) marker = '[x]';
    else if (item.checked === false) marker = '[ ]';

    let firstBlock = true;
    for (const child of item.children) {
      if (child.type === 'list') {
        await drawList(ctx, child as List, depth + 1);
        continue;
      }
      const width = ctx.xw - (indent - ctx.x0);
      if (child.type === 'paragraph') {
        const words = toWords(flattenInline((child as Paragraph).children));
        if (firstBlock) {
          const lines = buildLines(ctx.pdf, words, width, 10.25);
          for (const line of lines) {
            ensureSpace(ctx, 10.25 * 1.42);
            ctx.pdf.setFont('helvetica', 'normal');
            ctx.pdf.setFontSize(10.25);
            ctx.pdf.setTextColor(26, 26, 46);
            if (line === lines[0]) ctx.pdf.text(marker, indent - 14, ctx.y);
            let cx = indent;
            for (const word of line) {
              applyFont(ctx.pdf, word.style);
              ctx.pdf.setFontSize(10.25);
              ctx.pdf.setTextColor(26, 26, 46);
              ctx.pdf.text(word.text, cx, ctx.y);
              cx += ctx.pdf.getTextWidth(word.text);
            }
            ctx.y += 10.25 * 1.42;
          }
        } else {
          drawWrappedWords(ctx, words, indent, width, 10.25, 10.25 * 1.42);
        }
        firstBlock = false;
      } else {
        await renderBlock(ctx, child, { x: indent, width });
      }
    }
    ctx.y += 3;
    itemCount += 1;
    if (itemCount % 25 === 0) await yieldToBrowser();
  }
}

async function drawBlockquote(ctx: Ctx, node: Blockquote) {
  const startY = ctx.y;
  const x = ctx.x0 + 14;
  const width = ctx.xw - 14;
  for (const child of node.children) {
    await renderBlock(ctx, child, { x, width, italic: true });
  }
  ctx.pdf.setDrawColor(180, 184, 196);
  ctx.pdf.setLineWidth(2);
  ctx.pdf.line(ctx.x0 + 4, startY - 2, ctx.x0 + 4, Math.max(ctx.y - 4, startY - 2));
  ctx.pdf.setLineWidth(0.2);
}

function drawCodeBlock(ctx: Ctx, node: Pick<Code, 'lang' | 'value'>) {
  if (node.lang === 'mermaid') {
    drawMermaidBlock(ctx, node.value ?? '');
    return;
  }
  const fontSize = 8.5;
  const lineHeight = fontSize * 1.35;
  const pad = 6;
  const innerWidth = ctx.xw - pad * 2;
  const text = sanitizePdfText(node.value ?? '');

  ctx.pdf.setFont('courier', 'normal');
  ctx.pdf.setFontSize(fontSize);
  const wrapped: string[] = [];
  for (const raw of text.split('\n')) {
    if (raw === '') {
      wrapped.push('');
      continue;
    }
    let remaining = raw;
    while (ctx.pdf.getTextWidth(remaining) > innerWidth) {
      let cut = remaining.length;
      while (cut > 1 && ctx.pdf.getTextWidth(remaining.slice(0, cut)) > innerWidth) cut -= 1;
      const spaceIdx = remaining.lastIndexOf(' ', cut);
      const breakAt = spaceIdx > 10 ? spaceIdx : cut;
      wrapped.push(remaining.slice(0, breakAt));
      remaining = remaining.slice(breakAt).replace(/^ /, '');
    }
    wrapped.push(remaining);
  }

  let idx = 0;
  while (idx < wrapped.length) {
    ensureSpace(ctx, lineHeight + pad * 2);
    const maxRows = Math.max(1, Math.floor((ctx.pageHeight - ctx.marginBottom - ctx.y - pad * 2) / lineHeight));
    const rows = wrapped.slice(idx, idx + maxRows);
    const boxHeight = rows.length * lineHeight + pad * 2;
    ctx.pdf.setFillColor(244, 244, 248);
    ctx.pdf.setDrawColor(208, 211, 219);
    ctx.pdf.rect(ctx.x0, ctx.y, ctx.xw, boxHeight, 'FD');
    ctx.pdf.setFont('courier', 'normal');
    ctx.pdf.setFontSize(fontSize);
    ctx.pdf.setTextColor(26, 26, 46);
    let ty = ctx.y + pad + fontSize * 0.8;
    for (const row of rows) {
      ctx.pdf.text(row || ' ', ctx.x0 + pad, ty);
      ty += lineHeight;
    }
    ctx.y += boxHeight + 2;
    idx += rows.length;
  }
  ctx.y += 6;
}

function drawMermaidBlock(ctx: Ctx, source: string) {
  const caption = source.match(/%%\s*(.+?)\s*%%/)?.[1] ?? 'Diagram';
  const image = ctx.diagramImages[ctx.diagramIndex] ?? null;
  ctx.diagramIndex += 1;

  ensureSpace(ctx, 20);
  ctx.pdf.setFont('helvetica', 'italic');
  ctx.pdf.setFontSize(9.5);
  ctx.pdf.setTextColor(90, 94, 108);
  ctx.pdf.text(sanitizePdfText(`Diagram: ${caption}`), ctx.x0, ctx.y);
  ctx.y += 14;

  if (!image) {
    drawCodeBlock(ctx, { lang: 'text', value: `[Diagram could not be rendered for export]\n\n${source}` });
    return;
  }

  try {
    const props = ctx.pdf.getImageProperties(image);
    let width = ctx.xw;
    let height = (props.height / props.width) * width;
    const maxHeight = ctx.pageHeight - ctx.marginTop - ctx.marginBottom;
    if (height > maxHeight) {
      height = maxHeight;
      width = (props.width / props.height) * height;
    }
    ensureSpace(ctx, height);
    const offsetX = ctx.x0 + (ctx.xw - width) / 2;
    ctx.pdf.addImage(image, 'PNG', offsetX, ctx.y, width, height, undefined, 'FAST');
    ctx.y += height + 10;
  } catch {
    drawCodeBlock(ctx, { lang: 'text', value: source });
  }
}

function drawGridRow(ctx: Ctx, cellLines: Word[][][], numCols: number, colWidth: number, height: number, fontSize: number, lineHeight: number, pad: number, header: boolean) {
  if (header) {
    ctx.pdf.setFillColor(240, 242, 245);
    ctx.pdf.rect(ctx.x0, ctx.y, ctx.xw, height, 'F');
  }
  let cx = ctx.x0;
  for (let c = 0; c < numCols; c += 1) {
    ctx.pdf.setDrawColor(193, 196, 206);
    ctx.pdf.rect(cx, ctx.y, colWidth, height, 'S');
    let ty = ctx.y + pad + fontSize * 0.72;
    for (const line of cellLines[c] ?? []) {
      let tx = cx + pad;
      for (const word of line) {
        applyFont(ctx.pdf, word.style);
        ctx.pdf.setFontSize(fontSize);
        ctx.pdf.setTextColor(26, 26, 46);
        ctx.pdf.text(word.text, tx, ty);
        tx += ctx.pdf.getTextWidth(word.text);
      }
      ty += lineHeight;
    }
    cx += colWidth;
  }
  ctx.y += height;
}

async function drawTable(ctx: Ctx, node: Table) {
  const rows = node.children as TableRow[];
  if (!rows.length) return;
  const numCols = rows[0].children.length;
  if (numCols === 0) return;
  const pad = 5;
  const fontSize = 9.5;
  const lineHeight = fontSize * 1.3;
  const colWidth = ctx.xw / numCols;

  const cellsToLines = (cells: TableCell[], bold: boolean) => cells.map((cell) => {
    const words = toWords(flattenInline(cell.children, bold ? { bold: true } : {}));
    return buildLines(ctx.pdf, words, colWidth - pad * 2, fontSize);
  });

  const headerLines = cellsToLines(rows[0].children, true);
  const headerHeight = Math.max(...headerLines.map((lines) => lines.length), 1) * lineHeight + pad * 2;

  ensureSpace(ctx, headerHeight + lineHeight + pad * 2);
  drawGridRow(ctx, headerLines, numCols, colWidth, headerHeight, fontSize, lineHeight, pad, true);

  for (let r = 1; r < rows.length; r += 1) {
    const lines = cellsToLines(rows[r].children, false);
    const rowHeight = Math.max(...lines.map((l) => l.length), 1) * lineHeight + pad * 2;
    if (ctx.y + rowHeight > ctx.pageHeight - ctx.marginBottom) {
      ctx.pdf.addPage();
      ctx.y = ctx.marginTop;
      drawGridRow(ctx, headerLines, numCols, colWidth, headerHeight, fontSize, lineHeight, pad, true);
    }
    drawGridRow(ctx, lines, numCols, colWidth, rowHeight, fontSize, lineHeight, pad, false);
    if (r % 40 === 0) await yieldToBrowser();
  }
  ctx.y += 8;
}

/** Renders PP-MD's own hand-written `<table>` blocks (e.g. the document context/metadata table), which remark parses as a raw HTML block rather than a GFM table. */
function drawHtmlTableBlock(ctx: Ctx, html: string) {
  const rowMatches = [...html.matchAll(/<tr>([\s\S]*?)<\/tr>/gi)];
  const rows = rowMatches
    .map((match) => [...match[1].matchAll(/<td>([\s\S]*?)<\/td>/gi)].map((cell) => sanitizePdfText(stripTags(cell[1]).trim())))
    .filter((row) => row.some((cell) => cell.length > 0));
  if (!rows.length) return;

  const numCols = Math.max(...rows.map((row) => row.length));
  const pad = 5;
  const fontSize = 10;
  const lineHeight = fontSize * 1.3;
  const colWidth = ctx.xw / numCols;

  for (const row of rows) {
    const cellLines = Array.from({ length: numCols }, (_, i) => {
      const text = row[i] ?? '';
      const words = toWords([{ text, style: { bold: i === 0 } }]);
      return buildLines(ctx.pdf, words, colWidth - pad * 2, fontSize);
    });
    const rowHeight = Math.max(...cellLines.map((l) => l.length), 1) * lineHeight + pad * 2;
    ensureSpace(ctx, rowHeight);
    drawGridRow(ctx, cellLines, numCols, colWidth, rowHeight, fontSize, lineHeight, pad, false);
  }
  ctx.y += 8;
}

async function renderBlock(ctx: Ctx, node: RootContent, opts?: { x?: number; width?: number; italic?: boolean }) {
  switch (node.type) {
    case 'heading':
      drawHeading(ctx, node);
      break;
    case 'paragraph':
      drawParagraph(ctx, node, opts);
      break;
    case 'list':
      await drawList(ctx, node);
      break;
    case 'table':
      await drawTable(ctx, node);
      break;
    case 'code':
      drawCodeBlock(ctx, node);
      break;
    case 'blockquote':
      await drawBlockquote(ctx, node);
      break;
    case 'thematicBreak':
      drawThematicBreak(ctx);
      break;
    case 'html': {
      const raw = (node as Html).value ?? '';
      if (/<table/i.test(raw)) drawHtmlTableBlock(ctx, raw);
      break;
    }
    default:
      break;
  }
}

function parseMarkdown(markdown: string): Root {
  return unified().use(remarkParse).use(remarkGfm).parse(markdown) as Root;
}

function drawTitlePage(ctx: Ctx, title: string) {
  const centerX = ctx.x0 + ctx.xw / 2;
  ctx.pdf.setFont('helvetica', 'bold');
  ctx.pdf.setFontSize(26);
  ctx.pdf.setTextColor(26, 26, 46);
  const titleLines = ctx.pdf.splitTextToSize(sanitizePdfText(title), ctx.xw);
  let y = ctx.pageHeight / 2 - (titleLines.length * 32) / 2;
  for (const line of titleLines) {
    ctx.pdf.text(line, centerX, y, { align: 'center' });
    y += 32;
  }
  ctx.pdf.setFont('helvetica', 'normal');
  ctx.pdf.setFontSize(11);
  ctx.pdf.setTextColor(90, 94, 108);
  ctx.pdf.text('Power Platform solution documentation', centerX, y + 12, { align: 'center' });
  ctx.pdf.text(`Generated ${new Date().toLocaleString()}`, centerX, y + 30, { align: 'center' });
}

function addPageNumbers(pdf: jsPDF, title: string) {
  const pageCount = pdf.getNumberOfPages();
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  for (let i = 2; i <= pageCount; i += 1) {
    pdf.setPage(i);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(8.5);
    pdf.setTextColor(140, 143, 156);
    pdf.text(sanitizePdfText(title), 46, pageHeight - 24);
    pdf.text(`Page ${i - 1} of ${pageCount - 1}`, pageWidth - 46, pageHeight - 24, { align: 'right' });
  }
}

/**
 * Builds a genuine text-based, searchable PDF directly from the generated
 * Markdown source (not a screenshot of the rendered view). Mermaid diagrams
 * are embedded as raster images using pre-rendered PNGs supplied by the
 * caller (in document order); everything else — headings, paragraphs,
 * tables, lists, code blocks, links — is drawn as real vector text.
 */
export async function createRenderedDocumentationPdf(
  markdown: string,
  title: string,
  diagramImages: ReadonlyArray<string | null> = [],
): Promise<Blob> {
  const tree = parseMarkdown(markdown);
  const pdf = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4', compress: true });
  pdf.setProperties({ title, subject: 'Power Platform solution documentation' });

  const marginTop = 54;
  const marginBottom = 50;
  const marginX = 46;
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  const ctx: Ctx = {
    pdf,
    x0: marginX,
    xw: pageWidth - marginX * 2,
    y: marginTop,
    pageHeight,
    marginTop,
    marginBottom,
    diagramImages,
    diagramIndex: 0,
  };

  drawTitlePage(ctx, title || 'PP-MD Documentation');
  pdf.addPage();
  ctx.y = marginTop;

  let blockCount = 0;
  for (const node of tree.children) {
    await renderBlock(ctx, node);
    blockCount += 1;
    if (blockCount % 15 === 0) await yieldToBrowser();
  }

  addPageNumbers(pdf, title || 'PP-MD Documentation');

  return pdf.output('blob');
}
