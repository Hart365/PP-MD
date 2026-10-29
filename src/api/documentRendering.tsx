/**
 * @file documentRendering.tsx
 * @description Shared, DOM-free(ish) helpers for turning generated Markdown
 * into standalone HTML and rasterised diagram images. Used both by the
 * in-app Markdown viewer and by the batch "export all" actions, which render
 * documents without mounting them on screen.
 */

import type { ComponentPropsWithoutRef, ReactNode } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import ReactMarkdown from 'react-markdown';
import type { Components } from 'react-markdown';
import remarkGfm from 'remark-gfm';
import rehypeRaw from 'rehype-raw';
import rehypeSanitize, { defaultSchema } from 'rehype-sanitize';
import { renderMermaidToSvg, type MermaidDiagramProps } from '../components/ui/MermaidDiagram';

/** Generated Markdown is pure GFM, so the default strict schema is sufficient. */
export const markdownSanitizationSchema = defaultSchema;

const MAX_DIAGRAM_RASTER_DIMENSION = 4000;

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Extract plain text from nested React nodes so markdown headings can be
 * normalized into stable in-document anchors.
 */
export function extractText(node: ReactNode): string {
  if (typeof node === 'string' || typeof node === 'number') return String(node);
  if (Array.isArray(node)) return node.map(extractText).join(' ');
  if (node && typeof node === 'object' && 'props' in node) {
    return extractText((node as { props?: { children?: ReactNode } }).props?.children);
  }
  return '';
}

/**
 * Convert heading node content into a URL-safe slug used as the heading id.
 */
export function slugifyHeading(children: ReactNode): string {
  return extractText(children)
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-') || 'section';
}

/**
 * Wraps already-rendered, sanitized documentation markup in a standalone,
 * self-contained HTML document (embedded styles, no external requests) so it
 * can be opened directly in any browser outside of PP-MD.
 */
export function buildStandaloneHtmlDocument(title: string, bodyHtml: string): string {
  return `<!doctype html>
<html lang="en-GB">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>${escapeHtml(title)}</title>
<style>
  :root { color-scheme: light; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; line-height: 1.6; max-width: 68rem; margin: 0 auto; padding: 2rem; background: #ffffff; color: #1a1a2e; }
  h1, h2, h3, h4 { line-height: 1.25; margin: 1.5em 0 0.5em; }
  table { border-collapse: collapse; width: 100%; margin: 1em 0; overflow-x: auto; display: block; }
  th, td { border: 1px solid #c1c4ce; padding: 0.4em 0.6em; text-align: left; }
  th { background: #f0f2f5; }
  code { font-family: 'Cascadia Code', 'Consolas', monospace; background: #f4f4f8; padding: 0.1em 0.4em; border-radius: 4px; }
  pre { background: #f4f4f8; border: 1px solid #d0d3db; border-radius: 8px; padding: 1rem; overflow-x: auto; }
  figure { margin: 1.5em 0; overflow-x: auto; }
  figcaption { font-size: 0.875rem; color: #4a4a5e; margin-top: 0.5em; }
  svg { max-width: 100%; height: auto; }
  a { color: #2563eb; }
</style>
</head>
<body>
${bodyHtml}
</body>
</html>`;
}

/**
 * Rasterizes a rendered Mermaid `<svg>` into a PNG data URL so it can be
 * embedded in the PDF export (jsPDF cannot draw arbitrary SVG markup
 * directly). Returns null on any failure so the caller can fall back to
 * rendering the diagram's raw source as text instead.
 */
export async function svgToPngDataUrl(svgEl: SVGSVGElement): Promise<string | null> {
  try {
    const clone = svgEl.cloneNode(true) as SVGSVGElement;
    const viewBox = clone.getAttribute('viewBox')?.trim().split(/\s+/).map(Number);
    const width = viewBox && viewBox.length === 4 && viewBox[2] > 0 ? viewBox[2] : (clone.clientWidth || 800);
    const height = viewBox && viewBox.length === 4 && viewBox[3] > 0 ? viewBox[3] : (clone.clientHeight || 600);

    const svgNs = 'http://www.w3.org/2000/svg';
    const background = document.createElementNS(svgNs, 'rect');
    background.setAttribute('x', '0');
    background.setAttribute('y', '0');
    background.setAttribute('width', String(width));
    background.setAttribute('height', String(height));
    background.setAttribute('fill', '#ffffff');
    clone.insertBefore(background, clone.firstChild);
    clone.setAttribute('width', String(width));
    clone.setAttribute('height', String(height));

    const svgString = new XMLSerializer().serializeToString(clone);
    const svgUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgString)}`;

    let scale = 2;
    if (width * scale > MAX_DIAGRAM_RASTER_DIMENSION) scale = MAX_DIAGRAM_RASTER_DIMENSION / width;
    if (height * scale > MAX_DIAGRAM_RASTER_DIMENSION) scale = Math.min(scale, MAX_DIAGRAM_RASTER_DIMENSION / height);
    scale = Math.max(scale, 0.25);

    {
      const image = new Image();
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('Diagram rasterisation timed out')), 5000);
        image.onload = () => {
          try {
            const canvas = document.createElement('canvas');
            canvas.width = Math.max(1, Math.round(width * scale));
            canvas.height = Math.max(1, Math.round(height * scale));
            const context = canvas.getContext('2d');
            if (!context) throw new Error('Canvas 2D context unavailable');
            context.scale(scale, scale);
            context.drawImage(image, 0, 0, width, height);
            resolve(canvas.toDataURL('image/png'));
          } catch (error) {
            reject(error);
          } finally {
            clearTimeout(timer);
          }
        };
        image.onerror = () => {
          clearTimeout(timer);
          reject(new Error('Failed to rasterise diagram SVG'));
        };
        image.src = svgUrl;
      });
      return dataUrl;
    }
  } catch {
    return null;
  }
}

/** Rasterizes SVG markup (e.g. from `mermaid.render`) into a PNG data URL. */
export async function svgMarkupToPngDataUrl(svgMarkup: string): Promise<string | null> {
  try {
    const parsed = new DOMParser().parseFromString(svgMarkup, 'image/svg+xml');
    const svgEl = parsed.documentElement;
    if (!(svgEl instanceof SVGSVGElement)) return null;
    return await svgToPngDataUrl(svgEl);
  } catch {
    return null;
  }
}

export interface RenderedDocument {
  /** Sanitized HTML body markup with Mermaid diagrams inlined as SVG figures. */
  bodyHtml: string;
  /** Rendered SVG markup per `mermaid` fence, in document order (null when rendering failed). */
  diagramSvgs: Array<string | null>;
}

type DiagramTheme = MermaidDiagramProps['theme'];

const MERMAID_PLACEHOLDER = (index: number) => `<span data-pp-md-mermaid="${index}"></span>`;

function makeStaticHeading(Tag: 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6') {
  return ({ children, ...props }: ComponentPropsWithoutRef<'h1'> & { node?: unknown }) => {
    const { node: _node, ...rest } = props;
    void _node;
    return <Tag id={slugifyHeading(children)} {...rest}>{children}</Tag>;
  };
}

/**
 * Renders generated Markdown to standalone-ready HTML without mounting it in
 * the visible viewer. Mermaid fences are rendered (sequentially, via the
 * shared low-priority queue) and inlined as `<figure>` SVGs; failures fall
 * back to the diagram source in a code block.
 */
export async function renderMarkdownDocument(markdown: string, theme: DiagramTheme = 'neutral'): Promise<RenderedDocument> {
  const sources: Array<{ src: string; caption: string }> = [];
  const components: Components = {
    code: ({ className, children, node: _node, ...props }) => {
      void _node;
      const classNameText = typeof className === 'string' ? className : '';
      const lang = /language-(\w+)/.exec(classNameText)?.[1] ?? '';
      if (lang !== 'mermaid') {
        return <code className={classNameText || undefined} {...props}>{children}</code>;
      }
      const src = String(children).replace(/\n$/, '');
      const caption = src.match(/%%\s*(.+?)\s*%%/)?.[1] ?? 'Diagram';
      const index = sources.length;
      sources.push({ src, caption });
      return <span data-pp-md-mermaid={index} />;
    },
    h1: makeStaticHeading('h1'),
    h2: makeStaticHeading('h2'),
    h3: makeStaticHeading('h3'),
    h4: makeStaticHeading('h4'),
    h5: makeStaticHeading('h5'),
    h6: makeStaticHeading('h6'),
  };

  let bodyHtml = renderToStaticMarkup(
    <ReactMarkdown
      remarkPlugins={[remarkGfm]}
      rehypePlugins={[rehypeRaw, [rehypeSanitize, markdownSanitizationSchema]]}
      components={components}
    >
      {markdown}
    </ReactMarkdown>,
  );

  const diagramSvgs: Array<string | null> = [];
  for (let index = 0; index < sources.length; index += 1) {
    const { src, caption } = sources[index];
    let svg: string | null;
    try {
      svg = await renderMermaidToSvg(src, theme);
    } catch {
      svg = null;
    }
    diagramSvgs.push(svg);
    const replacement = svg
      ? `<figure role="group" aria-label="${escapeHtml(caption)}">${svg}<figcaption>${escapeHtml(caption)}</figcaption></figure>`
      : `<pre><code class="language-mermaid">${escapeHtml(src)}</code></pre>`;
    const placeholder = MERMAID_PLACEHOLDER(index);
    const wrapped = `<pre>${placeholder}</pre>`;
    bodyHtml = bodyHtml.includes(wrapped)
      ? bodyHtml.replace(wrapped, replacement)
      : bodyHtml.replace(placeholder, replacement);
  }

  return { bodyHtml, diagramSvgs };
}
