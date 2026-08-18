/**
 * @file MarkdownViewer.tsx
 * @description Renders Markdown with GFM, Mermaid diagram support, and a
 * raw/rendered toggle. WCAG compliant with accessible diagrams, headings,
 * tables, and copy/export actions.
 */

import { useState, useCallback, useRef, useMemo, type MouseEvent, type ReactNode } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import rehypeRaw from 'rehype-raw';
import { MermaidDiagram } from './MermaidDiagram';
import { exportRenderedMarkdownAsPdf } from '../../utils/pdfExport';
import { exportMarkdownTablesToExcel } from '../../utils/tableExport';
import styles from './MarkdownViewer.module.css';

export interface MarkdownViewerProps {
  markdown: string;
  title?: string;
  fileName?: string;
  onExport?: () => void;
  onStatusMessage?: (message: string) => void;
}

/**
 * Extract text from nested React nodes.
 */
function extractText(node: ReactNode): string {
  if (typeof node === 'string' || typeof node === 'number') return String(node);
  if (Array.isArray(node)) return node.map(extractText).join(' ');
  if (node && typeof node === 'object' && 'props' in node) {
    return extractText((node as { props?: { children?: ReactNode } }).props?.children);
  }
  return '';
}

/**
 * Convert node text to URL-safe heading ID.
 */
function slugifyHeading(children: ReactNode): string {
  return extractText(children)
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-') || 'section';
}

/**
 * Renders Markdown with diagram support, raw/rendered toggle, and export.
 */
export function MarkdownViewer({ markdown, title, fileName, onExport, onStatusMessage }: MarkdownViewerProps) {
  const [showRaw, setShowRaw] = useState(false);
  const [copyMsg, setCopyMsg] = useState('');
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const [mermaidProgressState, setMermaidProgressState] = useState<{
    docKey: string;
    renderedIds: Set<string>;
  }>({ docKey: '', renderedIds: new Set() });
  const contentRef = useRef<HTMLDivElement>(null);

  const totalMermaidDiagrams = useMemo(() => (markdown.match(/```mermaid/g) ?? []).length, [markdown]);
  const mermaidDocumentKey = useMemo(
    () => `${fileName || title || 'document'}::${markdown.length}::${totalMermaidDiagrams}`,
    [fileName, markdown.length, title, totalMermaidDiagrams],
  );

  const handleMermaidRenderState = useCallback((groupKey: string, diagramId: string, state: 'rendering' | 'done') => {
    setMermaidProgressState((prev) => {
      if (groupKey !== mermaidDocumentKey) {
        return prev;
      }

      const base = prev.docKey === groupKey
        ? prev
        : { docKey: groupKey, renderedIds: new Set<string>() };
      const next = new Set(base.renderedIds);
      if (state === 'done') {
        next.add(diagramId);
      } else {
        next.delete(diagramId);
      }
      return { docKey: groupKey, renderedIds: next };
    });
  }, [mermaidDocumentKey]);

  const scrollToHeading = useCallback((href: string) => {
    if (!contentRef.current || !href.startsWith('#')) return false;

    const rawTarget = href.slice(1).trim();
    if (!rawTarget) return false;

    const decodedTarget = (() => {
      try {
        return decodeURIComponent(rawTarget);
      } catch {
        return rawTarget;
      }
    })();

    const escapedId = typeof CSS !== 'undefined' && typeof CSS.escape === 'function'
      ? CSS.escape(decodedTarget)
      : decodedTarget.replace(/([ !"#$%&'()*+,./:;<=>?@[\\\]^`{|}~])/g, '\\$1');

    let target = contentRef.current.querySelector<HTMLElement>(`#${escapedId}`);

    if (!target) {
      const lower = decodedTarget.toLowerCase();
      target = Array
        .from(contentRef.current.querySelectorAll<HTMLElement>('[id]'))
        .find((el) => el.id.toLowerCase() === lower) ?? null;
    }

    if (!target) return false;

    target.scrollIntoView({ behavior: 'smooth', block: 'start', inline: 'nearest' });
    window.history.replaceState(null, '', `#${rawTarget}`);
    return true;
  }, []);

  const makeHeading = (Tag: 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6') => {
    return ({ children, ...props }: any) => {
      const id = slugifyHeading(children);
      return <Tag id={id} {...props}>{children}</Tag>;
    };
  };

  const components = useMemo(() => ({
    code: ({ className, children, ...props }: any) => {
      const lang = /language-(\w+)/.exec(className)?.[1] ?? '';
      if (lang !== 'mermaid') {
        return <code className={className} {...props}>{children}</code>;
      }
      const rawSrc = String(children).replace(/\n$/, ""); const src = rawSrc.replace(/([^\n])(linkStyle \d+)/g, "$1\n$2").replace(/(linkStyle \d+ stroke:[^\n]+)(linkStyle)/g, "$1\n$2");
      const cap = src.match(/%%\s*(.+?)\s*%%/)?.[1] ?? 'Diagram';
      return (
        <MermaidDiagram
          chart={src}
          caption={cap}
          mermaidGroupKey={mermaidDocumentKey}
          onRenderStateChange={handleMermaidRenderState}
        />
      );
    },
    h1: makeHeading('h1'),
    h2: makeHeading('h2'),
    h3: makeHeading('h3'),
    h4: makeHeading('h4'),
    h5: makeHeading('h5'),
    h6: makeHeading('h6'),
    a: ({ href, children, ...props }: any) => {
      const handleClick = (event: MouseEvent<HTMLAnchorElement>) => {
        if (!href?.startsWith('#')) return;
        if (!scrollToHeading(href)) return;
        event.preventDefault();
      };

      return <a href={href} {...props} onClick={handleClick}>{children}</a>;
    },
    table: ({ children, ...props }: any) => (
      <div style={{ overflowX: 'auto', margin: '1rem 0' }} role="region" aria-label="Table">
        <table {...props}>{children}</table>
      </div>
    ),
  }), [handleMermaidRenderState, mermaidDocumentKey, scrollToHeading]);

  const renderedMermaidCount = Math.min(
    mermaidProgressState.docKey === mermaidDocumentKey ? mermaidProgressState.renderedIds.size : 0,
    totalMermaidDiagrams,
  );
  const mermaidProgressLabel = totalMermaidDiagrams > 0
    ? `Diagrams ${renderedMermaidCount}/${totalMermaidDiagrams} rendered`
    : 'No diagrams in document';

  const markdownHasRawHtml = useMemo(() => /<\/?[a-z][\s\S]*>/i.test(markdown), [markdown]);
  const rehypePlugins = useMemo(() => (markdownHasRawHtml ? [rehypeRaw] : []), [markdownHasRawHtml]);

  const renderedMarkdown = useMemo(() => (
    <ReactMarkdown
      remarkPlugins={[remarkGfm]}
      rehypePlugins={rehypePlugins}
      components={components}
    >
      {markdown}
    </ReactMarkdown>
  ), [components, markdown, rehypePlugins]);

  const handleCopy = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(markdown);
      setCopyMsg('Copied!');
      setTimeout(() => setCopyMsg(''), 2000);
    } catch {
      setCopyMsg('Copy failed — select and copy manually.');
      setTimeout(() => setCopyMsg(''), 3000);
    }
  }, [markdown]);

  const handleContentClickCapture = useCallback((event: MouseEvent<HTMLDivElement>) => {
    const target = event.target as HTMLElement | null;
    if (!target) return;

    const anchor = target.closest('a[href]') as HTMLAnchorElement | null;
    if (!anchor) return;

    const href = anchor.getAttribute('href') || '';
    if (!href.startsWith('#')) return;

    if (scrollToHeading(href)) {
      event.preventDefault();
      event.stopPropagation();
    }
  }, [scrollToHeading]);

  const handleExportPdf = useCallback(async () => {
    if (isExportingPdf) {
      return;
    }

    if (showRaw) {
      const msg = 'Switch to Rendered view to export diagrams in PDF.';
      setCopyMsg(msg);
      onStatusMessage?.(msg);
      return;
    }

    if (!contentRef.current) {
      const msg = 'PDF export is unavailable until the document finishes rendering.';
      setCopyMsg(msg);
      onStatusMessage?.(msg);
      return;
    }

    const clone = contentRef.current.cloneNode(true) as HTMLElement;
    clone.querySelectorAll('button, details, summary, [role="toolbar"]').forEach((node) => node.remove());

    const exportedTitle = title?.trim() || fileName?.replace(/\.[^.]+$/g, '') || 'PP-MD Documentation';
    const headingNode = document.createElement('h1');
    headingNode.textContent = exportedTitle;
    const bodyHtml = `${headingNode.outerHTML}${clone.innerHTML}`;

    const progressMsg = 'Preparing PDF export. Save dialog should appear shortly.';
    setIsExportingPdf(true);
    setCopyMsg(progressMsg);
    onStatusMessage?.(progressMsg);

    // Yield so the status toast paints before heavy export preparation begins.
    await new Promise<void>((resolve) => {
      requestAnimationFrame(() => resolve());
    });

    let result;
    try {
      result = await exportRenderedMarkdownAsPdf({
        title: exportedTitle,
        fileName: fileName || `${exportedTitle}.pdf`,
        renderedHtml: bodyHtml,
        language: document.documentElement.lang || 'en',
      });
    } catch (error) {
      result = {
        cancelled: false,
        error: `PDF export failed: ${error instanceof Error ? error.message : 'Unexpected renderer exception.'}`,
      };
    } finally {
      setIsExportingPdf(false);
    }

    if (result.cancelled) {
      if (result.error) {
        setCopyMsg(result.error);
        onStatusMessage?.(result.error);
      }
      return;
    }

    if (result.error) {
      setCopyMsg(result.error);
      onStatusMessage?.(result.error);
      return;
    }

    const msg = 'PDF exported successfully.';
    setCopyMsg(msg);
    onStatusMessage?.(msg);
  }, [fileName, isExportingPdf, onStatusMessage, showRaw, title]);

  const handleExportExcel = useCallback(() => {
    const workbookFileName = fileName || title || 'PP-MD Tables.xlsx';
    const result = exportMarkdownTablesToExcel(markdown, workbookFileName);
    setCopyMsg(result.message);
    onStatusMessage?.(result.message);
  }, [fileName, markdown, onStatusMessage, title]);

  return (
    <section
      className={styles.viewer}
      aria-label={title ? `Documentation for ${title}` : 'Generated documentation'}
    >
      <div className={styles.toolbar} role="toolbar" aria-label="Documentation actions">
        {title && (
          <h2
            className={styles.viewerTitle}
            title={fileName || title}
          >
            {title}
          </h2>
        )}
        <div className={styles.actions}>
          <span className={styles.diagramProgress} aria-live="polite">{mermaidProgressLabel}</span>
          <button
            type="button"
            className={styles.toolbarBtn}
            onClick={() => setShowRaw((v) => !v)}
            aria-pressed={showRaw}
            aria-label={showRaw ? 'Switch to rendered view' : 'Switch to raw Markdown view'}
          >
            {showRaw ? '🖼️ Rendered' : '📝 Raw MD'}
          </button>
          <button
            type="button"
            className={styles.toolbarBtn}
            onClick={handleCopy}
            aria-label="Copy Markdown to clipboard"
          >
            📋 Copy
          </button>
          {onExport && (
            <button
              type="button"
              className={`${styles.toolbarBtn} ${styles.primary}`}
              onClick={onExport}
              aria-label="Download Markdown documentation as a .md file"
            >
              ⬇️ Export .md
            </button>
          )}
          <button
            type="button"
            className={`${styles.toolbarBtn} ${styles.primary}`}
            onClick={handleExportExcel}
            aria-label="Download Markdown tables as an Excel workbook with one tab per table"
            title="Export all Markdown tables to Excel"
          >
            ⬇️ Export .xlsx
          </button>
          <button
            type="button"
            className={`${styles.toolbarBtn} ${styles.primary}`}
            onClick={handleExportPdf}
            disabled={isExportingPdf}
            aria-label="Download rendered documentation as an accessible PDF file"
            title={showRaw
              ? 'Switch to Rendered view for diagram-inclusive PDF export'
              : isExportingPdf
                ? 'Preparing PDF export'
                : 'Export rendered documentation to PDF'}
          >
            {isExportingPdf ? '⏳ Preparing PDF...' : '⬇️ Export .pdf'}
          </button>
        </div>
      </div>
      {copyMsg && (
        <div aria-live="polite" className={styles.copyToast} role="status">
          {copyMsg}
        </div>
      )}
      {showRaw ? (
        <pre className={styles.rawSource} aria-label="Raw Markdown source">
          <code>{markdown}</code>
        </pre>
      ) : (
        <div
          ref={contentRef}
          className={`${styles.content} markdown-body`}
          onClickCapture={handleContentClickCapture}
        >
          {renderedMarkdown}
        </div>
      )}
    </section>
  );
}
