/**
 * @file MarkdownViewer.tsx
 * @description Renders Markdown with GFM, Mermaid diagram support, and a
 * raw/rendered toggle. WCAG compliant with accessible diagrams, headings,
 * tables, and copy/export actions.
 */

import { useState, useCallback, useEffect, useRef, useMemo, type ChangeEvent, type MouseEvent, type ReactNode } from 'react';
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
  const [isExportMenuOpen, setIsExportMenuOpen] = useState(false);
  const [exportMenuPosition, setExportMenuPosition] = useState<{ top: number; right: number }>({ top: 0, right: 0 });
  const [isSearchOpen, setIsSearchOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [searchMatchCount, setSearchMatchCount] = useState(0);
  const [currentSearchIndex, setCurrentSearchIndex] = useState(-1);
  const [mermaidProgressState, setMermaidProgressState] = useState<{
    docKey: string;
    renderedIds: Set<string>;
  }>({ docKey: '', renderedIds: new Set() });
  const contentRef = useRef<HTMLDivElement>(null);
  const exportButtonRef = useRef<HTMLButtonElement>(null);
  const exportMenuRef = useRef<HTMLDivElement>(null);
  const searchInputRef = useRef<HTMLInputElement>(null);
  const searchRangesRef = useRef<Range[]>([]);
  const supportsHighlightApi = typeof CSS !== 'undefined' && 'highlights' in CSS;

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

  const closeExportMenu = useCallback(() => {
    setIsExportMenuOpen(false);
  }, []);

  const toggleExportMenu = useCallback(() => {
    setIsExportMenuOpen((open) => {
      if (open) return false;
      const rect = exportButtonRef.current?.getBoundingClientRect();
      if (rect) {
        setExportMenuPosition({ top: rect.bottom + 4, right: window.innerWidth - rect.right });
      }
      return true;
    });
  }, []);

  // The menu is positioned relative to the viewport (not the toolbar) so it
  // is never clipped by the viewer's `overflow: hidden` container.
  useEffect(() => {
    if (!isExportMenuOpen) return;

    const handlePointerDown = (event: globalThis.MouseEvent) => {
      const target = event.target as Node;
      if (exportMenuRef.current?.contains(target) || exportButtonRef.current?.contains(target)) {
        return;
      }
      closeExportMenu();
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') closeExportMenu();
    };
    const handleReposition = () => {
      const rect = exportButtonRef.current?.getBoundingClientRect();
      if (rect) {
        setExportMenuPosition({ top: rect.bottom + 4, right: window.innerWidth - rect.right });
      }
    };

    document.addEventListener('mousedown', handlePointerDown);
    document.addEventListener('keydown', handleKeyDown);
    window.addEventListener('resize', handleReposition);
    window.addEventListener('scroll', handleReposition, true);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      document.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('resize', handleReposition);
      window.removeEventListener('scroll', handleReposition, true);
    };
  }, [closeExportMenu, isExportMenuOpen]);

  const toggleSearch = useCallback(() => {
    setIsSearchOpen((open) => {
      const next = !open;
      if (!next) {
        setSearchQuery('');
      }
      return next;
    });
  }, []);

  const handleSearchQueryChange = useCallback((event: ChangeEvent<HTMLInputElement>) => {
    setSearchQuery(event.target.value);
  }, []);

  const focusSearchMatch = useCallback((index: number) => {
    const ranges = searchRangesRef.current;
    if (!supportsHighlightApi || ranges.length === 0 || index < 0 || index >= ranges.length) return;

    const range = ranges[index];
    setCurrentSearchIndex(index);
    if (typeof Highlight !== 'undefined') {
      CSS.highlights.set('ppmd-search-active', new Highlight(range));
    }
    const container = range.startContainer instanceof Element
      ? range.startContainer
      : range.startContainer.parentElement;
    container?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }, [supportsHighlightApi]);

  const clearSearchHighlights = useCallback(() => {
    if (supportsHighlightApi) {
      CSS.highlights.delete('ppmd-search');
      CSS.highlights.delete('ppmd-search-active');
    }
    searchRangesRef.current = [];
    setSearchMatchCount(0);
    setCurrentSearchIndex(-1);
  }, [supportsHighlightApi]);

  const resetActiveMatch = useCallback(() => {
    if (supportsHighlightApi) {
      CSS.highlights.delete('ppmd-search-active');
    }
    setCurrentSearchIndex(-1);
  }, [supportsHighlightApi]);

  const handleSearchNext = useCallback(() => {
    const count = searchRangesRef.current.length;
    if (count === 0) return;
    focusSearchMatch((currentSearchIndex + 1 + count) % count);
  }, [currentSearchIndex, focusSearchMatch]);

  const handleSearchPrev = useCallback(() => {
    const count = searchRangesRef.current.length;
    if (count === 0) return;
    focusSearchMatch((currentSearchIndex - 1 + count) % count);
  }, [currentSearchIndex, focusSearchMatch]);

  // Highlights search matches via the CSS Custom Highlight API so the
  // read-only rendered content never needs to be mutated directly.
  useEffect(() => {
    if (!supportsHighlightApi) return;
    const highlights = CSS.highlights;

    if (!isSearchOpen || showRaw || !contentRef.current || !searchQuery.trim()) {
      clearSearchHighlights();
      return;
    }

    const query = searchQuery.trim().toLowerCase();
    const walker = document.createTreeWalker(contentRef.current, NodeFilter.SHOW_TEXT);
    const ranges: Range[] = [];
    let node: Node | null = walker.nextNode();
    while (node) {
      const text = node.textContent ?? '';
      const lowerText = text.toLowerCase();
      let fromIndex = 0;
      let matchIndex = lowerText.indexOf(query, fromIndex);
      while (matchIndex !== -1) {
        const range = new Range();
        range.setStart(node, matchIndex);
        range.setEnd(node, matchIndex + query.length);
        ranges.push(range);
        fromIndex = matchIndex + query.length;
        matchIndex = lowerText.indexOf(query, fromIndex);
      }
      node = walker.nextNode();
    }

    searchRangesRef.current = ranges;
    setSearchMatchCount(ranges.length);

    if (typeof Highlight !== 'undefined') {
      highlights.set('ppmd-search', new Highlight(...ranges));
    }

    if (ranges.length > 0) {
      focusSearchMatch(0);
    } else {
      // Synchronizes the external Highlight API registry with React state
      // for the "no matches" case; single call, no cascading renders.
      // eslint-disable-next-line react-hooks/set-state-in-effect
      resetActiveMatch();
    }
    // focusSearchMatch is intentionally omitted: re-running it on every
    // render would fight the user's manual next/previous navigation.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isSearchOpen, markdown, searchQuery, showRaw, supportsHighlightApi]);

  useEffect(() => {
    if (isSearchOpen) {
      searchInputRef.current?.focus();
    }
  }, [isSearchOpen]);

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

  const handleExportMdClick = useCallback(() => {
    closeExportMenu();
    onExport?.();
  }, [closeExportMenu, onExport]);

  const handleExportExcelClick = useCallback(() => {
    closeExportMenu();
    handleExportExcel();
  }, [closeExportMenu, handleExportExcel]);

  const handleExportPdfClick = useCallback(() => {
    closeExportMenu();
    void handleExportPdf();
  }, [closeExportMenu, handleExportPdf]);

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
            onClick={toggleSearch}
            aria-pressed={isSearchOpen}
            disabled={showRaw}
            aria-label={isSearchOpen ? 'Close document search' : 'Search document'}
            title={showRaw ? 'Switch to Rendered view to search the document' : 'Search document'}
          >
            🔍 Search
          </button>
          <button
            type="button"
            className={styles.toolbarBtn}
            onClick={handleCopy}
            aria-label="Copy Markdown to clipboard"
          >
            📋 Copy
          </button>
          <div className={styles.exportMenuWrapper}>
            <button
              ref={exportButtonRef}
              type="button"
              className={`${styles.toolbarBtn} ${styles.primary}`}
              onClick={toggleExportMenu}
              aria-haspopup="menu"
              aria-expanded={isExportMenuOpen}
              aria-label="Export documentation"
            >
              ⬇️ Export ▾
            </button>
          </div>
        </div>
      </div>
      {isExportMenuOpen && (
        <div
          ref={exportMenuRef}
          className={styles.exportMenu}
          role="menu"
          aria-label="Export format"
          style={{ top: exportMenuPosition.top, right: exportMenuPosition.right }}
        >
          {onExport && (
            <button
              type="button"
              role="menuitem"
              className={styles.exportMenuItem}
              onClick={handleExportMdClick}
            >
              📄 Export .md
            </button>
          )}
          <button
            type="button"
            role="menuitem"
            className={styles.exportMenuItem}
            onClick={handleExportExcelClick}
            title="Export all Markdown tables to Excel"
          >
            📊 Export .xlsx
          </button>
          <button
            type="button"
            role="menuitem"
            className={styles.exportMenuItem}
            onClick={handleExportPdfClick}
            disabled={isExportingPdf}
            title={showRaw
              ? 'Switch to Rendered view for diagram-inclusive PDF export'
              : isExportingPdf
                ? 'Preparing PDF export'
                : 'Export rendered documentation to PDF'}
          >
            {isExportingPdf ? '⏳ Preparing PDF...' : '🧾 Export .pdf'}
          </button>
        </div>
      )}
      {isSearchOpen && !showRaw && (
        <div className={styles.searchBar} role="search">
          <label htmlFor="markdown-search-input" className="sr-only">Search document</label>
          <input
            id="markdown-search-input"
            ref={searchInputRef}
            type="text"
            className={styles.searchInput}
            value={searchQuery}
            onChange={handleSearchQueryChange}
            placeholder="Search document…"
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                if (event.shiftKey) handleSearchPrev(); else handleSearchNext();
              } else if (event.key === 'Escape') {
                toggleSearch();
              }
            }}
          />
          <span className={styles.searchMatchCount} aria-live="polite">
            {searchQuery.trim()
              ? searchMatchCount > 0
                ? `${currentSearchIndex + 1} of ${searchMatchCount}`
                : 'No matches'
              : ''}
          </span>
          <button
            type="button"
            className={styles.toolbarBtn}
            onClick={handleSearchPrev}
            disabled={searchMatchCount === 0}
            aria-label="Previous match"
          >
            ▲
          </button>
          <button
            type="button"
            className={styles.toolbarBtn}
            onClick={handleSearchNext}
            disabled={searchMatchCount === 0}
            aria-label="Next match"
          >
            ▼
          </button>
          <button
            type="button"
            className={styles.toolbarBtn}
            onClick={toggleSearch}
            aria-label="Close search"
          >
            ✕
          </button>
        </div>
      )}
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
