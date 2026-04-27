/**
 * @file MarkdownViewer.tsx
 * @description Renders a Markdown string using react-markdown with GitHub
 * Flavored Markdown support.  Mermaid code fence blocks are intercepted and
 * rendered via the {@link MermaidDiagram} component.
 *
 * WCAG compliance:
 *  - 1.1.1 Non-text Content: Mermaid diagrams have text alternatives.
 *  - 1.3.1 Info & Relationships: Semantic HTML headings/lists/tables preserved.
 *  - 2.1.1 Keyboard: All interactive elements (copy button, export) are
 *    keyboard-accessible.
 *  - 4.1.3 Status Messages: Copy confirmation announced via aria-live.
 */

import { useState, useCallback, useRef, type ReactNode } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import rehypeRaw from 'rehype-raw';
import { MermaidDiagram } from './MermaidDiagram';
import styles from './MarkdownViewer.module.css';

export interface MarkdownViewerProps {
  /** The raw Markdown string to render */
  markdown: string;
  /** Optional title shown above the content */
  title?: string;
  /** Callback for exporting the Markdown as a .md file */
  onExport?: () => void;
}

function extractText(node: ReactNode): string {
  if (typeof node === 'string' || typeof node === 'number') return String(node);
  if (Array.isArray(node)) return node.map((child) => extractText(child)).join(' ');
  if (node && typeof node === 'object' && 'props' in node) {
    return extractText((node as { props?: { children?: ReactNode } }).props?.children);
  }
  return '';
}

function headingId(children: ReactNode): string {
  const raw = extractText(children)
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-');
  return raw || 'section';
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

/**
 * Renders Markdown content with Mermaid diagram support, a raw-source toggle,
 * and export/copy actions.
 */
export function MarkdownViewer({ markdown, title, onExport }: MarkdownViewerProps) {
  /** Whether to show raw Markdown source rather than rendered HTML */
  const [showRaw,   setShowRaw]   = useState(false);
  /** Confirmation message after copying to clipboard */
  const [copyMsg,   setCopyMsg]   = useState('');
  const headingCountsRef = useRef<Map<string, number>>(new Map());
  headingCountsRef.current = new Map();

  const uniqueHeadingId = (children: ReactNode): string => {
    const base = headingId(children);
    const counts = headingCountsRef.current;
    const current = counts.get(base) ?? 0;
    counts.set(base, current + 1);
    return current === 0 ? base : `${base}-${current}`;
  };

  const handleCopy = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(markdown);
      setCopyMsg('Copied!');
      setTimeout(() => setCopyMsg(''), 2000);
    } catch {
      setCopyMsg('Copy failed — please select and copy manually.');
      setTimeout(() => setCopyMsg(''), 3000);
    }
  }, [markdown]);

  return (
    <section
      className={styles.viewer}
      aria-label={title ? `Documentation for ${title}` : 'Generated documentation'}
    >
      {/* ── Toolbar ──────────────────────────────────────────────────────── */}
      <div className={styles.toolbar} role="toolbar" aria-label="Documentation actions">
        {title && <h2 className={styles.viewerTitle}>{title}</h2>}

        <div className={styles.actions}>
          {/* Raw/Rendered toggle */}
          <button
            type="button"
            className={styles.toolbarBtn}
            onClick={() => setShowRaw((v) => !v)}
            aria-pressed={showRaw}
            aria-label={showRaw ? 'Switch to rendered view' : 'Switch to raw Markdown view'}
          >
            {showRaw ? '🖼️ Rendered' : '📝 Raw MD'}
          </button>

          {/* Copy */}
          <button
            type="button"
            className={styles.toolbarBtn}
            onClick={handleCopy}
            aria-label="Copy Markdown to clipboard"
          >
            📋 Copy
          </button>

          {/* Export */}
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
        </div>
      </div>

      {/* Copy confirmation — announced to screen readers via aria-live */}
      {copyMsg && (
        <div
          aria-live="polite"
          className={styles.copyToast}
          role="status"
        >
          {copyMsg}
        </div>
      )}

      {/* ── Content area ─────────────────────────────────────────────────── */}
      {showRaw ? (
        /* Raw Markdown source */
        <pre className={styles.rawSource} aria-label="Raw Markdown source">
          <code>{markdown}</code>
        </pre>
      ) : (
        /* Rendered Markdown */
        <div className={`${styles.content} markdown-body`}>
          <ReactMarkdown
            remarkPlugins={[remarkGfm]}
            rehypePlugins={[rehypeRaw]}
            components={{
              /**
               * Override code blocks: intercept ```mermaid``` fences and render
               * them as live diagrams; all other fences render as normal code.
               */
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              code({ className, children, ...props }: any) {
                const language = /language-(\w+)/.exec(className || '')?.[1] ?? '';
                const codeStr  = String(children).replace(/\n$/, '');

                if (language === 'mermaid') {
                  // Extract the title from the Mermaid %% comment if present
                  const titleMatch = codeStr.match(/%%\s*(.+?)\s*%%/);
                  const diagramTitle = titleMatch ? titleMatch[1] : 'Diagram';
                  return <MermaidDiagram chart={codeStr} caption={diagramTitle} />;
                }

                return (
                  <code className={className} {...props}>
                    {children}
                  </code>
                );
              },

              /**
               * Add id anchors to headings so the Table of Contents links work.
               */
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              h1({ children, ...props }: any) {
                const id = uniqueHeadingId(children as ReactNode);
                return <h1 id={id} {...props}>{children}</h1>;
              },
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              h2({ children, ...props }: any) {
                const id = uniqueHeadingId(children as ReactNode);
                return <h2 id={id} {...props}>{children}</h2>;
              },
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              h3({ children, ...props }: any) {
                const id = uniqueHeadingId(children as ReactNode);
                return <h3 id={id} {...props}>{children}</h3>;
              },

              /**
               * Ensure table elements are wrapped in a scrollable container
               * so they don't overflow on narrow viewports (WCAG 1.4.10 Reflow).
               */
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              table({ children, ...props }: any) {
                return (
                  <div style={{ overflowX: 'auto', margin: '1rem 0' }} role="region" aria-label="Table">
                    <table {...props}>{children}</table>
                  </div>
                );
              },
            }}
          >
            {markdown}
          </ReactMarkdown>
        </div>
      )}
    </section>
  );
}
