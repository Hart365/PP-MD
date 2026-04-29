/**
 * @file MarkdownViewer.tsx
 * @description Renders Markdown with GFM, Mermaid diagram support, and a
 * raw/rendered toggle. WCAG compliant with accessible diagrams, headings,
 * tables, and copy/export actions.
 */

import { useState, useCallback, useRef, type MouseEvent, type ReactNode } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import rehypeRaw from 'rehype-raw';
import { MermaidDiagram } from './MermaidDiagram';
import styles from './MarkdownViewer.module.css';

export interface MarkdownViewerProps {
  markdown: string;
  title?: string;
  onExport?: () => void;
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
export function MarkdownViewer({ markdown, title, onExport }: MarkdownViewerProps) {
  const [showRaw, setShowRaw] = useState(false);
  const [copyMsg, setCopyMsg] = useState('');
  const contentRef = useRef<HTMLDivElement>(null);

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

  const components = {
    code: ({ className, children, ...props }: any) => {
      const lang = /language-(\w+)/.exec(className)?.[1] ?? '';
      if (lang !== 'mermaid') {
        return <code className={className} {...props}>{children}</code>;
      }
      const src = String(children).replace(/\n$/, '');
      const cap = src.match(/%%\s*(.+?)\s*%%/)?.[1] ?? 'Diagram';
      return <MermaidDiagram chart={src} caption={cap} />;
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
  };

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

  return (
    <section
      className={styles.viewer}
      aria-label={title ? `Documentation for ${title}` : 'Generated documentation'}
    >
      <div className={styles.toolbar} role="toolbar" aria-label="Documentation actions">
        {title && <h2 className={styles.viewerTitle}>{title}</h2>}
        <div className={styles.actions}>
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
          <ReactMarkdown
            remarkPlugins={[remarkGfm]}
            rehypePlugins={[rehypeRaw]}
            components={components}
          >
            {markdown}
          </ReactMarkdown>
        </div>
      )}
    </section>
  );
}
