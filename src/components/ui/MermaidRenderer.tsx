/**
 * @file MermaidDiagram.tsx
 * @description Renders a Mermaid diagram with zoom controls, fullscreen support,
 * and accessible fallback. Dynamically imports Mermaid to keep bundle small.
 * WCAG: 1.1.1 (text alternative), 4.1.3 (status/error messages), 1.4.3 (contrast).
 */

import { useEffect, useRef, useState, useId, useMemo } from 'react';
import { LruCache } from '../../utils/lruCache';

const mermaidCfg = {
  startOnLoad: false,
  securityLevel: 'strict' as const,
  fontFamily: "'Segoe UI', system-ui, sans-serif",
  themeVariables: { fontFamily: "'Segoe UI', system-ui, sans-serif", fontSize: '11px' },
  flowchart: { useMaxWidth: false, htmlLabels: false },
  er: { useMaxWidth: false },
};

const btnStyle = {
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  height: 'var(--control-height)',
  minWidth: 'var(--control-height)',
  padding: '0 0.65rem',
  background: 'var(--color-surface)',
  color: 'var(--color-text-primary)',
  border: '1px solid var(--color-border)',
  borderRadius: 'var(--control-radius)',
  fontFamily: 'inherit',
  lineHeight: 1,
  fontSize: '0.8125rem',
  fontWeight: 500,
  cursor: 'pointer',
  transition: 'all var(--transition-fast)',
} as const;

/**
 * Serializes heavy `mermaid.render()` calls across every diagram instance on
 * the page, yielding to the browser between each one. Without this, a
 * document with many diagrams (e.g. a consolidated multi-solution report)
 * fires them all in the same tick and freezes/locks the viewer.
 */
let renderQueue: Promise<void> = Promise.resolve();
function scheduleRender<T>(task: () => Promise<T>): Promise<T> {
  const result = renderQueue.then(() => new Promise<T>((resolve, reject) => {
    const run = () => { task().then(resolve, reject); };
    if (typeof requestIdleCallback === 'function') {
      requestIdleCallback(run, { timeout: 500 });
    } else {
      setTimeout(run, 0);
    }
  }));
  renderQueue = result.then(() => undefined, () => undefined);
  return result;
}

/**
 * Caches rendered SVG markup by chart+theme so re-mounting a previously
 * rendered diagram (raw/rendered toggle, switching documents and back, etc.)
 * is instant instead of re-running mermaid's layout engine. Stores the
 * original render id so it can be rewritten to each consuming instance's own
 * id, avoiding duplicate DOM ids if the same diagram appears more than once.
 * Bounded (least recently used) so long sessions cannot grow memory without limit.
 */
const MERMAID_SVG_CACHE_MAX_ENTRIES = 100;
const svgCache = new LruCache<string, { svg: string; sourceId: string }>(MERMAID_SVG_CACHE_MAX_ENTRIES);

let offscreenRenderCounter = 0;

/**
 * Renders a Mermaid chart to SVG markup outside of the React tree (used by
 * batch exports). Shares the viewer's serialized render queue and config so
 * it never competes with on-screen diagrams for the main thread.
 */
export async function renderMermaidToSvg(chart: string, theme: MermaidDiagramProps['theme'] = 'neutral'): Promise<string> {
  const m = (await import('mermaid')).default;
  offscreenRenderCounter += 1;
  const renderId = `m_export_${offscreenRenderCounter}`;
  const { svg } = await scheduleRender(() => {
    m.initialize({ ...mermaidCfg, theme });
    return m.render(renderId, chart);
  });
  return svg;
}

export interface MermaidDiagramProps {
  chart: string;
  caption?: string;
  /** Called once this diagram finishes rendering (success or failure), with its unique id. */
  onRendered?: (id: string) => void;
  /** Mermaid built-in colour theme (defaults to 'neutral'). */
  theme?: 'neutral' | 'default' | 'dark' | 'forest' | 'base';
  /** Bypasses the viewport-proximity gate so this diagram renders immediately (e.g. before PDF export). */
  forceRender?: boolean;
}

/**
 * Every currently-mounted diagram registers an "activate" callback here.
 * `forceRenderAllMountedDiagrams()` lets a caller (e.g. PDF export) tell every
 * already-mounted diagram to render immediately without changing the
 * `components` object passed to `react-markdown` — mutating that object
 * would remount (and re-parse) the entire document tree, which is what
 * previously caused a multi-second UI freeze on the first PDF export click.
 */
const activeDiagramActivators = new Set<() => void>();
export function forceRenderAllMountedDiagrams(): void {
  activeDiagramActivators.forEach((activate) => activate());
}

/**
 * Low-priority background render scheduler. Mounted diagrams register here and
 * are rendered one at a time, only while the browser is idle and the user is
 * not scrolling. Diagrams currently in the viewport take priority; otherwise
 * diagrams render top-down in document order. Scrolling pauses the queue and
 * it resumes shortly after scrolling stops.
 */
interface ScheduledDiagram {
  el: HTMLElement;
  activate: () => void;
  visible: boolean;
  done: boolean;
}

const SCROLL_IDLE_MS = 200;
const BETWEEN_RENDERS_MS = 120;
const scheduledDiagrams = new Set<ScheduledDiagram>();
let inFlightDiagram: ScheduledDiagram | null = null;
let isUserScrolling = false;
let scrollIdleTimer: ReturnType<typeof setTimeout> | undefined;
let pumpScheduled = false;
let visibilityObserver: IntersectionObserver | null = null;
const diagramsByElement = new WeakMap<Element, ScheduledDiagram>();

function handleAnyScroll(): void {
  isUserScrolling = true;
  if (scrollIdleTimer) clearTimeout(scrollIdleTimer);
  scrollIdleTimer = setTimeout(() => {
    isUserScrolling = false;
    pumpRenderQueue();
  }, SCROLL_IDLE_MS);
}

function compareDocumentOrder(a: ScheduledDiagram, b: ScheduledDiagram): number {
  if (a.el === b.el) return 0;
  return a.el.compareDocumentPosition(b.el) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
}

function runNextRender(): void {
  pumpScheduled = false;
  if (isUserScrolling || inFlightDiagram) return;
  const pending = Array.from(scheduledDiagrams).filter((d) => !d.done && d.el.isConnected);
  if (pending.length === 0) return;
  const visible = pending.filter((d) => d.visible).sort(compareDocumentOrder);
  const next = visible[0] ?? pending.sort(compareDocumentOrder)[0];
  inFlightDiagram = next;
  next.activate();
}

function pumpRenderQueue(delayMs = 0): void {
  if (pumpScheduled) return;
  pumpScheduled = true;
  const schedule = () => {
    if (typeof requestIdleCallback === 'function') {
      requestIdleCallback(runNextRender, { timeout: 2000 });
    } else {
      setTimeout(runNextRender, 50);
    }
  };
  if (delayMs > 0) setTimeout(schedule, delayMs);
  else schedule();
}

function completeScheduledDiagram(entry: ScheduledDiagram): void {
  if (entry.done) return;
  entry.done = true;
  if (inFlightDiagram === entry) inFlightDiagram = null;
  pumpRenderQueue(BETWEEN_RENDERS_MS);
}

function registerScheduledDiagram(el: HTMLElement, activate: () => void): ScheduledDiagram {
  const entry: ScheduledDiagram = { el, activate, visible: false, done: false };
  if (scheduledDiagrams.size === 0 && typeof document !== 'undefined') {
    document.addEventListener('scroll', handleAnyScroll, { capture: true, passive: true });
  }
  scheduledDiagrams.add(entry);
  diagramsByElement.set(el, entry);
  if (typeof IntersectionObserver !== 'undefined') {
    visibilityObserver ??= new IntersectionObserver((observed) => {
      let becameVisible = false;
      observed.forEach((item) => {
        const target = diagramsByElement.get(item.target);
        if (!target) return;
        target.visible = item.isIntersecting;
        if (item.isIntersecting && !target.done) becameVisible = true;
      });
      if (becameVisible) pumpRenderQueue();
    }, { rootMargin: '100px 0px' });
    visibilityObserver.observe(el);
  }
  pumpRenderQueue();
  return entry;
}

function unregisterScheduledDiagram(entry: ScheduledDiagram): void {
  scheduledDiagrams.delete(entry);
  diagramsByElement.delete(entry.el);
  visibilityObserver?.unobserve(entry.el);
  if (inFlightDiagram === entry) inFlightDiagram = null;
  if (scheduledDiagrams.size === 0) {
    if (typeof document !== 'undefined') {
      document.removeEventListener('scroll', handleAnyScroll, { capture: true });
    }
    visibilityObserver?.disconnect();
    visibilityObserver = null;
  } else {
    pumpRenderQueue();
  }
}

/**
 * Renders a Mermaid diagram with zoom/pan/fullscreen controls.
 */
export function MermaidDiagram({ chart, caption = 'Diagram', onRendered, theme = 'neutral', forceRender = false }: MermaidDiagramProps) {
  const figRef = useRef<HTMLElement>(null);
  const ctrRef = useRef<HTMLDivElement>(null);
  const scrlRef = useRef<HTMLDivElement>(null);
  const [error, setError] = useState('');
  const [rendered, setRendered] = useState(false);
  const [isNearViewport, setIsNearViewport] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [fs, setFs] = useState(false);
  const [w, setW] = useState<number | null>(null);
  const pct = useMemo(() => `${Math.round(zoom * 100)}%`, [zoom]);
  const id = useId().replace(/:/g, '_');

  // Register with the module-level activator set so an external caller (PDF
  // export) can force this diagram to render without touching react-markdown's
  // `components` prop (see `forceRenderAllMountedDiagrams` above).
  useEffect(() => {
    const activate = () => setIsNearViewport(true);
    activeDiagramActivators.add(activate);
    return () => {
      activeDiagramActivators.delete(activate);
    };
  }, []);

  // Render via the low-priority background scheduler: one diagram at a time,
  // only while idle and not scrolling, visible diagrams first, then top-down.
  const scheduledRef = useRef<ScheduledDiagram | null>(null);
  useEffect(() => {
    if (forceRender) {
      queueMicrotask(() => setIsNearViewport(true));
      return;
    }
    const el = figRef.current;
    if (!el) return;
    const entry = registerScheduledDiagram(el, () => setIsNearViewport(true));
    scheduledRef.current = entry;
    return () => {
      unregisterScheduledDiagram(entry);
      scheduledRef.current = null;
    };
  }, [forceRender]);

  const fitZoom = () => {
    if (!scrlRef.current || !w) return;
    const pw = scrlRef.current.clientWidth - 16;
    setZoom(Math.max(0.1, Math.min(8, (pw / w) * 0.97)));
  };

  const toggleFs = async () => {
    if (!figRef.current) return;
    try {
      if (document.fullscreenElement === figRef.current) {
        await document.exitFullscreen();
      } else {
        await figRef.current.requestFullscreen();
      }
    } catch {
      // noop
    }
  };

  useEffect(() => {
    const handleFsChange = () => {
      setFs(document.fullscreenElement === figRef.current);
    };
    document.addEventListener('fullscreenchange', handleFsChange);
    return () => document.removeEventListener('fullscreenchange', handleFsChange);
  }, []);

  useEffect(() => {
    if (!isNearViewport) return;
    let active = true;
    const cacheKey = `${theme}::${chart}`;

    const render = async () => {
      try {
        setError('');
        setRendered(false);

        const cached = svgCache.get(cacheKey);
        let svg: string;
        if (cached) {
          svg = cached.sourceId === id ? cached.svg : cached.svg.split(cached.sourceId).join(id);
        } else {
          const m = (await import('mermaid')).default;
          m.initialize({ ...mermaidCfg, theme });
          if (!active) return;
          const renderId = `m_${id}`;
          ({ svg } = await scheduleRender(() => m.render(renderId, chart)));
          svgCache.set(cacheKey, { svg, sourceId: id });
        }

        if (!active || !ctrRef.current) return;
        ctrRef.current.innerHTML = svg;

        const svgEl = ctrRef.current.querySelector('svg');
        if (!svgEl) return;

        svgEl.setAttribute('role', 'img');
        svgEl.setAttribute('data-pdf-mermaid', 'true');
        svgEl.setAttribute('aria-label', caption);
        svgEl.setAttribute('preserveAspectRatio', 'xMinYMin meet');

        // Extract viewBox and set dimensions
        const vb = svgEl.getAttribute('viewBox')?.trim()?.split(/\s+/);
        if (vb && vb.length === 4) {
          const vw = Number(vb[2]);
          if (Number.isFinite(vw) && vw > 0) {
            svgEl.style.width = `${vw}px`;
            svgEl.style.minWidth = `${vw}px`;
            setW(vw);
          }
          const vh = Number(vb[3]);
          if (Number.isFinite(vh) && vh > 0) {
            svgEl.style.height = `${vh}px`;
          }
        }

        svgEl.removeAttribute('width');
        svgEl.removeAttribute('height');
        Object.assign(svgEl.style, { maxWidth: 'none', display: 'block', margin: 0, fontSize: '8pt' });

        // Enforce minimum text size
        svgEl.querySelectorAll('text').forEach((el) => {
          const sz = el.getAttribute('font-size');
          if (!sz) {
            el.setAttribute('font-size', '8pt');
          } else {
            const num = parseFloat(sz);
            const isPt = sz.toLowerCase().includes('pt');
            const isPx = sz.toLowerCase().includes('px');
            if ((isPt && num < 8) || (isPx && num < 10.67)) {
              el.setAttribute('font-size', '8pt');
            }
          }
        });

        setRendered(true);
      } catch (e) {
        if (active) {
          setError(`Failed to render: ${(e as Error).message}`);
        }
      } finally {
        if (active) onRendered?.(id);
        if (scheduledRef.current) completeScheduledDiagram(scheduledRef.current);
      }
    };

    render();
    return () => { active = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chart, caption, id, isNearViewport, theme]);

  const toolbarStyle = {
    display: 'flex',
    flexWrap: 'wrap' as const,
    alignItems: 'center',
    justifyContent: 'flex-end',
    gap: '0.5rem',
    marginBottom: '0.5rem',
    position: 'sticky' as const,
    top: 0,
    zIndex: 1,
    background: 'var(--color-surface)',
    padding: '0.25rem 0',
  };

  return (
    <figure
      ref={figRef}
      data-pdf-mermaid="true"
      aria-label={caption}
      style={{
        margin: '1.5rem 0',
        width: '100%',
        height: fs ? '100vh' : 'auto',
        display: 'flex',
        flexDirection: 'column',
        overflow: 'visible',
        background: 'var(--color-surface)',
        color: 'var(--color-text-primary)',
        padding: fs ? '0.75rem' : 0,
        boxSizing: 'border-box',
      }}
    >
      <div style={toolbarStyle}>
        <span style={{ fontSize: '0.8rem', color: 'var(--color-text-muted)' }}>{pct}</span>
        <button
          type="button"
          onClick={() => setZoom((c) => Math.max(0.1, c - 0.1))}
          aria-label="Zoom out"
          style={btnStyle}
          onMouseEnter={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-accent)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 0 0 2px rgba(37,99,235,0.1)';
          }}
          onMouseLeave={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-border)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = 'none';
          }}
        >
          −
        </button>
        <button
          type="button"
          onClick={() => setZoom(1)}
          aria-label="Reset zoom"
          style={btnStyle}
          onMouseEnter={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-accent)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 0 0 2px rgba(37,99,235,0.1)';
          }}
          onMouseLeave={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-border)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = 'none';
          }}
        >
          1×
        </button>
        <button
          type="button"
          onClick={fitZoom}
          aria-label="Fit to width"
          disabled={!w}
          style={btnStyle}
          onMouseEnter={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-accent)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 0 0 2px rgba(37,99,235,0.1)';
          }}
          onMouseLeave={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-border)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = 'none';
          }}
        >
          Fit
        </button>
        <button
          type="button"
          onClick={() => setZoom((c) => Math.min(8, c + 0.1))}
          aria-label="Zoom in"
          style={btnStyle}
          onMouseEnter={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-accent)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 0 0 2px rgba(37,99,235,0.1)';
          }}
          onMouseLeave={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-border)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = 'none';
          }}
        >
          +
        </button>
        <button
          type="button"
          onClick={toggleFs}
          aria-label={fs ? 'Exit fullscreen' : 'Fullscreen'}
          style={btnStyle}
          onMouseEnter={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-accent)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 0 0 2px rgba(37,99,235,0.1)';
          }}
          onMouseLeave={(e) => {
            (e.currentTarget as HTMLButtonElement).style.borderColor = 'var(--color-border)';
            (e.currentTarget as HTMLButtonElement).style.boxShadow = 'none';
          }}
        >
          {fs ? '✕ Exit FS' : '⛶ Full Screen'}
        </button>
      </div>

      <div
        ref={scrlRef}
        style={{
          flex: fs ? 1 : undefined,
          minHeight: fs ? 0 : rendered ? 'auto' : '60px',
          overflowX: 'auto',
          overflowY: 'auto',
          border: '1px solid var(--color-code-border)',
          borderRadius: 'var(--border-radius-md)',
          padding: '0.5rem',
          maxHeight: fs ? 'none' : '85vh',
          background: 'var(--color-surface)',
          cursor: zoom < 1 ? 'zoom-in' : zoom > 1 ? 'grab' : 'default',
        }}
      >
        <div
          style={{
            transform: `scale(${zoom})`,
            transformOrigin: 'top left',
            width: 'max-content',
            minWidth: 'max-content',
          }}
        >
          <div ref={ctrRef} aria-hidden={!rendered} />
        </div>
      </div>

      {error && (
        <div
          role="alert"
          style={{
            padding: '0.75rem 1rem',
            background: 'var(--color-error-bg)',
            color: 'var(--color-error)',
            borderRadius: 'var(--border-radius-md)',
            fontSize: '0.875rem',
            marginTop: '0.5rem',
          }}
        >
          ⚠️ {error}
        </div>
      )}

      <details style={{ marginTop: '0.5rem' }}>
        <summary
          style={{
            fontSize: '0.8125rem',
            color: 'var(--color-text-muted)',
            cursor: 'pointer',
          }}
        >
          Source ({caption})
        </summary>
        <pre
          aria-label={`Source: ${caption}`}
          style={{
            marginTop: '0.5rem',
            fontSize: '0.75rem',
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
            background: 'var(--color-code-bg)',
            padding: '0.75rem',
            borderRadius: 'var(--border-radius-md)',
            border: '1px solid var(--color-code-border)',
          }}
        >
          <code>{chart}</code>
        </pre>
      </details>

      {caption && (
        <figcaption
          style={{
            textAlign: 'center',
            fontSize: '0.875rem',
            color: 'var(--color-text-secondary)',
            marginTop: '0.25rem',
          }}
        >
          {caption}
        </figcaption>
      )}
    </figure>
  );
}
