/**
 * @file MermaidDiagram.tsx
 * @description Renders a Mermaid diagram with zoom controls, fullscreen support,
 * and accessible fallback. Dynamically imports Mermaid to keep bundle small.
 * WCAG: 1.1.1 (text alternative), 4.1.3 (status/error messages), 1.4.3 (contrast).
 */

import { useEffect, useRef, useState, useId, useMemo } from 'react';

const mermaidCfg = {
  startOnLoad: false,
  theme: 'neutral' as const,
  securityLevel: 'strict' as const,
  fontFamily: "'Segoe UI', system-ui, sans-serif",
  themeVariables: { fontFamily: "'Segoe UI', system-ui, sans-serif", fontSize: '11px' },
  flowchart: { useMaxWidth: false, htmlLabels: false },
  er: { useMaxWidth: false },
};

const btnStyle = {
  padding: '0.35rem 0.65rem',
  background: 'var(--color-surface)',
  color: 'var(--color-text-primary)',
  border: '1px solid var(--color-border)',
  borderRadius: 'var(--border-radius-md)',
  fontSize: '0.8125rem',
  fontWeight: 500,
  cursor: 'pointer',
  transition: 'all var(--transition-fast)',
} as const;

export interface MermaidDiagramProps {
  chart: string;
  caption?: string;
}

/**
 * Renders a Mermaid diagram with zoom/pan/fullscreen controls.
 */
export function MermaidDiagram({ chart, caption = 'Diagram' }: MermaidDiagramProps) {
  const figRef = useRef<HTMLElement>(null);
  const ctrRef = useRef<HTMLDivElement>(null);
  const scrlRef = useRef<HTMLDivElement>(null);
  const [error, setError] = useState('');
  const [rendered, setRendered] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [fs, setFs] = useState(false);
  const [w, setW] = useState<number | null>(null);
  const pct = useMemo(() => `${Math.round(zoom * 100)}%`, [zoom]);
  const id = useId().replace(/:/g, '_');

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
    let active = true;

    const render = async () => {
      try {
        const m = (await import('mermaid')).default;
        m.initialize(mermaidCfg);
        if (!active) return;
        setError('');
        setRendered(false);
        const { svg } = await m.render(`m_${id}`, chart);

        if (!active || !ctrRef.current) return;
        ctrRef.current.innerHTML = svg;

        const svgEl = ctrRef.current.querySelector('svg');
        if (!svgEl) return;

        svgEl.setAttribute('role', 'img');
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
      }
    };

    render();
    return () => { active = false; };
  }, [chart, caption, id]);

  const toolbarStyle = {
    display: 'flex',
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

