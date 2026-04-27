/**
 * @file MermaidDiagram.tsx
 * @description Renders a Mermaid diagram from a DSL string using the Mermaid
 * library loaded dynamically.  Provides a text fallback for screen readers
 * (the raw Mermaid source in a <details> block).
 *
 * WCAG compliance:
 *  - 1.1.1 Non-text Content: An accessible text alternative is exposed via
 *    a <details> / <summary> toggle beneath each diagram.
 *  - 4.1.3 Status Messages: Renders a loading and error state.
 *  - 1.4.3 Contrast: Error/info messages use semantic colour tokens.
 */

import { useEffect, useRef, useState, useId, useMemo } from 'react';

export interface MermaidDiagramProps {
  /** Mermaid DSL source string (without the fence markers) */
  chart: string;
  /** Accessible caption / title for the diagram */
  caption?: string;
}

/**
 * Renders a single Mermaid diagram asynchronously.
 * Falls back to the raw source if Mermaid fails to render.
 */
export function MermaidDiagram({ chart, caption = 'Diagram' }: MermaidDiagramProps) {
  const figureRef = useRef<HTMLElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const [error,    setError]    = useState<string>('');
  const [rendered, setRendered] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [svgNaturalWidth, setSvgNaturalWidth] = useState<number | null>(null);
  const zoomPercent = useMemo(() => `${Math.round(zoom * 100)}%`, [zoom]);
  // Generate a unique ID for this diagram instance (required by Mermaid)
  const diagramId = useId().replace(/:/g, '_');

  const fitZoom = () => {
    if (!scrollRef.current || !svgNaturalWidth) return;
    const containerWidth = scrollRef.current.clientWidth - 16; // subtract padding
    const fit = Number(Math.max(0.1, (containerWidth / svgNaturalWidth) * 0.97).toFixed(2));
    setZoom(fit);
  };

  const toggleFullscreen = async () => {
    if (!figureRef.current) return;
    try {
      if (document.fullscreenElement === figureRef.current) {
        await document.exitFullscreen();
      } else {
        await figureRef.current.requestFullscreen();
      }
    } catch {
      // best-effort only
    }
  };

  useEffect(() => {
    const onFullscreenChange = () => {
      setIsFullscreen(document.fullscreenElement === figureRef.current);
    };

    document.addEventListener('fullscreenchange', onFullscreenChange);
    return () => {
      document.removeEventListener('fullscreenchange', onFullscreenChange);
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    setError('');
    setRendered(false);

    const render = async () => {
      try {
        // Dynamic import keeps Mermaid out of the main bundle chunk
        const mermaid = (await import('mermaid')).default;

        mermaid.initialize({
          startOnLoad: false,
          theme:       'neutral',
          securityLevel: 'strict',
          fontFamily: "'Segoe UI', system-ui, sans-serif",
          themeVariables: {
            fontFamily: "'Segoe UI', system-ui, sans-serif",
            // 8pt ~= 10.67px. Use 11px minimum for legibility.
            fontSize: '11px',
          },
          flowchart: { useMaxWidth: false, htmlLabels: false },
          er: { useMaxWidth: false },
        });

        const { svg } = await mermaid.render(`m_${diagramId}`, chart);

        if (cancelled) return;

        if (containerRef.current) {
          // DOMPurify is not available here but Mermaid's strict security
          // level sanitises the SVG output itself.
          containerRef.current.innerHTML = svg;

          // Add role="img" and aria-label to the SVG for AT
          const svgEl = containerRef.current.querySelector('svg');
          if (svgEl) {
            svgEl.setAttribute('role', 'img');
            svgEl.setAttribute('aria-label', caption);
            svgEl.setAttribute('preserveAspectRatio', 'xMinYMin meet');

            // Keep natural diagram dimensions so overflow scrolling and zoom
            // can reveal dense content without shrinking text to fit container width.
            const viewBox = svgEl.getAttribute('viewBox')?.trim();
            if (viewBox) {
              const parts = viewBox.split(/\s+/);
              if (parts.length === 4) {
                const vbWidth = Number(parts[2]);
                const vbHeight = Number(parts[3]);
                if (Number.isFinite(vbWidth) && vbWidth > 0) {
                  svgEl.style.width = `${vbWidth}px`;
                  svgEl.style.minWidth = `${vbWidth}px`;
                  setSvgNaturalWidth(vbWidth);
                }
                if (Number.isFinite(vbHeight) && vbHeight > 0) {
                  svgEl.style.height = `${vbHeight}px`;
                }
              }
            }

            svgEl.removeAttribute('width');
            svgEl.removeAttribute('height');
            svgEl.style.maxWidth = 'none';
            svgEl.style.display = 'block';
            svgEl.style.margin = '0';

            // Enforce minimum label size of 8pt for readability.
            svgEl.style.fontSize = '8pt';
            svgEl.querySelectorAll('text').forEach((textNode) => {
              const textEl = textNode as SVGTextElement;
              const explicit = textEl.getAttribute('font-size');
              if (!explicit) {
                textEl.setAttribute('font-size', '8pt');
                return;
              }

              const parsed = Number.parseFloat(explicit);
              if (!Number.isFinite(parsed)) return;

              const isPt = explicit.toLowerCase().includes('pt');
              const isPx = explicit.toLowerCase().includes('px');
              if (isPt && parsed < 8) textEl.setAttribute('font-size', '8pt');
              if (isPx && parsed < 10.67) textEl.setAttribute('font-size', '8pt');
            });
          }
        }
        setRendered(true);
      } catch (err: unknown) {
        if (!cancelled) {
          setError(`Diagram could not be rendered: ${(err as Error).message}`);
        }
      }
    };

    render();
    return () => { cancelled = true; };
  }, [chart, caption, diagramId]);

  return (
    <figure
      ref={figureRef}
      aria-label={caption}
      style={{
        margin: '1.5rem 0',
        width: '100%',
        overflow: 'visible',
        background: 'var(--color-surface)',
        color: 'var(--color-text-primary)',
        padding: isFullscreen ? '0.75rem' : 0,
      }}
    >
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'flex-end',
          gap: '0.5rem',
          marginBottom: '0.5rem',
          position: 'sticky',
          top: 0,
          zIndex: 1,
          background: 'var(--color-surface)',
          padding: '0.25rem 0',
        }}
      >
        <span style={{ fontSize: '0.8rem', color: 'var(--color-text-muted)' }}>{zoomPercent}</span>
        <button
          type="button"
          onClick={() => setZoom((current) => Math.max(0.1, Number((current - 0.1).toFixed(2))))}
          aria-label="Zoom out diagram"
        >
          −
        </button>
        <button
          type="button"
          onClick={() => setZoom(1)}
          aria-label="Reset diagram zoom to 100%"
        >
          1×
        </button>
        <button
          type="button"
          onClick={fitZoom}
          aria-label="Fit diagram to container width"
          disabled={!svgNaturalWidth}
        >
          Fit
        </button>
        <button
          type="button"
          onClick={() => setZoom((current) => Math.min(8, Number((current + 0.1).toFixed(2))))}
          aria-label="Zoom in diagram"
        >
          +
        </button>
        <button
          type="button"
          onClick={toggleFullscreen}
          aria-label={isFullscreen ? 'Exit diagram fullscreen' : 'Open diagram in fullscreen'}
        >
          {isFullscreen ? 'Exit Full Screen' : 'Full Screen'}
        </button>
      </div>

      {/* The Mermaid SVG will be injected here */}
      <div
        ref={scrollRef}
        style={{
          minHeight: rendered ? 'auto' : '60px',
          overflowX: 'auto',
          overflowY: 'auto',
          border: '1px solid var(--color-code-border)',
          borderRadius: 'var(--border-radius-md)',
          padding: '0.5rem',
          maxHeight: isFullscreen ? 'calc(100vh - 8rem)' : '85vh',
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
          <div ref={containerRef} aria-hidden={!rendered} />
        </div>
      </div>

      {/* Error state */}
      {error && (
        <div
          role="alert"
          style={{
            padding: '0.75rem 1rem',
            background: 'var(--color-error-bg)',
            color: 'var(--color-error)',
            borderRadius: 'var(--border-radius-md)',
            fontSize: '0.875rem',
          }}
        >
          ⚠️ {error}
        </div>
      )}

      {/* Accessible text fallback — always present */}
      <details style={{ marginTop: '0.5rem' }}>
        <summary
          style={{
            fontSize: '0.8125rem',
            color: 'var(--color-text-muted)',
            cursor: 'pointer',
          }}
        >
          View diagram source ({caption})
        </summary>
        <pre
          aria-label={`Mermaid source for: ${caption}`}
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

      {/* Visible caption */}
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
