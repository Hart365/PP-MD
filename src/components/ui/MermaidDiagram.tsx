import { useCallback, useEffect, useId } from 'react';
import { MermaidDiagram as MermaidRenderer, type MermaidDiagramProps as RendererProps } from './MermaidRenderer';

export { renderMermaidToSvg, forceRenderAllMountedDiagrams } from './MermaidRenderer';

export interface MermaidDiagramProps extends RendererProps {
  mermaidGroupKey?: string;
  onRenderStateChange?: (groupKey: string, diagramId: string, state: 'rendering' | 'done') => void;
}

export function MermaidDiagram({ mermaidGroupKey = 'default', onRenderStateChange, onRendered, ...props }: MermaidDiagramProps) {
  const diagramId = useId();
  useEffect(() => {
    onRenderStateChange?.(mermaidGroupKey, diagramId, 'rendering');
  }, [diagramId, mermaidGroupKey, onRenderStateChange, props.chart, props.theme]);
  const handleRendered = useCallback((renderId: string) => {
    onRenderStateChange?.(mermaidGroupKey, diagramId, 'done');
    onRendered?.(renderId);
  }, [diagramId, mermaidGroupKey, onRenderStateChange, onRendered]);
  return <MermaidRenderer {...props} onRendered={handleRendered} />;
}