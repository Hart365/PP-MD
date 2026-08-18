import { render } from '@testing-library/react';
import { describe, expect, it, vi, beforeEach } from 'vitest';
import { MarkdownViewer } from '../components/ui/MarkdownViewer';
import { MermaidDiagram } from '../components/ui/MermaidDiagram';

vi.mock('../components/ui/MermaidDiagram', () => ({
  MermaidDiagram: vi.fn(() => <div data-testid="mermaid-diagram" />),
}));

describe('MarkdownViewer', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('preserves Mermaid line breaks from fenced code blocks', () => {
    render(
      <MarkdownViewer
        markdown={['```mermaid', 'flowchart LR', 'A-->B', 'linkStyle 0 stroke:#16a34a', '```'].join('\n')}
        title="Document"
      />,
    );

    expect(MermaidDiagram).toHaveBeenCalled();
    const props = vi.mocked(MermaidDiagram).mock.calls.at(-1)?.[0] as { chart?: string } | undefined;
    expect(props?.chart).toContain('flowchart LR\nA-->B\nlinkStyle 0 stroke:#16a34a');
    expect(props?.chart).not.toContain('flowchart LR,A-->B');
  });

  it('repairs merged linkStyle directives in Mermaid fences', () => {
    render(
      <MarkdownViewer
        markdown={[
          '```mermaid',
          'flowchart LR',
          'A-->B',
          'linkStyle 0 stroke:#0f766elinkStyle 1 stroke:#2563eb',
          '```',
        ].join('\n')}
        title="Document"
      />,
    );

    expect(MermaidDiagram).toHaveBeenCalled();
    const props = vi.mocked(MermaidDiagram).mock.calls.at(-1)?.[0] as { chart?: string } | undefined;
    expect(props?.chart).toContain('linkStyle 0 stroke:#0f766e\nlinkStyle 1 stroke:#2563eb');
  });
});
