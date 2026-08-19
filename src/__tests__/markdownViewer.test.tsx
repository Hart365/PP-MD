import { fireEvent, render, screen } from '@testing-library/react';
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

  it('merges Export actions into a single dropdown menu', () => {
    const onExport = vi.fn();
    render(
      <MarkdownViewer
        markdown="# Title\n\nSome content."
        title="Document"
        onExport={onExport}
      />,
    );

    // Individual export buttons should no longer be rendered directly on the toolbar.
    expect(screen.queryByRole('button', { name: /^export \.md$/i })).not.toBeInTheDocument();
    expect(screen.queryByRole('button', { name: /^export \.xlsx$/i })).not.toBeInTheDocument();
    expect(screen.queryByRole('button', { name: /^export \.pdf$/i })).not.toBeInTheDocument();

    const trigger = screen.getByRole('button', { name: /export documentation/i });
    expect(trigger).toHaveAttribute('aria-expanded', 'false');
    expect(screen.queryByRole('menu', { name: /export format/i })).not.toBeInTheDocument();

    fireEvent.click(trigger);

    expect(trigger).toHaveAttribute('aria-expanded', 'true');
    const menu = screen.getByRole('menu', { name: /export format/i });
    expect(menu).toBeInTheDocument();

    fireEvent.click(screen.getByRole('menuitem', { name: /export \.md/i }));
    expect(onExport).toHaveBeenCalledTimes(1);
    expect(screen.queryByRole('menu', { name: /export format/i })).not.toBeInTheDocument();
  });

  it('closes the export menu on Escape', () => {
    render(<MarkdownViewer markdown="# Title" title="Document" onExport={vi.fn()} />);

    fireEvent.click(screen.getByRole('button', { name: /export documentation/i }));
    expect(screen.getByRole('menu', { name: /export format/i })).toBeInTheDocument();

    fireEvent.keyDown(document, { key: 'Escape' });
    expect(screen.queryByRole('menu', { name: /export format/i })).not.toBeInTheDocument();
  });

  it('opens and closes an in-document search bar', () => {
    render(<MarkdownViewer markdown="# Title\n\nSome searchable content." title="Document" />);

    expect(screen.queryByPlaceholderText(/search document/i)).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole('button', { name: /^search document$/i }));
    expect(screen.getByPlaceholderText(/search document/i)).toBeInTheDocument();

    fireEvent.click(screen.getByRole('button', { name: /close search/i }));
    expect(screen.queryByPlaceholderText(/search document/i)).not.toBeInTheDocument();
  });
});
