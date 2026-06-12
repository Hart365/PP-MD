import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import App from '../App';

vi.mock('../components/SolutionSidebar', () => ({
  SolutionSidebar: () => null,
}));

vi.mock('../components/ui/DropZone', () => ({
  DropZone: () => <div data-testid="drop-zone" />,
}));

vi.mock('../components/ui/ThemeToggle', () => ({
  ThemeToggle: () => null,
}));

vi.mock('../components/ui/UpdateChecker', () => ({
  UpdateChecker: () => null,
}));

describe('settings controls and migration', () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it('migrates legacy saved configuration shape to schema v2 defaults', async () => {
    localStorage.setItem(
      'pp-md-doc-configurations',
      JSON.stringify([
        {
          id: 'legacy-1',
          name: 'Legacy Config',
          client: 'Contoso',
          project: 'Legacy Project',
          contract: '',
          sow: '',
          sprint: '',
          releaseDate: '',
        },
      ]),
    );

    render(<App />);

    const configSelect = await screen.findByLabelText(/select document configuration/i);
    fireEvent.change(configSelect, { target: { value: 'legacy-1' } });

    await waitFor(() => {
      expect((screen.getByLabelText(/select attribute selection mode/i) as HTMLSelectElement).value).toBe('all');
      expect((screen.getByLabelText(/include default columns/i) as HTMLInputElement).checked).toBe(true);
    });
  });

  it('applies validation rule for unsupported summary-mode combinations', async () => {
    render(<App />);

    const detailSelect = await screen.findByLabelText(/select documentation detail level/i);
    const attributeModeSelect = screen.getByLabelText(/select attribute selection mode/i);

    fireEvent.change(attributeModeSelect, { target: { value: 'option-set-focused' } });
    fireEvent.change(detailSelect, { target: { value: 'summary' } });

    await waitFor(() => {
      expect((attributeModeSelect as HTMLSelectElement).value).toBe('all');
      expect(screen.getByRole('status').textContent || '').toMatch(/attribute selection mode is only supported in detailed mode/i);
    });
  });

  it('persists metadata settings when saving a configuration', async () => {
    render(<App />);

    fireEvent.click(await screen.findByLabelText(/include audit info/i));
    fireEvent.change(screen.getByLabelText(/select attribute selection mode/i), {
      target: { value: 'manually-selected' },
    });
    fireEvent.change(screen.getByLabelText(/manual attribute schema names/i), {
      target: { value: 'new_name, new_status' },
    });
    fireEvent.change(screen.getByLabelText(/configuration name/i), {
      target: { value: 'Phase2 Persisted' },
    });
    fireEvent.click(screen.getByRole('button', { name: /save configuration/i }));

    await waitFor(() => {
      const saved = JSON.parse(localStorage.getItem('pp-md-doc-configurations') || '[]') as Array<Record<string, unknown>>;
      expect(saved.length).toBe(1);
      expect(saved[0].schemaVersion).toBe(2);
      expect((saved[0].documentationSettings as { metadata?: { includeAuditInfo?: boolean } }).metadata?.includeAuditInfo).toBe(false);
      expect((saved[0].documentationSettings as { metadata?: { attributeSelectionMode?: string } }).metadata?.attributeSelectionMode).toBe('manually-selected');
      expect((saved[0].documentationSettings as { metadata?: { manuallySelectedAttributes?: string[] } }).metadata?.manuallySelectedAttributes).toEqual(['new_name', 'new_status']);
    });
  });
});
