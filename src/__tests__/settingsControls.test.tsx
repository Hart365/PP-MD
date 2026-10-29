import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import App, { rebuildResults } from '../App';
import { DEFAULT_DOCUMENTATION_SETTINGS } from '../generator/markdownGenerator';
import { AttributeType, type ParsedSolution } from '../types/solution';

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

/** Expand a collapsible section by its heading text. */
async function expandSection(heading: RegExp) {
  const btn = await screen.findByRole('button', { name: heading });
  if (btn.getAttribute('aria-expanded') === 'false') {
    fireEvent.click(btn);
  }
}

describe('settings controls and migration', () => {
  const documentContext = { client: '', project: '', contract: '', sow: '', sprint: '', releaseDate: '' };
  const solution: ParsedSolution = {
    metadata: { uniqueName: 'sample', displayName: 'Sample', version: '1', publisherName: 'Test', isManaged: false, dependencies: [], componentInventory: [] },
    entities: [{ name: 'new_case', logicalName: 'new_case', isCustom: true, description: 'Large document '.repeat(4000), attributes: [
      { name: 'new_name', type: AttributeType.String, required: false, isCustom: true },
    ], relationships: [] }],
    optionSets: [], forms: [], views: [], processes: [], apps: [], agents: [], aiModels: [], desktopFlows: [],
    dataflows: [], customApis: [], offlineProfiles: [], webResources: [], securityRoles: [], fieldSecurityProfiles: [],
    connectionReferences: [], environmentVariables: [], emailTemplates: [], reports: [], dashboards: [], pluginAssemblies: [], warnings: [],
  };

  for (const count of [1, 2]) {
    for (const separateDiagramsDocument of [false, true]) {
      it(`honours separate diagrams=${separateDiagramsDocument} for ${count} large solution documents`, () => {
        const inputs = Array.from({ length: count }, (_, index) => ({
          solution: { ...solution, metadata: { ...solution.metadata, uniqueName: `sample_${index}` } },
          markdown: '', fileName: `Sample ${index}.md`,
        }));
        const results = rebuildResults(inputs, 'detailed-relationships', documentContext, { ...DEFAULT_DOCUMENTATION_SETTINGS, separateDiagramsDocument });
        const companions = results.filter((result) => result.isDiagramCompanion);
        expect(companions).toHaveLength(separateDiagramsDocument ? count + (count > 1 ? 1 : 0) : 0);
        const main = results.filter((result) => !result.isReport && !result.isDiagramCompanion);
        expect(main).toHaveLength(count + (count > 1 ? 1 : 0));
        main.forEach((result) => {
          expect(result.markdown.length).toBeGreaterThan(45000);
          expect(result.markdown.includes('```mermaid')).toBe(!separateDiagramsDocument);
        });
      });
    }
  }

  it('creates no companion documents when diagrams are excluded even if separation is enabled', () => {
    const results = rebuildResults([{ solution, markdown: '', fileName: 'Sample.md' }], 'compact', documentContext, {
      ...DEFAULT_DOCUMENTATION_SETTINGS, includeDiagrams: false, separateDiagramsDocument: true,
    });
    expect(results.some((result) => result.isDiagramCompanion)).toBe(false);
    expect(results.every((result) => !result.markdown.includes('```mermaid'))).toBe(true);
  });

  for (const createConsolidatedDocument of [true, false]) {
    for (const includeDiagrams of [true, false]) {
      for (const separateDiagramsDocument of [true, false]) {
        it(`honours combined=${createConsolidatedDocument}, diagrams=${includeDiagrams}, separate=${separateDiagramsDocument}`, () => {
          const inputs = [solution, { ...solution, metadata: { ...solution.metadata, uniqueName: 'other' } }]
            .map((item) => ({ solution: item, markdown: '', fileName: `${item.metadata.uniqueName}.md` }));
          const results = rebuildResults(inputs, 'compact', documentContext, {
            ...DEFAULT_DOCUMENTATION_SETTINGS, createConsolidatedDocument, includeDiagrams, separateDiagramsDocument,
          });
          expect(results.some((result) => result.isConsolidated)).toBe(createConsolidatedDocument);
          expect(results.filter((result) => result.isDiagramCompanion)).toHaveLength(includeDiagrams && separateDiagramsDocument ? 2 + Number(createConsolidatedDocument) : 0);
          if (!includeDiagrams) expect(results.every((result) => !result.markdown.includes('```mermaid'))).toBe(true);
        });
      }
    }
  }

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

    await expandSection(/document options/i);
    await expandSection(/table options/i);

    await waitFor(() => {
      expect((screen.getByLabelText(/select attribute selection mode/i) as HTMLSelectElement).value).toBe('all');
      expect((screen.getByLabelText(/include default columns/i) as HTMLInputElement).checked).toBe(true);
    });
  });

  it('applies validation rule for unsupported summary-mode combinations', async () => {
    render(<App />);

    await expandSection(/document options/i);

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

    await expandSection(/table options/i);
    await expandSection(/column metadata/i);
    fireEvent.click(await screen.findByLabelText(/^audited$/i));

    await expandSection(/security role options/i);
    fireEvent.click(screen.getByLabelText(/only include tables in current solution/i));
    fireEvent.click(screen.getByLabelText(/only include custom tables in security roles/i));

    await expandSection(/diagram options/i);
    fireEvent.click(screen.getByLabelText(/generate companion diagrams document/i));
    fireEvent.click(screen.getByLabelText('Include diagrams', { exact: true }));
    fireEvent.click(screen.getByLabelText('Create All Selected Solutions document', { exact: true }));
    fireEvent.click(screen.getByLabelText('Include missing import dependencies', { exact: true }));
    fireEvent.change(screen.getByLabelText('Diagram colour theme'), { target: { value: 'forest' } });

    await expandSection(/document options/i);
    fireEvent.change(screen.getByLabelText(/select attribute selection mode/i), {
      target: { value: 'custom-only' },
    });

    fireEvent.change(screen.getByLabelText(/configuration name/i), {
      target: { value: 'Phase2 Persisted' },
    });
    fireEvent.click(screen.getByRole('button', { name: /save configuration/i }));

    await waitFor(() => {
      const saved = JSON.parse(localStorage.getItem('pp-md-doc-configurations') || '[]') as Array<Record<string, unknown>>;
      expect(saved.length).toBe(1);
      expect(saved[0].schemaVersion).toBe(3);
      expect((saved[0].documentationSettings as { metadata?: { includeAuditInfo?: boolean } }).metadata?.includeAuditInfo).toBe(false);
      expect((saved[0].documentationSettings as { metadata?: { attributeSelectionMode?: string } }).metadata?.attributeSelectionMode).toBe('custom-only');
      expect((saved[0].documentationSettings as { securityRoleFilters?: { onlyTablesInCurrentSolution?: boolean; onlyCustomTables?: boolean } }).securityRoleFilters?.onlyTablesInCurrentSolution).toBe(true);
      expect((saved[0].documentationSettings as { securityRoleFilters?: { onlyTablesInCurrentSolution?: boolean; onlyCustomTables?: boolean } }).securityRoleFilters?.onlyCustomTables).toBe(true);
      expect((saved[0].documentationSettings as { separateDiagramsDocument?: boolean }).separateDiagramsDocument).toBe(true);
      expect(saved[0].documentationSettings).toMatchObject({ includeDiagrams: false, createConsolidatedDocument: false, generateMissingDependencies: true, mermaidColourTheme: 'forest' });
    });
  });
});
