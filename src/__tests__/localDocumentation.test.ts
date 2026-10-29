import JSZip from 'jszip';
import { parseSolutionZip } from '../parser/solutionParser';
import { consolidateSolutions, DEFAULT_DOCUMENTATION_SETTINGS, fillSolutionGapsFromPeerSolutions, generateMarkdown } from '../generator/markdownGenerator';
import { generateDependencyReportMarkdown, generateDiagnosticsMarkdown } from '../generator/localDocumentation';
import { splitMarkdownByCategory } from '../generator/documentSplitter';

async function sampleSolution(name = 'sample') {
  const zip = new JSZip();
  zip.file('solution.xml', `<ImportExportXml><SolutionManifest><UniqueName>${name}</UniqueName><Version>1</Version><MissingDependencies><MissingDependency><Required type="1" schemaName="account"/><Dependent type="2" schemaName="new_customer" parentSchemaName="new_case"/></MissingDependency></MissingDependencies></SolutionManifest></ImportExportXml>`);
  zip.file('customizations.xml', '<ImportExportXml/>');
  return parseSolutionZip(await zip.generateAsync({ type: 'blob' }));
}

describe('local documentation parity', () => {
  it('fills only matching peer components and never imports peer prerequisites or warnings', async () => {
    const source = await sampleSolution('first');
    source.entities = [{ name: 'new_case', logicalName: 'new_case', isCustom: true, attributes: [], relationships: [] }];
    const peer = await sampleSolution('second');
    peer.entities = [{ ...source.entities[0], description: 'Peer description' }, { ...source.entities[0], name: 'account', logicalName: 'account' }];
    peer.warnings = ['Peer warning'];
    const enriched = fillSolutionGapsFromPeerSolutions(source, [peer]);
    expect(enriched.entities).toHaveLength(1);
    expect(enriched.entities[0].description).toBe('Peer description');
    expect(source.entities[0].description).toBeUndefined();
    expect(enriched.missingDependencies).toEqual(source.missingDependencies);
    expect(enriched.warnings).toEqual([]);
  });

  it('ignores headings inside code fences and gives category files their own contents', () => {
    const categories = splitMarkdownByCategory('# Example\n\n## Power Apps\n\n```text\n## Not a section\n```\n\nApp contents');
    expect(categories).toHaveLength(1);
    expect(categories[0].markdown).toContain('## Table of Contents');
    expect(categories[0].markdown).toContain('## Not a section');
  });
  it('keeps declared prerequisites scoped to their source solutions when consolidating', async () => {
    const first = await sampleSolution('first');
    const second = await sampleSolution('second');
    const consolidated = consolidateSolutions([first, second]);
    expect(consolidated.missingDependencies?.map((dependency) => dependency.solutionUniqueName)).toEqual(['first', 'second']);
    expect(generateDependencyReportMarkdown(consolidated)).toContain('new_case.new_customer');
    expect(first.missingDependencies).toHaveLength(1);
  });

  it('generates separate diagnostics and optional dependency documentation', async () => {
    const solution = await sampleSolution();
    expect(generateDiagnosticsMarkdown(solution)).toBe('');
    expect(generateDiagnosticsMarkdown({ ...solution, warnings: ['Optional file unreadable'] })).toContain('Optional file unreadable');
    const markdown = generateMarkdown(solution, { documentationSettings: {
      ...DEFAULT_DOCUMENTATION_SETTINGS, includeDiagrams: false, generateMissingDependencies: true,
    } });
    expect(markdown).not.toContain('```mermaid');
    expect(markdown).toContain('## Missing Dependencies');
    expect(splitMarkdownByCategory(markdown).some((category) => category.markdown.includes('new_case.new_customer'))).toBe(true);
  });

  it('does not expose current environment variable values or mistake defaults for current values', async () => {
    const solution = await sampleSolution();
    solution.environmentVariables = [
      { name: 'new_api', schemaName: 'new_api', type: 'String', currentValue: 'DO_NOT_EXPORT', hasCurrentValue: true },
      { name: 'new_default', schemaName: 'new_default', type: 'String', defaultValue: 'default-only', hasCurrentValue: false },
      { name: 'new_unknown', schemaName: 'new_unknown', type: 'String' },
    ];
    const markdown = generateMarkdown(solution);
    expect(markdown).not.toContain('DO_NOT_EXPORT');
    expect(markdown).toContain('Has Current Value');
    expect(markdown).toContain('| No | default-only |');
    expect(markdown).toContain('| Unknown |');
  });
});