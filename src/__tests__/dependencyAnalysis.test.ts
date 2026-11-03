import JSZip from 'jszip';
import { parseSolutionZip } from '../parser/solutionParser';
import { generateDependencyReportMarkdown, generateInstallationMarkdown } from '../generator/localDocumentation';
import { analyzeDependencies } from '../generator/dependencyAnalysis';
import type { ParsedSolution } from '../types/solution';

const missing = (required: string, dependent = '<Dependent type="2" schemaName="new_x" parentSchemaName="new_case" />') =>
  `<MissingDependency>${required}${dependent}</MissingDependency>`;

async function zipSolution(name: string, options: { version?: string; prefix?: string; manifest?: string; customizations?: string; files?: Record<string, string> } = {}) {
  const zip = new JSZip();
  zip.file('solution.xml', `<ImportExportXml><SolutionManifest><UniqueName>${name}</UniqueName><Version>${options.version ?? '1.0.0.0'}</Version>`
    + `<Publisher><UniqueName>pub</UniqueName><CustomizationPrefix>${options.prefix ?? 'new'}</CustomizationPrefix></Publisher>${options.manifest ?? ''}</SolutionManifest></ImportExportXml>`);
  zip.file('customizations.xml', options.customizations ?? '<ImportExportXml/>');
  Object.entries(options.files ?? {}).forEach(([path, content]) => zip.file(path, content));
  return parseSolutionZip(await zip.generateAsync({ type: 'blob' }));
}

const entityXml = (logicalName: string, attributes: string[] = [], forms = '') => `<Entity><Name>${logicalName}</Name><EntityInfo><entity Name="${logicalName}"><attributes>${
  attributes.map((attribute) => `<attribute PhysicalName="${attribute}"><Type>nvarchar</Type><Name>${attribute}</Name><LogicalName>${attribute}</LogicalName></attribute>`).join('')
}</attributes></entity></EntityInfo>${forms}</Entity>`;

const statuses = (solutions: ParsedSolution[]) => analyzeDependencies(solutions).resolutions
  .map((resolution) => `${resolution.consumer}:${resolution.dependency.requiredName}:${resolution.status}`).sort();

describe('dependency analysis', () => {
  it('treats components defined in the same ZIP as satisfied, including ID-only forms and columns under a table shell', async () => {
    const formId = '11111111-2222-3333-4444-555555555555';
    const solution = await zipSolution('self', {
      manifest: `<RootComponents><RootComponent type="1" schemaName="new_case" behavior="2" /></RootComponents><MissingDependencies>${
        missing(`<Required type="60" displayName="Case Main" parentDisplayName="Case" id="{${formId.toUpperCase()}}" solution="Active" />`)
      }${missing('<Required type="2" schemaName="new_Detail" parentSchemaName="new_case" solution="Active" />')}</MissingDependencies>`,
      customizations: `<ImportExportXml><Entities>${entityXml('new_case', ['new_detail'], `<FormXml><forms type="main"><systemform><formid>{${formId}}</formid></systemform></forms></FormXml>`)}</Entities></ImportExportXml>`,
    });
    expect(solution.ownedComponentIds).toContain(formId);
    expect(statuses([solution])).toEqual(expect.arrayContaining([`self:Case Main:included`, 'self:new_Detail:included']));
    const markdown = generateInstallationMarkdown([solution]);
    expect(markdown).toContain('| 1 | self |');
    expect(markdown).toContain('No missing dependencies.');
    expect(markdown).toContain('Dependencies Included in the Same Solution');
  });

  it('resolves components defined in other selected ZIPs, including id.* keyed environment variables, and orders providers first', async () => {
    const provider = await zipSolution('z_provider', {
      customizations: `<ImportExportXml><Entities>${entityXml('new_case', ['new_detail'])}</Entities></ImportExportXml>`,
      files: { 'environmentvariabledefinitions/new_Url/environmentvariabledefinition.xml': '<environmentvariabledefinition><schemaname>new_Url</schemaname><type>String</type></environmentvariabledefinition>' },
    });
    const consumer = await zipSolution('a_consumer', { manifest: `<MissingDependencies>${
      missing('<Required type="1" schemaName="new_case" solution="Active" />')
    }${missing('<Required type="2" schemaName="new_detail" parentSchemaName="new_case" solution="Active" />')
    }${missing('<Required type="EnvironmentVariableDefinition" displayName="Url" id.schemaname="new_Url" solution="Active" />')}</MissingDependencies>` });
    expect(consumer.missingDependencies?.find((dependency) => dependency.requiredType === 'Environment Variable')?.requiredName).toBe('new_Url');
    expect(statuses([consumer, provider])).toEqual(['a_consumer:new_Url:selected', 'a_consumer:new_case:selected', 'a_consumer:new_detail:selected']);
    const markdown = generateInstallationMarkdown([consumer, provider]);
    expect(markdown).toContain('| 1 | z_provider |');
    expect(markdown).toContain('| 2 | a_consumer | 1.0.0.0 | z_provider | Ready |');
    expect(markdown).toContain('No missing dependencies.');
  });

  it('accepts a selected owner declared by the manifest even when the component type is not itemised', async () => {
    const provider = await zipSolution('controls', { version: '2.0.0.0' });
    const consumer = await zipSolution('app', { manifest: `<MissingDependencies>${
      missing('<Required type="66" schemaName="new_Namespace.Control" solution="controls (2.0.0.0)" />')}</MissingDependencies>` });
    expect(statuses([consumer, provider])).toEqual(['app:new_Namespace.Control:selected']);
    expect(generateInstallationMarkdown([consumer, provider])).toContain('| 2 | app |');
  });

  it('flags itemised components absent from the declared selected owner for verification without inventing a gap', async () => {
    const provider = await zipSolution('z_core');
    const consumer = await zipSolution('a_app', { manifest: `<MissingDependencies>${
      missing('<Required type="1" schemaName="new_case" solution="z_core (1.0.0.0)" />')}</MissingDependencies>` });
    const markdown = generateInstallationMarkdown([consumer, provider]);
    expect(markdown).toContain('| 2 | a_app |');
    expect(markdown).toContain('was not found in that ZIP. Verify the selected version contains it.');
  });

  it('matches connection references by logical or display name without matching other component types', async () => {
    const provider = await zipSolution('z_conn');
    provider.connectionReferences = [{ name: 'new_sharedoffice365', displayName: 'Office 365 Outlook', connectorId: 'shared_office365' }];
    provider.webResources = [{ name: 'new_case', schemaName: 'new_case', resourceType: provider.webResources[0]?.resourceType ?? ('JavaScript' as never) }];
    const consumer = await zipSolution('a_flow', { manifest: `<MissingDependencies>${
      missing('<Required type="connectionreference" displayName="Office 365 Outlook" solution="Active" />')}${
      missing('<Required type="1" schemaName="new_case" solution="Active" />')}</MissingDependencies>` });
    expect(statuses([consumer, provider])).toEqual(['a_flow:Office 365 Outlook:selected', 'a_flow:new_case:missing']);
  });

  it('never flags standard tables, columns, platform owners or Microsoft-supplied components as missing', async () => {
    const solution = await zipSolution('custom', { manifest: `<MissingDependencies>${[
      '<Required type="1" schemaName="account" />',
      '<Required type="1" schemaName="contact" displayName="Contact" solution="Active" />',
      '<Required type="2" schemaName="address1_city" parentSchemaName="account" solution="Active" />',
      '<Required type="2" schemaName="emailaddress1" parentSchemaName="contact" solution="System" />',
      '<Required type="1" schemaName="systemuser" solution="System" />',
      '<Required type="1" schemaName="msdyn_aimodel" solution="msdyn_AISolution (202509.2.2.1)" />',
      '<Required type="1" schemaName="bot" solution="PowerVirtualAgents (2026.1.2.18829146)" />',
      '<Required type="1" schemaName="environmentvariabledefinition" solution="EnvironmentVariables (1.0.2406.4)"><package appName="EnvironmentVariables" version="1.0.2406.4">EnvironmentVariablesSolutionAnchor (1.0.2406.4)</package></Required>',
      '<Required type="60" displayName="Information" parentDisplayName="Account" id="{aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}" solution="Active" />',
    ].map((required) => missing(required)).join('')}</MissingDependencies>` });
    const results = analyzeDependencies([solution]).resolutions;
    expect(results.filter((result) => result.status === 'missing')).toEqual([]);
    expect(results.filter((result) => result.status === 'platform').map((result) => result.dependency.requiredName).sort())
      .toEqual(['account', 'address1_city', 'contact', 'emailaddress1', 'systemuser']);
    expect(results.filter((result) => result.status === 'environment').map((result) => result.dependency.requiredName).sort())
      .toEqual(['Information', 'bot', 'environmentvariabledefinition', 'msdyn_aimodel']);
    expect(results.find((result) => result.dependency.requiredName === 'environmentvariabledefinition')?.owner).toBe('EnvironmentVariables 1.0.2406.4');
    const markdown = generateInstallationMarkdown([solution]);
    expect(markdown).toContain('| 1 | custom |');
    expect(markdown).toContain('Target Environment Prerequisites');
    expect(markdown).toContain('Standard Platform Components');
    expect(markdown).not.toContain('## Blocked Solutions');
  });

  it('reports genuine gaps with the dependent solution, component, expected owner and action', async () => {
    const solution = await zipSolution('consumer', { manifest: `<MissingDependencies>${[
      '<Required type="1" schemaName="new_absent" solution="Active" />',
      '<Required type="2" schemaName="new_score" parentSchemaName="account" solution="Active" />',
      '<Required type="61" schemaName="contoso_/script.js" solution="contoso_core (3.1.0.0)" />',
    ].map((required) => missing(required)).join('')}</MissingDependencies>` });
    expect(statuses([solution])).toEqual(['consumer:contoso_/script.js:missing', 'consumer:new_absent:missing', 'consumer:new_score:missing']);
    const markdown = generateInstallationMarkdown([solution]);
    expect(markdown).toContain('| consumer | Column: new_case.new_x | Table: new_absent | Active |');
    expect(markdown).toContain('contoso_core (3.1.0.0)');
    expect(markdown).toContain('Owned by managed solution contoso_core (3.1.0.0), which is not among the selected ZIPs');
    expect(markdown).toContain('No solution can be installed');
    const provider = await zipSolution('extension', { customizations: `<ImportExportXml><Entities>${entityXml('account', ['new_score'])}</Entities></ImportExportXml>` });
    expect(statuses([solution, provider])).toContain('consumer:new_score:selected');
  });

  it('prefers a full table owner over a shell in the declaring solution', async () => {
    const owner = await zipSolution('z_owner', { customizations: `<ImportExportXml><Entities>${entityXml('new_case')}</Entities></ImportExportXml>` });
    const shell = await zipSolution('a_shell', {
      manifest: `<RootComponents><RootComponent type="1" schemaName="new_case" behavior="2" /></RootComponents><MissingDependencies>${missing('<Required type="1" schemaName="new_case" solution="Active" />')}</MissingDependencies>`,
      customizations: `<ImportExportXml><Entities>${entityXml('new_case')}</Entities></ImportExportXml>`,
    });
    expect(statuses([shell, owner])).toEqual(['a_shell:new_case:selected']);
    expect(statuses([shell])).toEqual(['a_shell:new_case:included']);
  });

  it('matches ID-less forms by display name and parent table display name', async () => {
    const provider = await zipSolution('z_forms');
    provider.entities = [{ name: 'new_case', logicalName: 'new_case', displayName: 'Case', isCustom: true, attributes: [], relationships: [] }];
    provider.forms = [{ name: 'Case Main', displayName: 'Case Main', entityLogicalName: 'new_case', formType: 'Main', fields: [] }];
    const consumer = await zipSolution('a_app', { manifest: `<MissingDependencies>${missing('<Required type="60" displayName="Case Main" parentDisplayName="Case" solution="Active" />')}</MissingDependencies>` });
    expect(statuses([consumer, provider])).toEqual(['a_app:Case Main:selected']);
  });

  it('does not create solution-level failures from component owners and deduplicates rows', async () => {
    const consumer = await zipSolution('consumer', { manifest: `<MissingDependencies>${
      missing('<Required type="1" schemaName="bot" solution="PowerVirtualAgents (2026.1)" />')}${
      missing('<Required type="1" schemaName="bot" solution="PowerVirtualAgents (2026.1)" />')}</MissingDependencies>` });
    const markdown = generateInstallationMarkdown([consumer, consumer]);
    expect(markdown).not.toContain('Not among the selected ZIPs');
    expect(markdown.match(/\| \d+ \| consumer \|/g)).toHaveLength(1);
    expect(markdown.match(/PowerVirtualAgents \(2026\.1\) \| Table: bot/g)).toHaveLength(1);
  });

  it('blocks insufficient explicit minimum versions, transitive consumers, cycles and duplicate versions', async () => {
    const base = await zipSolution('base', { version: '1.0.0.1' });
    const middle = await zipSolution('middle', { manifest: '<Dependencies><Dependency solution="base" version="1.0.0.2" /></Dependencies>' });
    const final = await zipSolution('final', { manifest: '<Dependencies><Dependency solution="middle" /></Dependencies>' });
    const versions = generateInstallationMarkdown([final, middle, base]);
    expect(versions).toContain('Requires base 1.0.0.2 or later; the selected ZIP is 1.0.0.1.');
    expect(versions).toContain('Waits for blocked prerequisite middle.');
    expect(versions).toContain('| 1 | base |');
    expect(versions).not.toContain('| final | 1.0.0.0');
    const first = await zipSolution('first', { manifest: '<Dependencies><Dependency solution="second" /></Dependencies>' });
    const second = await zipSolution('second', { manifest: '<Dependencies><Dependency solution="first" /></Dependencies>' });
    const cycle = generateInstallationMarkdown([first, second]);
    expect(cycle).toContain('Circular dependency');
    expect(cycle).not.toMatch(/\| 1 \|/);
    const newer = await zipSolution('base', { version: '2.0.0.0' });
    expect(generateInstallationMarkdown([base, newer])).toContain('Several ZIPs of base are selected');
    expect(generateInstallationMarkdown([{ ...base, missingDependencies: undefined }])).toContain('Import dependencies have not been collected');
  });

  it('orders email templates after selected table and column providers for every input permutation', async () => {
    const table = await zipSolution('z_tables', { customizations: `<ImportExportXml><Entities>${entityXml('new_case')}</Entities></ImportExportXml>` });
    const column = await zipSolution('y_columns', {
      manifest: '<Dependencies><Dependency solution="z_tables" /></Dependencies>',
      customizations: `<ImportExportXml><Entities>${entityXml('new_case', ['new_detail'])}</Entities></ImportExportXml>`,
    });
    const email = await zipSolution('a_email', { customizations: '<ImportExportXml><templates><template><title>Case notification</title><templatetypecode>new_case</templatetypecode><body>{!new_case:new_detail;} {!account:name;}</body></template></templates></ImportExportXml>' });
    for (const inputs of [[email, column, table], [email, table, column], [column, email, table], [column, table, email], [table, email, column], [table, column, email]]) {
      const markdown = generateInstallationMarkdown(inputs);
      expect(markdown).toContain('| 1 | z_tables |');
      expect(markdown).toContain('| 2 | y_columns |');
      expect(markdown).toContain('| 3 | a_email |');
      expect(markdown).toContain('| a_email | Case notification | Table: account | Standard Dataverse component present in every environment. |');
    }
  });

  it('marks inferred references that conflict with declared order for review instead of creating cycles', async () => {
    const first = await zipSolution('first', {
      manifest: '<Dependencies><Dependency solution="second" /></Dependencies>',
      customizations: `<ImportExportXml><Entities>${entityXml('new_case')}</Entities></ImportExportXml>`,
    });
    const second = await zipSolution('second', { customizations: '<ImportExportXml><templates><template><title>Email</title><templatetypecode>new_case</templatetypecode></template></templates></ImportExportXml>' });
    const markdown = generateInstallationMarkdown([first, second]);
    expect(markdown).toContain('conflicts with the declared order');
    expect(markdown).not.toContain('Circular dependency');
    expect(markdown).toContain('## Blocked Solutions');
  });

  it('classifies per-solution and consolidated dependency reports against the selected context', async () => {
    const provider = await zipSolution('provider', { customizations: `<ImportExportXml><Entities>${entityXml('new_case')}</Entities></ImportExportXml>` });
    const consumer = await zipSolution('consumer', { manifest: `<MissingDependencies>${
      missing('<Required type="1" schemaName="new_case" solution="Active" />')}${missing('<Required type="1" schemaName="account" />')}</MissingDependencies>` });
    const alone = generateDependencyReportMarkdown(consumer);
    expect(alone).toContain('2 declared prerequisites: 1 missing');
    const withContext = generateDependencyReportMarkdown(consumer, [consumer, provider]);
    expect(withContext).toContain('0 missing');
    expect(withContext).toContain('## Provided by Selected Solutions');
    expect(withContext).toContain('## Standard Platform Components');
    expect(withContext).toContain('No missing dependencies.');
  });
});
