import { describe, expect, it } from 'vitest';
import { generateMetadataGridMarkdown, buildMetadataGridRows } from '../generator/markdownGenerator';
import { AttributeType, type ParsedSolution } from '../types/solution';

const sampleSolution: ParsedSolution = {
  metadata: {
    uniqueName: 'phase1_grid',
    displayName: 'Phase1 Grid',
    version: '1.0.5',
    publisherName: 'Contoso',
    isManaged: false,
    dependencies: [],
    componentInventory: [],
  },
  entities: [
    {
      name: 'account',
      displayName: 'Account',
      logicalName: 'account',
      isCustom: false,
      attributes: [
        {
          name: 'name',
          displayName: 'Account Name',
          type: AttributeType.String,
          required: true,
          isCustom: false,
        },
      ],
      relationships: [
        {
          name: 'account_primary_contact',
          type: 'OneToMany',
          referencedEntity: 'contact',
          referencingEntity: 'account',
          referencingAttribute: 'primarycontactid',
        },
      ],
    },
  ],
  optionSets: [],
  forms: [],
  views: [],
  processes: [],
  apps: [],
  agents: [],
  aiModels: [],
  desktopFlows: [],
  dataflows: [],
  customApis: [],
  offlineProfiles: [],
  webResources: [],
  securityRoles: [],
  fieldSecurityProfiles: [],
  connectionReferences: [],
  environmentVariables: [],
  emailTemplates: [],
  reports: [],
  dashboards: [],
  pluginAssemblies: [],
  warnings: [],
};

describe('metadata grid export', () => {
  it('builds entity, attribute, and relationship rows', () => {
    const rows = buildMetadataGridRows(sampleSolution);

    expect(rows.find((row) => row.value.startsWith('Entity:'))).toBeTruthy();
    expect(rows.find((row) => row.attribute === 'name')).toBeTruthy();
    expect(rows.find((row) => row.relationship === 'account_primary_contact')).toBeTruthy();
  });

  it('renders markdown for metadata grid output', () => {
    const markdown = generateMetadataGridMarkdown(sampleSolution);

    expect(markdown).toContain('## Solution Metadata Grid');
    expect(markdown).toContain('| Entity | Attribute | Relationship | Value |');
    expect(markdown).toContain('| account | name | – | String (required) |');
  });
});
