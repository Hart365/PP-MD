import {
  DEFAULT_DOCUMENTATION_SETTINGS,
  generateConsolidatedMarkdown,
  generateMarkdown,
  splitMarkdownForDiagramCompanion,
} from '../generator/markdownGenerator';
import {
  AppType,
  AttributeType,
  ProcessCategory,
  WebResourceType,
  type ParsedSolution,
} from '../types/solution';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

function normalizeForComparison(value: string): string {
  return value.replace(/\r\n/g, '\n').trimEnd();
}

const sampleSolution: ParsedSolution = {
  metadata: {
    uniqueName: 'contoso_core',
    displayName: 'Contoso Core',
    version: '1.0.4',
    publisherName: 'Contoso',
    isManaged: false,
    dependencies: [],
    componentInventory: [],
  },
  entities: [
    {
      name: 'new_project',
      displayName: 'Project',
      logicalName: 'new_project',
      isCustom: true,
      attributes: [
        {
          name: 'new_name',
          displayName: 'Name',
          type: AttributeType.String,
          required: true,
          isCustom: true,
        },
      ],
      relationships: [],
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

describe('documentation option output matrix', () => {
  const solution: ParsedSolution = {
    ...sampleSolution,
    entities: [{ ...sampleSolution.entities[0], attributes: [
      { name: 'new_onform', displayName: 'On Form', type: AttributeType.String, required: true, isCustom: true, isManaged: false, description: 'Column description', isAuditEnabled: true, isSecured: true, isValidForAdvancedFind: true },
      { name: 'new_choice', type: AttributeType.OptionSet, required: false, isCustom: true, isManaged: true },
      { name: 'system_default', type: AttributeType.String, required: false, isCustom: false, isManaged: true },
      { name: 'new_virtual', type: AttributeType.Virtual, required: false, isCustom: true, isManaged: false },
    ] }],
    forms: [{ name: 'Main', entityLogicalName: 'new_project', formType: 'Main', fields: [{ attributeName: 'new_onform' }] }],
    processes: [{ name: 'AUTOMATION_SENTINEL', uniqueName: 'audit_flow', category: ProcessCategory.PowerAutomateFlow, primaryEntity: 'new_project', steps: [] }],
    apps: [{ name: 'APP_SENTINEL', uniqueName: 'audit_app', appType: AppType.ModelDriven, entities: ['new_project'] }],
    securityRoles: [{ name: 'ROLE_SENTINEL', privileges: [{ privilegeName: 'prvReadnew_project', depth: 4 }] }],
    connectionReferences: [{ name: 'CONNECTION_SENTINEL', connectorId: 'shared_dataverse' }],
    pluginAssemblies: [{ name: 'PLUGIN_SENTINEL', assemblyName: 'PLUGIN_SENTINEL', isolationMode: 2, steps: [] }],
    reports: [{ name: 'REPORT_SENTINEL', relatedEntities: ['new_project'] }],
    webResources: [{ name: 'WEB_SENTINEL', schemaName: 'audit_script', resourceType: WebResourceType.JavaScript }],
    desktopFlows: [{ name: 'DESKTOP_SENTINEL', sourcePath: 'desktop.json' }],
    dataflows: [{ name: 'DATAFLOW_SENTINEL', sourcePath: 'dataflow.json' }],
    customApis: [{ name: 'API_SENTINEL', sourcePath: 'api.xml' }],
    offlineProfiles: [{ name: 'OFFLINE_SENTINEL', sourcePath: 'offline.xml' }],
    agents: [{ name: 'AGENT_SENTINEL', sourcePath: 'agent.yaml' }],
    aiModels: [{ name: 'AI_SENTINEL', sourcePath: 'ai.json' }],
  };

  const scopeCases = [
    ['flows', ['Processes & Automation'], ['AUTOMATION_SENTINEL']],
    ['apps', ['Power Apps'], ['APP_SENTINEL']],
    ['security', ['Security Roles'], ['ROLE_SENTINEL']],
    ['integration', ['Connection References', 'Connections'], ['CONNECTION_SENTINEL']],
    ['plugins', ['Plugin Assemblies & Steps'], ['PLUGIN_SENTINEL']],
    ['reports', ['Reports & Dashboards'], ['REPORT_SENTINEL']],
    ['webResources', ['Web Resources'], ['WEB_SENTINEL']],
    ['desktopFlows', ['Desktop Flows', 'Dataflows'], ['DESKTOP_SENTINEL', 'DATAFLOW_SENTINEL']],
    ['customApis', ['Custom APIs', 'Offline Profiles'], ['API_SENTINEL', 'OFFLINE_SENTINEL']],
    ['copilotAgents', ['Copilot Studio Agents', 'AI Models'], ['AGENT_SENTINEL', 'AI_SENTINEL']],
  ] as const;
  for (const [key, headings, markers] of scopeCases) {
    it(`honours the ${key} scope checkbox throughout output`, () => {
      const enabled = generateMarkdown(solution);
      const disabled = generateMarkdown(solution, { documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS, scope: { ...DEFAULT_DOCUMENTATION_SETTINGS.scope, [key]: false },
      } });
      headings.forEach((title) => {
        expect(enabled).toContain(`## ${title}`);
        expect(disabled).not.toContain(`## ${title}`);
      });
      markers.forEach((marker) => {
        expect(enabled).toContain(marker);
        expect(disabled).not.toContain(marker);
      });
    });
  }

  const columnCases = [
    ['includeTypeColumn', 'Type'], ['includeRequiredLevelInfo', 'Required'],
    ['includeCustomColumn', 'Custom'], ['includeAuditInfo', 'Audited'],
    ['includeFieldSecurityFlags', 'Field Security'], ['includeValidForAdvancedFindInfo', 'Advanced Find'],
    ['includeMetadataDiagnosticInfo', 'Metadata Source'], ['includeNotesColumn', 'Notes'],
    ['includeDescriptionColumn', 'Description'],
  ] as const;
  for (const [key, title] of columnCases) {
    it(`honours the ${title} column checkbox`, () => {
      for (const enabled of [true, false]) {
        const markdown = generateMarkdown(solution, { documentationSettings: {
          ...DEFAULT_DOCUMENTATION_SETTINGS, metadata: { ...DEFAULT_DOCUMENTATION_SETTINGS.metadata, [key]: enabled },
        } });
        const header = markdown.split('#### Columns')[1]?.split('\n').find((line) => line.startsWith('| Display Name')) ?? '';
        expect(header.split('|').map((cell) => cell.trim()).includes(title)).toBe(enabled);
      }
    });
  }

  const selectionCases = [
    ['all', ['new_onform', 'new_choice', 'system_default', 'new_virtual']],
    ['custom-only', ['new_onform', 'new_choice', 'new_virtual']],
    ['attributes-on-form', ['new_onform']],
    ['attributes-not-on-form', ['new_choice', 'system_default', 'new_virtual']],
    ['option-set-focused', ['new_choice']],
    ['unmanaged-only', ['new_onform', 'new_virtual']],
  ] as const;
  for (const [mode, expected] of selectionCases) {
    it(`outputs the expected columns in ${mode} mode`, () => {
      const markdown = generateMarkdown(solution, { documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS, metadata: { ...DEFAULT_DOCUMENTATION_SETTINGS.metadata, attributeSelectionMode: mode },
      } });
      const columns = markdown.split('#### Columns')[1]?.split('####')[0] ?? '';
      const names = solution.entities[0].attributes.filter((attribute) => columns.includes(`\`${attribute.name}\``)).map((attribute) => attribute.name);
      expect(names).toEqual(expected);
    });
  }

  it('honours default-column and virtual-column inclusion independently', () => {
    for (const includeDefaultColumns of [true, false]) {
      for (const excludeVirtualAttributes of [true, false]) {
        const markdown = generateMarkdown(solution, { documentationSettings: {
          ...DEFAULT_DOCUMENTATION_SETTINGS, metadata: { ...DEFAULT_DOCUMENTATION_SETTINGS.metadata, includeDefaultColumns, excludeVirtualAttributes },
        } });
        const columns = markdown.split('#### Columns')[1]?.split('####')[0] ?? '';
        expect(columns.includes('`system_default`')).toBe(includeDefaultColumns);
        expect(columns.includes('`new_virtual`')).toBe(!excludeVirtualAttributes);
      }
    }
  });

  it('does not leave diagram headings or links when diagrams are excluded', () => {
    const markdown = generateMarkdown(solution, { documentationSettings: { ...DEFAULT_DOCUMENTATION_SETTINGS, includeDiagrams: false } });
    expect(markdown).not.toContain('```mermaid');
    expect(markdown).not.toContain('## Entity Relationship Diagram');
    expect(markdown).not.toContain('](#entity-relationship-diagram)');
    expect(markdown).not.toContain('## Solution Component Relationship Graph');
  });

  it('uses the selected ERD detail mode without changing the table documentation', () => {
    const compact = generateMarkdown(solution, { erdMode: 'compact' });
    const detailed = generateMarkdown(solution, { erdMode: 'detailed-relationships' });
    expect(compact).toContain('#### Columns');
    expect(detailed).toContain('#### Columns');
    const erd = (markdown: string) => markdown.split('## Entity Relationship Diagram')[1]?.split('## Tables & Columns')[0] ?? '';
    expect(erd(compact)).not.toContain('logical_name PK');
    expect(erd(detailed)).toContain('logical_name PK');
  });

  it('allows selecting neutral again after a custom palette', () => {
    const markdown = generateMarkdown(solution, { documentationSettings: {
      ...DEFAULT_DOCUMENTATION_SETTINGS, mermaidColourTheme: 'neutral',
      mermaidPalette: { sourceOne: '#123456', sourceTwo: '#654321', sourceThree: '#abcdef', sourceFour: '#fedcba' },
    } });
    markdown.split('\n').filter((line) => line.startsWith('%%{ init:')).forEach((line) => expect(line).toContain("'theme': 'neutral'"));
  });

  it('uses all four custom diagram colour controls', () => {
    const markdown = generateMarkdown({
      ...solution,
      views: [{ name: 'Palette View', entityLogicalName: 'new_project', viewType: 'Public', columns: [] }],
      dashboards: [{ name: 'Palette Dashboard', entityLogicalName: 'new_project', dashboardType: 'System', components: [] }],
    }, { documentationSettings: {
      ...DEFAULT_DOCUMENTATION_SETTINGS, mermaidColourTheme: 'base',
      mermaidPalette: { sourceOne: '#123456', sourceTwo: '#654321', sourceThree: '#abcdef', sourceFour: '#fedcba' },
    } });
    for (const colour of ['#123456', '#654321', '#abcdef', '#fedcba']) expect(markdown).toContain(`stroke:${colour}`);
  });

  for (const theme of ['neutral', 'default', 'dark', 'forest', 'base'] as const) {
    it(`uses the selected ${theme} theme in exported diagrams`, () => {
      const markdown = generateMarkdown(solution, { documentationSettings: { ...DEFAULT_DOCUMENTATION_SETTINGS, mermaidColourTheme: theme } });
      const initializers = markdown.split('\n').filter((line) => line.startsWith('%%{ init:'));
      expect(initializers.length).toBeGreaterThan(0);
      initializers.forEach((line) => expect(line).toContain(`'theme': '${theme}'`));
    });
  }

  it('honours disabled scope in consolidated summaries', () => {
    const summary = generateConsolidatedMarkdown([solution, { ...solution, metadata: { ...solution.metadata, uniqueName: 'other' } }], {
      documentationSettings: { ...DEFAULT_DOCUMENTATION_SETTINGS, scope: { ...DEFAULT_DOCUMENTATION_SETTINGS.scope, apps: false, flows: false } },
    });
    expect(summary).not.toContain('APP_SENTINEL');
    expect(summary).not.toContain('AUTOMATION_SENTINEL');
    expect(summary).not.toContain('## Apps Across Solutions');
    expect(summary).not.toContain('## Processes Across Solutions');
  });
});

describe('generateMarkdown', () => {
  it('matches the golden markdown baseline', () => {
    const markdown = normalizeForComparison(
      generateMarkdown(sampleSolution)
        .replace(/^> Generated on: .*$/m, '> Generated on: <normalized>'),
    );
    const expected = readFileSync(
      resolve(__dirname, 'fixtures', 'markdownGenerator.sample.golden.md'),
      'utf8',
    );

    expect(markdown).toBe(normalizeForComparison(expected));
  });

  it('omits detailed-only sections when detail level is summary', () => {
    const markdown = generateMarkdown(sampleSolution, {
      documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS,
        detailLevel: 'summary',
      },
    });

    expect(markdown).not.toContain('## Entity Relationship Diagram');
    expect(markdown).not.toContain('## Tables & Columns');
    expect(markdown).not.toContain('## Global Option Sets');
    expect(markdown).not.toContain('## Forms & Views');
    expect(markdown).not.toContain('## Web Resources');
  });

  it('respects scope toggles for apps and flows', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        processes: [
          {
            name: 'Sync Contacts',
            uniqueName: 'contoso_sync_contacts',
            category: ProcessCategory.PowerAutomateFlow,
            steps: [],
          },
        ],
        apps: [
          {
            name: 'Sales Hub',
            uniqueName: 'contoso_saleshub',
            appType: AppType.ModelDriven,
          },
        ],
      },
      {
        documentationSettings: {
          ...DEFAULT_DOCUMENTATION_SETTINGS,
          scope: {
            ...DEFAULT_DOCUMENTATION_SETTINGS.scope,
            flows: false,
            apps: false,
          },
        },
      },
    );

    expect(markdown).not.toContain('## Processes & Automation');
    expect(markdown).not.toContain('## Power Apps');
  });

  it('splits process summaries by automation type with type-specific columns', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        processes: [
          {
            name: 'account-sync',
            displayName: 'Account Sync',
            uniqueName: 'account-sync',
            category: ProcessCategory.PowerAutomateFlow,
            primaryEntity: 'account',
            relatedEntities: ['account'],
            flowConnectors: ['Microsoft Dataverse'],
            flowConnectionReferences: ['cr_dataverse'],
            flowEnvironmentVariables: ['contoso_FeatureFlag'],
            flowTrigger: 'When a row is added',
            isActivated: true,
            steps: [],
          },
          {
            name: 'legacy-workflow',
            displayName: 'Legacy Workflow',
            uniqueName: 'legacy-workflow',
            category: ProcessCategory.Workflow,
            primaryEntity: 'account',
            relatedEntities: ['account'],
            triggerType: 'OnDemand',
            runAs: 'Owner',
            scope: 'Organization',
            isActivated: false,
            steps: [],
          },
        ],
      },
    );

    expect(markdown).toContain('### Power Automate Flow Summary');
    expect(markdown).toContain('### Workflow Summary');
    expect(markdown).toContain('| Name | Primary Table | Referenced Tables | Connectors | Connection References | Environment Variables | Trigger |');
    expect(markdown).toContain('| Name | Primary Table | Referenced Tables | Trigger | Run As | Scope |');
    expect(markdown).not.toContain('| Name | Category | Primary Table | Referenced Tables | Connectors | Connection References | Environment Variables | Trigger | Status |');
  });

  it('applies metadata filters for virtual exclusion and option-set-focused mode', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        entities: [
          {
            ...sampleSolution.entities[0],
            attributes: [
              {
                name: 'new_name',
                displayName: 'Name',
                type: AttributeType.String,
                required: true,
                isCustom: true,
              },
              {
                name: 'new_status',
                displayName: 'Status',
                type: AttributeType.OptionSet,
                required: false,
                isCustom: true,
              },
              {
                name: 'new_virtualscore',
                displayName: 'Virtual Score',
                type: AttributeType.Virtual,
                required: false,
                isCustom: false,
              },
            ],
          },
        ],
      },
      {
        documentationSettings: {
          ...DEFAULT_DOCUMENTATION_SETTINGS,
          metadata: {
            ...DEFAULT_DOCUMENTATION_SETTINGS.metadata,
            attributeSelectionMode: 'option-set-focused',
            excludeVirtualAttributes: true,
          },
        },
      },
    );

    expect(markdown).toContain('`new_status`');
    expect(markdown).not.toContain('`new_name`');
    expect(markdown).not.toContain('`new_virtualscore`');
  });

  it('filters security role permissions to solution tables and custom tables', () => {
    const solutionWithRoles: ParsedSolution = {
      ...sampleSolution,
      entities: [
        ...sampleSolution.entities,
        {
          name: 'account',
          logicalName: 'account',
          displayName: 'Account',
          isCustom: false,
          attributes: [],
          relationships: [],
        },
      ],
      securityRoles: [
        {
          name: 'sales_role',
          displayName: 'Sales Role',
          privileges: [
            { privilegeName: 'prvReadnew_project', depth: 4 },
            { privilegeName: 'prvWritenew_project', depth: 2 },
            { privilegeName: 'prvReadaccount', depth: 1 },
            { privilegeName: 'prvReadexternaltable', depth: 4 },
          ],
        },
      ],
    };

    const solutionTablesOnlyMarkdown = generateMarkdown(solutionWithRoles, {
      documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS,
        securityRoleFilters: {
          onlyTablesInCurrentSolution: true,
          onlyCustomTables: false,
        },
      },
    });

    const solutionTablesOnlySecuritySection = solutionTablesOnlyMarkdown.split('## Security Roles')[1] ?? '';

    expect(solutionTablesOnlyMarkdown).toContain('| Sales Role | 3 |');
    expect(solutionTablesOnlySecuritySection).toContain('`new_project`');
    expect(solutionTablesOnlySecuritySection).toContain('`account`');
    expect(solutionTablesOnlySecuritySection).not.toContain('`externaltable`');

    const customTablesOnlyMarkdown = generateMarkdown(solutionWithRoles, {
      documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS,
        securityRoleFilters: {
          onlyTablesInCurrentSolution: false,
          onlyCustomTables: true,
        },
      },
    });

    const customTablesOnlySecuritySection = customTablesOnlyMarkdown.split('## Security Roles')[1] ?? '';

    expect(customTablesOnlyMarkdown).toContain('| Sales Role | 2 |');
    expect(customTablesOnlySecuritySection).toContain('`new_project`');
    expect(customTablesOnlySecuritySection).not.toContain('`account`');
    expect(customTablesOnlySecuritySection).not.toContain('`externaltable`');

    const solutionAndCustomTablesMarkdown = generateMarkdown(solutionWithRoles, {
      documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS,
        securityRoleFilters: {
          onlyTablesInCurrentSolution: true,
          onlyCustomTables: true,
        },
      },
    });

    const solutionAndCustomTablesSecuritySection = solutionAndCustomTablesMarkdown.split('## Security Roles')[1] ?? '';

    expect(solutionAndCustomTablesMarkdown).toContain('| Sales Role | 2 |');
    expect(solutionAndCustomTablesSecuritySection).toContain('`new_project`');
    expect(solutionAndCustomTablesSecuritySection).not.toContain('`account`');
    expect(solutionAndCustomTablesSecuritySection).not.toContain('`externaltable`');
  });

  it('matches security role privilege tables by entity logical name and entity set name', () => {
    const solutionWithEntitySetRole: ParsedSolution = {
      ...sampleSolution,
      entities: [
        {
          ...sampleSolution.entities[0],
          logicalName: 'new_project',
          entitySetName: 'new_projects',
          isCustom: true,
        },
      ],
      securityRoles: [
        {
          name: 'project_reader',
          displayName: 'Project Reader',
          privileges: [
            { privilegeName: 'prvReadnew_projects', depth: 4 },
          ],
        },
      ],
    };

    const markdown = generateMarkdown(solutionWithEntitySetRole, {
      documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS,
        securityRoleFilters: {
          onlyTablesInCurrentSolution: true,
          onlyCustomTables: true,
        },
      },
    });

    const securitySection = markdown.split('## Security Roles')[1] ?? '';

    expect(markdown).toContain('| Project Reader | 1 |');
    expect(securitySection).toContain('`new_projects`');
  });

  it('keeps custom tables in security filtering even when IsCustomEntity is unavailable', () => {
    const solutionWithCustomNameOnly: ParsedSolution = {
      ...sampleSolution,
      entities: [
        {
          ...sampleSolution.entities[0],
          logicalName: 'abc_orders',
          name: 'abc_orders',
          displayName: 'Orders',
          isCustom: false,
        },
      ],
      securityRoles: [
        {
          name: 'ops_role',
          displayName: 'Operations Role',
          privileges: [
            { privilegeName: 'prvReadabc_orders', depth: 4 },
            { privilegeName: 'prvReadaccount', depth: 2 },
          ],
        },
      ],
    };

    const markdown = generateMarkdown(solutionWithCustomNameOnly, {
      documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS,
        securityRoleFilters: {
          onlyTablesInCurrentSolution: false,
          onlyCustomTables: true,
        },
      },
    });

    const securitySection = markdown.split('## Security Roles')[1] ?? '';
    expect(markdown).toContain('| Operations Role | 1 |');
    expect(securitySection).toContain('`abc_orders`');
    expect(securitySection).not.toContain('`account`');
  });

  it('moves Mermaid diagrams into a companion markdown document', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        entities: [
          {
            ...sampleSolution.entities[0],
            logicalName: 'account',
            name: 'account',
            displayName: 'Account',
          },
        ],
        processes: [
          {
            name: 'Account Sync Flow',
            uniqueName: 'contoso_account_sync',
            category: ProcessCategory.PowerAutomateFlow,
            primaryEntity: 'account',
            steps: [
              {
                id: 'trigger',
                name: 'When a row is added',
                stepType: 'Trigger',
              },
            ],
          },
        ],
      },
    );

    const { mainMarkdown, companionMarkdown } = splitMarkdownForDiagramCompanion(markdown);

    expect(mainMarkdown).not.toContain('```mermaid');
    expect(mainMarkdown).toContain('companion diagrams document');
    expect(mainMarkdown).not.toContain('#### Relationship Diagram');
    expect(mainMarkdown).not.toContain('This diagram shows the process and the solution components it depends on.');
    expect(companionMarkdown).toContain('## Table of Contents');
    expect(companionMarkdown).toContain('```mermaid');
    expect(companionMarkdown).toContain('[Back to Top](#table-of-contents)');
  });

  it('includes table-of-contents anchor in consolidated markdown for Back to Top links', () => {
    const markdown = generateConsolidatedMarkdown([
      sampleSolution,
      {
        ...sampleSolution,
        metadata: {
          ...sampleSolution.metadata,
          uniqueName: 'contoso_ops',
          displayName: 'Contoso Ops',
        },
      },
    ]);

    expect(markdown).toContain('## Table of Contents');
    expect(markdown).toContain('[Back to Top](#table-of-contents)');
  });

  it('keeps security roles section when filters remove all table privileges', () => {
    const solutionWithSystemRoleOnly: ParsedSolution = {
      ...sampleSolution,
      entities: [
        {
          name: 'account',
          logicalName: 'account',
          displayName: 'Account',
          isCustom: false,
          attributes: [],
          relationships: [],
        },
      ],
      securityRoles: [
        {
          name: 'system_reader',
          displayName: 'System Reader',
          privileges: [
            { privilegeName: 'prvReadaccount', depth: 4 },
            { privilegeName: 'prvWriteaccount', depth: 2 },
          ],
        },
      ],
    };

    const markdown = generateMarkdown(solutionWithSystemRoleOnly, {
      documentationSettings: {
        ...DEFAULT_DOCUMENTATION_SETTINGS,
        securityRoleFilters: {
          onlyTablesInCurrentSolution: false,
          onlyCustomTables: true,
        },
      },
    });

    const securitySection = markdown.split('## Security Roles')[1] ?? '';

    expect(markdown).toContain('## Security Roles');
    expect(markdown).toContain('| System Reader | 0 |');
    expect(securitySection).toContain('_No table privileges matched the active filters for this role._');
    expect(securitySection).not.toContain('`account`');
  });

  it('renders security role permission cells with colored-circle badges', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        entities: [
          {
            ...sampleSolution.entities[0],
            logicalName: 'account',
            name: 'account',
            displayName: 'Account',
          },
        ],
        securityRoles: [
          {
            name: 'auditor',
            displayName: 'Auditor',
            privileges: [
              { privilegeName: 'prvReadaccount', depth: 4 },
              { privilegeName: 'prvWriteaccount', depth: 1 },
            ],
          },
        ],
      },
      {
        documentationSettings: {
          ...DEFAULT_DOCUMENTATION_SETTINGS,
          securityRoleFilters: {
            onlyTablesInCurrentSolution: true,
            onlyCustomTables: false,
          },
        },
      },
    );

    expect(markdown).toContain('Permission Depth Legend:');
    expect(markdown).toContain('🟠 Organization');
    expect(markdown).toContain('🔵 User');
    expect(markdown).toContain('<span style="color:');
  });

  it('excludes default columns from table docs while preserving relationship columns in relationship documentation and ERD', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        entities: [
          {
            ...sampleSolution.entities[0],
            logicalName: 'account',
            name: 'account',
            displayName: 'Account',
            attributes: [
              {
                name: 'new_name',
                displayName: 'Name',
                type: AttributeType.String,
                required: true,
                isCustom: true,
              },
              {
                name: 'ownerid',
                displayName: 'Owner',
                type: AttributeType.Owner,
                required: false,
                isCustom: false,
              },
              {
                name: 'createdon',
                displayName: 'Created On',
                type: AttributeType.DateTime,
                required: false,
                isCustom: false,
              },
            ],
            relationships: [
              {
                name: 'owner_account_owner',
                type: 'ManyToOne',
                referencedEntity: 'systemuser',
                referencingEntity: 'account',
                referencingAttribute: 'ownerid',
                referencedAttribute: 'systemuserid',
              },
            ],
          },
          {
            name: 'systemuser',
            logicalName: 'systemuser',
            displayName: 'System User',
            isCustom: false,
            attributes: [
              {
                name: 'systemuserid',
                displayName: 'System User ID',
                type: AttributeType.UniqueIdentifier,
                required: true,
                isCustom: false,
              },
            ],
            relationships: [],
          },
        ],
      },
      {
        documentationSettings: {
          ...DEFAULT_DOCUMENTATION_SETTINGS,
          metadata: {
            ...DEFAULT_DOCUMENTATION_SETTINGS.metadata,
            includeDefaultColumns: false,
          },
        },
      },
    );

    expect(markdown).toContain('`new_name`');
    expect(markdown).not.toContain('`ownerid` | Owner');
    expect(markdown).not.toContain('`createdon`');
    expect(markdown).toContain('#### Relationships');
    expect(markdown).toContain('owner_account_owner ownerid');
    expect(markdown).toContain('owner ownerid');
  });

  it('renders a process dependency graph for process table relationships', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        entities: [
          {
            ...sampleSolution.entities[0],
            logicalName: 'account',
            name: 'account',
            displayName: 'Account',
          },
          {
            name: 'contact',
            logicalName: 'contact',
            displayName: 'Contact',
            isCustom: false,
            attributes: [],
            relationships: [],
          },
        ],
        processes: [
          {
            name: 'Sync Contacts',
            uniqueName: 'contoso_sync_contacts',
            category: ProcessCategory.PowerAutomateFlow,
            primaryEntity: 'account',
            relatedEntities: ['contact'],
            steps: [],
          },
        ],
      },
    );

    expect(markdown).toContain('#### Relationship Diagram');
    expect(markdown).toContain('```mermaid');
    expect(markdown).toContain('flowchart LR');
    expect(markdown).toContain('process_contoso_sync_contacts["Sync Contacts - contoso_sync_contacts"]');
    expect(markdown).toContain('entity_account["Account"]');
    expect(markdown).toContain('entity_contact["Contact"]');
    expect(markdown).toContain('process_contoso_sync_contacts --> entity_account');
    expect(markdown).toContain('process_contoso_sync_contacts --> entity_contact');
    expect(markdown).toContain('linkStyle 0 stroke:#2563eb,stroke-width:3px;stroke-linecap:round');
    expect(markdown).toContain('linkStyle 1 stroke:#2563eb,stroke-width:3px;stroke-linecap:round');
  });

  it('uses configurable Mermaid palette colors when provided', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        processes: [
          {
            name: 'Palette Flow',
            uniqueName: 'palette_flow',
            category: ProcessCategory.PowerAutomateFlow,
            primaryEntity: 'account',
            relatedEntities: ['contact'],
            steps: [
              {
                id: 'trigger',
                name: 'Start',
                stepType: 'Trigger',
                children: [
                  {
                    id: 'update',
                    name: 'Update',
                    stepType: 'Dataverse.Update',
                    referencedEntities: ['account'],
                  },
                ],
              },
            ],
          },
        ],
      },
      {
        documentationSettings: {
          ...DEFAULT_DOCUMENTATION_SETTINGS,
          mermaidPalette: {
            sourceOne: '#123456',
            sourceTwo: '#654321',
            sourceThree: '#abcdef',
            sourceFour: '#fedcba',
          },
        },
      },
    );

    expect(markdown).toContain('%%{ init: { \'theme\': \'neutral\'' );
    expect(markdown).toContain('stroke:#123456,stroke-width:3px;stroke-linecap:round');
    expect(markdown).toContain('stroke:#654321,stroke-width:3px;stroke-linecap:round');
    expect(markdown).toContain('stroke:#abcdef,stroke-width:3px;stroke-linecap:round');
  });

  it('renders process flow diagrams when process steps are available', () => {
    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        entities: [
          {
            ...sampleSolution.entities[0],
            logicalName: 'account',
            name: 'account',
            displayName: 'Account',
          },
        ],
        processes: [
          {
            name: 'Account Sync Flow',
            uniqueName: 'contoso_account_sync',
            category: ProcessCategory.PowerAutomateFlow,
            primaryEntity: 'account',
            steps: [
              {
                id: 'trigger',
                name: 'When a row is added',
                stepType: 'Trigger',
                children: [
                  {
                    id: 'update',
                    name: 'Update row',
                    stepType: 'Dataverse.Update',
                    referencedEntities: ['account'],
                  },
                ],
              },
            ],
          },
        ],
      },
    );

    expect(markdown).toContain('#### Process Flow Diagram');
    expect(markdown).toContain('%% Process flow for Account Sync Flow (contoso_account_sync) %%');
    expect(markdown).toContain('flowchart TD');
    expect(markdown).toContain('process_flow_contoso_account_sync -->|starts|');
    expect(markdown).toContain('-->|then|');
    expect(markdown).toContain('-->|touches|');
    expect(markdown).toContain('linkStyle 0 stroke:#2563eb,stroke-width:3px;stroke-linecap:round');
    expect(markdown).toContain('linkStyle 1 stroke:#16a34a,stroke-width:3px;stroke-linecap:round');
    expect(markdown).toContain('linkStyle 2 stroke:#7c3aed,stroke-width:3px;stroke-linecap:round');
  });

  it('splits large component relationship graphs and large process flows for readability', () => {
    const processes = Array.from({ length: 90 }, (_, index) => ({
      name: `Bulk Flow ${index}`,
      uniqueName: `contoso_bulk_flow_${index}`,
      category: ProcessCategory.PowerAutomateFlow,
      primaryEntity: `entity_${index}`,
      relatedEntities: [`entity_${index + 1}`],
      steps: Array.from({ length: 48 }, (_, stepIndex) => ({
        id: `step_${index}_${stepIndex}`,
        name: `Step ${stepIndex}`,
        stepType: 'Action',
      })),
    }));

    const markdown = generateMarkdown(
      {
        ...sampleSolution,
        entities: Array.from({ length: 100 }, (_, index) => ({
          name: `entity_${index}`,
          logicalName: `entity_${index}`,
          displayName: `Entity ${index}`,
          isCustom: true,
          attributes: [],
          relationships: [],
        })),
        processes,
      },
    );

    expect(markdown).toContain('## Solution Component Relationship Graph');
    expect(markdown).toContain('Large dependency graph detected');
    expect(markdown).toContain('### Component Relationship Map (Part 1)');
    expect(markdown).toContain('Large process detected (48 steps). Flow diagram is split for readability.');
    expect(markdown).toContain('##### Flow Segment 1');
  });

  it('keeps large markdown generation within guardrail runtime', () => {
    const entities = Array.from({ length: 180 }, (_, index) => ({
      name: `entity_${index}`,
      logicalName: `entity_${index}`,
      displayName: `Entity ${index}`,
      isCustom: true,
      attributes: [],
      relationships: [],
    }));

    const processes = Array.from({ length: 120 }, (_, index) => ({
      name: `Perf Flow ${index}`,
      uniqueName: `perf_flow_${index}`,
      category: ProcessCategory.PowerAutomateFlow,
      primaryEntity: `entity_${index % entities.length}`,
      relatedEntities: [`entity_${(index + 1) % entities.length}`],
      steps: Array.from({ length: 20 }, (_, stepIndex) => ({
        id: `perf_step_${index}_${stepIndex}`,
        name: `Perf Step ${stepIndex}`,
        stepType: 'Action',
      })),
    }));

    const start = Date.now();
    const markdown = generateMarkdown({
      ...sampleSolution,
      entities,
      processes,
    });
    const elapsedMs = Date.now() - start;

    expect(markdown.length).toBeGreaterThan(1000);
    expect(elapsedMs).toBeLessThan(5000);
  });

  it('renders Phase 4 sections for agents, AI models, and desktop flows', () => {
    const markdown = generateMarkdown({
      ...sampleSolution,
      agents: [
        {
          name: 'sales_assistant',
          displayName: 'Sales Assistant',
          sourcePath: 'Agents/sales-assistant.json',
          agentType: 'TaskAgent',
          language: 'en-US',
          trigger: 'Teams',
          connectors: ['Dataverse', 'Outlook'],
        },
      ],
      aiModels: [
        {
          name: 'case_classifier',
          displayName: 'Case Classifier',
          sourcePath: 'AIModels/case-classifier.json',
          modelType: 'Classification',
          provider: 'Azure OpenAI',
          version: '2026-05-01',
          endpoint: 'contoso-ai',
        },
      ],
      desktopFlows: [
        {
          name: 'invoice_reconciliation',
          displayName: 'Invoice Reconciliation',
          sourcePath: 'DesktopFlows/invoice-reconciliation.json',
          folder: 'Finance',
          isEnabled: true,
          stepCount: 12,
          connectors: ['SAP', 'Excel'],
        },
      ],
    });

    expect(markdown).toContain('## Copilot Studio Agents');
    expect(markdown).toContain('## AI Models');
    expect(markdown).toContain('## Desktop Flows');
    expect(markdown).toContain('Sales Assistant');
    expect(markdown).toContain('Case Classifier');
    expect(markdown).toContain('Invoice Reconciliation');
  });

  it('renders enriched attribute metadata and choice defaults', () => {
    const markdown = generateMarkdown({
      ...sampleSolution,
      entities: [
        {
          ...sampleSolution.entities[0],
          attributes: [
            {
              name: 'new_lookup',
              displayName: 'Account Lookup',
              type: AttributeType.Lookup,
              required: false,
              isCustom: true,
              lookupTargets: ['account', 'contact'],
            },
            {
              name: 'new_budget',
              displayName: 'Budget',
              type: AttributeType.Decimal,
              required: false,
              isCustom: true,
              minValue: 0,
              maxValue: 1000,
              precision: 2,
              defaultValue: '100',
              format: 'Currency',
            },
            {
              name: 'new_status',
              displayName: 'Status',
              type: AttributeType.OptionSet,
              required: false,
              isCustom: true,
              options: [
                { value: 1, label: 'Open', isDefault: true },
                { value: 2, label: 'Closed' },
              ],
            },
          ],
        },
      ],
    });

    expect(markdown).toContain('Min: 0');
    expect(markdown).toContain('Max Value: 1000');
    expect(markdown).toContain('Format: Currency');
    expect(markdown).toContain('Default: 100');
    expect(markdown).toContain('→ `account`, `contact`');
    expect(markdown).toContain('| Label | Value | Default | Color | Description |');
    expect(markdown).toContain('| Open | 1 | ✅ |');
  });
});
