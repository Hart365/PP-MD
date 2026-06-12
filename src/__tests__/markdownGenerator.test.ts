import {
  DEFAULT_DOCUMENTATION_SETTINGS,
  generateMarkdown,
} from '../generator/markdownGenerator';
import {
  AppType,
  AttributeType,
  ProcessCategory,
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
