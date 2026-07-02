/**
 * @file markdownGenerator.ts
 * @description Converts a {@link ParsedSolution} into comprehensive Markdown
 * documentation.  Each section of the output is produced by a dedicated
 * function which assembles the text and, where applicable, embeds a Mermaid
 * diagram code block.
 *
 * Mermaid diagram types used:
 *  - erDiagram          → Entity Relationship Diagram
 *  - flowchart          → Process and component relationship diagrams
 */

import type {
  ParsedSolution,
  EntityDefinition,
  ProcessDefinition,
  ProcessStep,
  AppDefinition,
  AgentDefinition,
  AIModelDefinition,
  DesktopFlowDefinition,
  DataflowDefinition,
  CustomAPIDefinition,
  OfflineProfileDefinition,
  WebResourceDefinition,
  PluginAssemblyDefinition,
  PluginStepDefinition,
  ConnectionReferenceDefinition,
  EnvironmentVariableDefinition,
  EmailTemplateDefinition,
  SecurityRoleDefinition,
  FieldSecurityProfileDefinition,
  ReportDefinition,
  DashboardDefinition,
  OptionSetDefinition,
  SolutionDependency,
  SolutionComponentInventoryItem,
  FormDefinition,
} from '../types/solution';

import { ProcessCategory, WebResourceType, AttributeType, AppType } from '../types/solution';
import {
  createMetadataGridSection,
  createOutputSection,
  createTableSection,
  renderMarkdownDocument,
  type OutputMetadataGridRow,
  type OutputSection,
} from './outputModel';

export interface MarkdownGenerationOptions {
  erdMode?: 'compact' | 'detailed-relationships';
  documentContext?: DocumentContext;
  documentationSettings?: DocumentationSettings;
}

export type DocumentationDetailLevel = 'summary' | 'detailed';

export interface DocumentationScopeSettings {
  flows: boolean;
  apps: boolean;
  security: boolean;
  integration: boolean;
  plugins: boolean;
  reports: boolean;
}

export type AttributeSelectionMode =
  | 'all'
  | 'custom-only'
  | 'attributes-on-form'
  | 'attributes-not-on-form'
  | 'option-set-focused'
  | 'manually-selected'
  | 'unmanaged-only';

export interface DocumentationMetadataSettings {
  includeDefaultColumns: boolean;
  includeAuditInfo: boolean;
  includeFieldSecurityFlags: boolean;
  includeRequiredLevelInfo: boolean;
  includeValidForAdvancedFindInfo: boolean;
  includeMetadataDiagnosticInfo: boolean;
  excludeVirtualAttributes: boolean;
  attributeSelectionMode: AttributeSelectionMode;
  manuallySelectedAttributes: string[];
}

export interface DocumentationSecurityRoleFilters {
  onlyTablesInCurrentSolution: boolean;
  onlyCustomTables: boolean;
}

export interface DocumentationSettings {
  detailLevel: DocumentationDetailLevel;
  scope: DocumentationScopeSettings;
  metadata: DocumentationMetadataSettings;
  securityRoleFilters: DocumentationSecurityRoleFilters;
  separateDiagramsDocument: boolean;
}

export const DEFAULT_DOCUMENTATION_SETTINGS: DocumentationSettings = {
  detailLevel: 'detailed',
  scope: {
    flows: true,
    apps: true,
    security: true,
    integration: true,
    plugins: true,
    reports: true,
  },
  metadata: {
    includeDefaultColumns: true,
    includeAuditInfo: true,
    includeFieldSecurityFlags: true,
    includeRequiredLevelInfo: true,
    includeValidForAdvancedFindInfo: true,
    includeMetadataDiagnosticInfo: false,
    excludeVirtualAttributes: false,
    attributeSelectionMode: 'all',
    manuallySelectedAttributes: [],
  },
  securityRoleFilters: {
    onlyTablesInCurrentSolution: false,
    onlyCustomTables: false,
  },
  separateDiagramsDocument: false,
};

export interface DocumentContext {
  client: string;
  project: string;
  contract: string;
  sow: string;
  sprint: string;
  releaseDate: string;
}

// ---------------------------------------------------------------------------
// Utility helpers
// ---------------------------------------------------------------------------

/**
 * Escapes characters special to Markdown tables (pipe and backslash).
 *
 * @param text - Raw text to escape
 * @returns Escaped text safe for use in a Markdown table cell
 */
function mdEscape(text: string | undefined | null): string {
  if (!text) return '';
  return text.replace(/\\/g, '\\\\').replace(/\|/g, '\\|');
}

function htmlEscape(text: string | undefined | null): string {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Escapes characters that are special in Mermaid node labels:
 * quotation marks, angle brackets, and parentheses.
 *
 * @param text - Raw text
 * @returns Mermaid-safe label string
 */
function mermaidLabel(text: string | undefined | null): string {
  if (!text) return 'Unknown';
  return text
    .replace(/"/g, "'")
    .replace(/</g, '‹')
    .replace(/>/g, '›')
    .replace(/[(){}[\]]/g, '')
    .substring(0, 60); // Mermaid renders poorly with very long labels
}

function stripTrailingGuid(text: string | undefined | null): string {
  if (!text) return '';
  return text.replace(/[\s_-]?[0-9a-f]{8}(?:[\s_-]?[0-9a-f]{4}){3}[\s_-]?[0-9a-f]{12}$/i, '').trim();
}

/**
 * Returns a Markdown heading at the requested level.
 *
 * @param level - Heading level 1–6
 * @param text  - Heading text
 */
function heading(level: number, text: string): string {
  return `${'#'.repeat(Math.min(6, Math.max(1, level)))} ${text}`;
}

function headingAnchor(text: string): string {
  return `#${text
    .toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-')}`;
}

function sortByLabel<T>(items: T[], labelFn: (item: T) => string): T[] {
  return [...items].sort((a, b) => labelFn(a).localeCompare(labelFn(b), undefined, { sensitivity: 'base' }));
}

/**
 * Wraps a multi-line string in a Mermaid fenced code block.
 *
 * @param mermaidContent - Raw Mermaid DSL text (without the fence)
 * @param title          - Optional accessible title line for screen readers
 */
function mermaidBlock(mermaidContent: string, title?: string): string {
  const titleLine = title ? `%%{ init: { 'theme': 'neutral' } }%%\n%% ${title} %%\n` : '';
  return `\`\`\`mermaid\n${titleLine}${mermaidContent}\n\`\`\``;
}

/**
 * Formats a number of bytes as a human-readable string.
 *
 * @param bytes - Number of bytes
 */
function formatBytes(bytes: number | undefined): string {
  if (!bytes) return '–';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

/**
 * Returns a human-readable stage label for an SDK Message Processing Step stage.
 *
 * @param stage - Stage number (10, 20, 40)
 */
function stageLabel(stage: number): string {
  const labels: Record<number, string> = { 10: 'Pre-Validation', 20: 'Pre-Operation', 40: 'Post-Operation' };
  return labels[stage] ?? `Stage ${stage}`;
}

function modeLabel(mode: number): string {
  return mode === 0 ? 'Sync' : 'Async';
}

function processCategoryLabel(category: ProcessCategory): string {
  const labels: Record<ProcessCategory, string> = {
    [ProcessCategory.Workflow]: 'Workflow',
    [ProcessCategory.Dialog]: 'Dialog',
    [ProcessCategory.BusinessRule]: 'Business Rule',
    [ProcessCategory.Action]: 'Action',
    [ProcessCategory.BusinessProcessFlow]: 'Business Process Flow',
    [ProcessCategory.CustomAction]: 'Custom Action',
    [ProcessCategory.PowerAutomateFlow]: 'Power Automate Flow',
  };
  return labels[category] ?? category;
}

function appTypeLabel(appType: AppType): string {
  const labels: Record<AppType, string> = {
    [AppType.ModelDriven]: 'Model-Driven App',
    [AppType.Canvas]: 'Canvas App',
    [AppType.CustomPage]: 'Custom Page',
    [AppType.CodeApp]: 'Code App',
    [AppType.AIPlugin]: 'Code App',
  };
  return labels[appType] ?? appType;
}

function processStatusLabel(status: boolean | undefined, withIcon = false): string {
  if (status === undefined) return '';
  if (withIcon) return status ? '✅ Active' : '⛔ Inactive';
  return status ? 'Active' : 'Inactive';
}

function labelWithSchema(displayName: string | undefined | null, schemaName: string | undefined | null): string {
  const display = stripTrailingGuid(displayName).trim();
  const schema = stripTrailingGuid(schemaName).trim();

  if (display && schema && display.toLowerCase() !== schema.toLowerCase()) {
    return `${display} (${schema})`;
  }
  return display || schema || 'Unknown';
}

/**
 * Builds a lookup map from attribute logical name -> display name.
 *
 * Note: attribute names can repeat across entities. We keep the first
 * available display label so matrix output can show a friendly value while
 * still including the logical name for disambiguation.
 */
function buildAttributeDisplayMap(entities: EntityDefinition[]): Map<string, string> {
  const map = new Map<string, string>();
  entities.forEach((entity) => {
    entity.attributes.forEach((attribute) => {
      const key = attribute.name.toLowerCase();
      if (!map.has(key) && attribute.displayName && attribute.displayName !== attribute.name) {
        map.set(key, attribute.displayName);
      }
    });
  });
  return map;
}

function attributeLookupCandidates(attributeName: string): string[] {
  const normalized = attributeName.trim().toLowerCase();
  if (!normalized) return [];
  const candidates = new Set<string>([normalized]);

  // Some profile exports use qualified names such as table.column.
  const dotIdx = normalized.lastIndexOf('.');
  if (dotIdx > -1 && dotIdx < normalized.length - 1) {
    candidates.add(normalized.slice(dotIdx + 1));
  }

  const slashIdx = normalized.lastIndexOf('/');
  if (slashIdx > -1 && slashIdx < normalized.length - 1) {
    candidates.add(normalized.slice(slashIdx + 1));
  }

  return Array.from(candidates);
}

function resolveAttributeDisplayName(attributeName: string, attributeDisplayMap: Map<string, string>): string {
  for (const candidate of attributeLookupCandidates(attributeName)) {
    const label = attributeDisplayMap.get(candidate);
    if (label) return label;
  }
  return '';
}

function accessDepthBadge(depth: number): string {
  const normalized = Math.max(0, Math.min(4, depth));
  const stylesByDepth: Record<number, { label: string; background: string }> = {
    0: { label: 'None', background: '#fee2e2' },
    1: { label: 'User', background: '#fef9c3' },
    2: { label: 'Business Unit', background: '#dbeafe' },
    3: { label: 'Parent Child Business Unit', background: '#ede9fe' },
    4: { label: 'Org', background: '#dcfce7' },
  };
  const value = stylesByDepth[normalized] ?? { label: String(normalized), background: '#e5e7eb' };
  return `<span style="display:inline-block;padding:0.15rem 0.5rem;border-radius:0.4rem;background:${value.background};color:#1f2937;font-weight:600;">${htmlEscape(value.label)}</span>`;
}

function allowedBadge(allowed: boolean): string {
  const label = allowed ? 'Allowed' : 'Not Allowed';
  const background = allowed ? '#dcfce7' : '#fee2e2';
  return `<span style="display:inline-block;padding:0.15rem 0.5rem;border-radius:0.4rem;background:${background};color:#1f2937;font-weight:600;">${label}</span>`;
}

function processTitle(proc: ProcessDefinition): string {
  return labelWithSchema(proc.displayName || proc.name, proc.uniqueName);
}

function appTitle(app: AppDefinition): string {
  return stripTrailingGuid(app.displayName || app.name || app.uniqueName) || 'Unknown';
}

function uniqueStrings(values: Array<string | undefined | null>): string[] {
  const byLower = new Map<string, string>();
  values
    .filter((value): value is string => !!value && value.trim().length > 0)
    .forEach((value) => {
      const normalized = value.trim();
      const key = normalized.toLowerCase();
      if (!byLower.has(key)) byLower.set(key, normalized);
    });
  return Array.from(byLower.values());
}

function normalizeDocumentationSettings(settings: DocumentationSettings | undefined): DocumentationSettings {
  const isValidSelectionMode = (mode: string | undefined): mode is AttributeSelectionMode => {
    return [
      'all',
      'custom-only',
      'attributes-on-form',
      'attributes-not-on-form',
      'option-set-focused',
      'manually-selected',
      'unmanaged-only',
    ].includes(mode ?? '');
  };

  const manualAttributes = (settings?.metadata?.manuallySelectedAttributes ?? [])
    .filter((name): name is string => typeof name === 'string')
    .map((name) => name.trim().toLowerCase())
    .filter((name, idx, all) => name.length > 0 && all.indexOf(name) === idx);

  return {
    detailLevel: settings?.detailLevel === 'summary' ? 'summary' : 'detailed',
    scope: {
      flows: settings?.scope?.flows ?? DEFAULT_DOCUMENTATION_SETTINGS.scope.flows,
      apps: settings?.scope?.apps ?? DEFAULT_DOCUMENTATION_SETTINGS.scope.apps,
      security: settings?.scope?.security ?? DEFAULT_DOCUMENTATION_SETTINGS.scope.security,
      integration: settings?.scope?.integration ?? DEFAULT_DOCUMENTATION_SETTINGS.scope.integration,
      plugins: settings?.scope?.plugins ?? DEFAULT_DOCUMENTATION_SETTINGS.scope.plugins,
      reports: settings?.scope?.reports ?? DEFAULT_DOCUMENTATION_SETTINGS.scope.reports,
    },
    metadata: {
      includeDefaultColumns: settings?.metadata?.includeDefaultColumns ?? DEFAULT_DOCUMENTATION_SETTINGS.metadata.includeDefaultColumns,
      includeAuditInfo: settings?.metadata?.includeAuditInfo ?? DEFAULT_DOCUMENTATION_SETTINGS.metadata.includeAuditInfo,
      includeFieldSecurityFlags: settings?.metadata?.includeFieldSecurityFlags ?? DEFAULT_DOCUMENTATION_SETTINGS.metadata.includeFieldSecurityFlags,
      includeRequiredLevelInfo: settings?.metadata?.includeRequiredLevelInfo ?? DEFAULT_DOCUMENTATION_SETTINGS.metadata.includeRequiredLevelInfo,
      includeValidForAdvancedFindInfo:
        settings?.metadata?.includeValidForAdvancedFindInfo ?? DEFAULT_DOCUMENTATION_SETTINGS.metadata.includeValidForAdvancedFindInfo,
      includeMetadataDiagnosticInfo:
        settings?.metadata?.includeMetadataDiagnosticInfo ?? DEFAULT_DOCUMENTATION_SETTINGS.metadata.includeMetadataDiagnosticInfo,
      excludeVirtualAttributes: settings?.metadata?.excludeVirtualAttributes ?? DEFAULT_DOCUMENTATION_SETTINGS.metadata.excludeVirtualAttributes,
      attributeSelectionMode: isValidSelectionMode(settings?.metadata?.attributeSelectionMode)
        ? settings.metadata.attributeSelectionMode
        : DEFAULT_DOCUMENTATION_SETTINGS.metadata.attributeSelectionMode,
      manuallySelectedAttributes: manualAttributes,
    },
    securityRoleFilters: {
      onlyTablesInCurrentSolution: settings?.securityRoleFilters?.onlyTablesInCurrentSolution ?? DEFAULT_DOCUMENTATION_SETTINGS.securityRoleFilters.onlyTablesInCurrentSolution,
      onlyCustomTables: settings?.securityRoleFilters?.onlyCustomTables ?? DEFAULT_DOCUMENTATION_SETTINGS.securityRoleFilters.onlyCustomTables,
    },
    separateDiagramsDocument: settings?.separateDiagramsDocument ?? DEFAULT_DOCUMENTATION_SETTINGS.separateDiagramsDocument,
  };
}

function looksLikeCustomTableName(name: string | undefined): boolean {
  if (!name) return false;
  // Dataverse custom table logical/entity-set names use publisher_prefix_name.
  return /^[a-z0-9]+_[a-z0-9_]+$/i.test(name.trim());
}

function isLikelyCustomTable(entity: EntityDefinition): boolean {
  if (entity.isCustom) return true;
  if ((entity.objectTypeCode ?? 0) >= 10000) return true;
  if (looksLikeCustomTableName(entity.logicalName)) return true;
  if (looksLikeCustomTableName(entity.entitySetName)) return true;
  return false;
}

function maxDefinedNumber(...values: Array<number | undefined>): number | undefined {
  const max = Math.max(...values.map((value) => value ?? 0));
  return max > 0 ? max : undefined;
}

function tableBulletedList(values: string[]): string {
  if (values.length === 0) return '–';
  return values.map((value) => `- ${mdEscape(value)}`).join('<br>');
}

function mergeByKey<T>(items: T[], keyFn: (item: T) => string, mergeFn: (current: T, incoming: T) => T): T[] {
  const map = new Map<string, T>();
  items.forEach((item) => {
    const key = keyFn(item);
    const existing = map.get(key);
    map.set(key, existing ? mergeFn(existing, item) : item);
  });
  return Array.from(map.values());
}

/**
 * Builds a lookup map from entity logical name → display name for use across
 * all section generators.
 */
function buildEntityDisplayMap(entities: EntityDefinition[]): Map<string, string> {
  const map = new Map<string, string>();
  entities.forEach((entity) => {
    if (entity.displayName && entity.displayName !== entity.logicalName) {
      map.set(entity.logicalName.toLowerCase(), entity.displayName);
    }
  });
  return map;
}

/**
 * Returns a human-readable label for an entity: "Display Name (`logicalName`)"
 * when a display name is available, otherwise just `\`logicalName\``.
 */
function entityDisplayLabel(logicalName: string | undefined | null, entityMap: Map<string, string>): string {
  if (!logicalName) return '–';
  const display = entityMap.get(logicalName.toLowerCase());
  if (display) return `${display} (\`${logicalName}\`)`;
  return `\`${logicalName}\``;
}

type PrivilegeOperation = 'Create' | 'Read' | 'Write' | 'Delete' | 'Append' | 'AppendTo' | 'Assign' | 'Share' | 'Unshare';

/**
 * Produces a human-readable table name for an entity logical name when the
 * entity does not appear in the solution's own entity map (e.g. OOB tables).
 * Strips a common publisher prefix, splits on underscores, and title-cases
 * each word.
 */
function humanizeEntityName(logicalName: string): string {
  const words = logicalName.split('_').filter(Boolean);
  // If there's a short prefix (≤4 chars) followed by more words, drop the prefix
  const display = (words.length > 1 && words[0].length <= 4) ? words.slice(1) : words;
  return display.map((w) => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}


function parseRolePrivilege(privilegeName: string): { operation: PrivilegeOperation; table: string } | undefined {
  const match = privilegeName.match(/^prv(Create|Read|Write|Delete|AppendTo|Append|Assign|Share|Unshare)(.+)$/i);
  if (!match) return undefined;
  const operation = match[1].toLowerCase();
  const table = match[2].toLowerCase();
  const opMap: Record<string, PrivilegeOperation> = {
    create: 'Create',
    read: 'Read',
    write: 'Write',
    delete: 'Delete',
    append: 'Append',
    appendto: 'AppendTo',
    assign: 'Assign',
    share: 'Share',
    unshare: 'Unshare',
  };
  if (!opMap[operation]) return undefined;
  return { operation: opMap[operation], table };
}

interface SecurityRolePrivilegeMatrix {
  role: SecurityRoleDefinition;
  matrix: Map<string, Record<PrivilegeOperation, number>>;
  privilegeCount: number;
}

function buildSecurityRolePrivilegeMatrices(
  roles: SecurityRoleDefinition[],
  entities: EntityDefinition[],
  filters: DocumentationSecurityRoleFilters,
): SecurityRolePrivilegeMatrix[] {
  const solutionTableNames = new Set<string>();
  const customTableNames = new Set<string>();

  entities.forEach((entity) => {
    const logical = entity.logicalName.toLowerCase();
    solutionTableNames.add(logical);
    if (entity.entitySetName) {
      solutionTableNames.add(entity.entitySetName.toLowerCase());
    }

    if (isLikelyCustomTable(entity)) {
      customTableNames.add(logical);
      if (entity.entitySetName) {
        customTableNames.add(entity.entitySetName.toLowerCase());
      }
    }
  });

  return sortByLabel(roles, (role) => role.displayName || role.name)
    .map((role) => {
      const matrix = new Map<string, Record<PrivilegeOperation, number>>();
      let privilegeCount = 0;

      role.privileges.forEach((priv) => {
        const parsed = parseRolePrivilege(priv.privilegeName);
        if (!parsed) return;
        if (filters.onlyTablesInCurrentSolution && !solutionTableNames.has(parsed.table)) return;
        if (filters.onlyCustomTables && !customTableNames.has(parsed.table)) return;

        if (!matrix.has(parsed.table)) {
          matrix.set(parsed.table, {
            Create: 0,
            Read: 0,
            Write: 0,
            Delete: 0,
            Append: 0,
            AppendTo: 0,
            Assign: 0,
            Share: 0,
            Unshare: 0,
          });
        }

        const row = matrix.get(parsed.table)!;
        row[parsed.operation] = Math.max(row[parsed.operation], priv.depth);
        privilegeCount += 1;
      });

      return { role, matrix, privilegeCount };
    });
}

function processStepDescription(step: ProcessStep): string {
  const description = (step.description ?? '').trim();
  if (!description) return '';
  const normalizedDescription = description.toLowerCase();
  const normalizedType = (step.stepType || '').trim().toLowerCase();
  if (normalizedDescription === normalizedType || normalizedDescription === `type: ${normalizedType}`) {
    return '';
  }
  return description;
}

function isDocumentedApp(app: AppDefinition): boolean {
  return [AppType.ModelDriven, AppType.Canvas, AppType.CustomPage, AppType.CodeApp, AppType.AIPlugin].includes(app.appType);
}

export function consolidateSolutions(solutions: ParsedSolution[]): ParsedSolution {
  const items = solutions.filter((solution) => !!solution.metadata.uniqueName);

  const entities = mergeByKey(
    items.flatMap((solution) => solution.entities),
    (entity) => entity.logicalName.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: current.displayName || incoming.displayName,
      description: current.description || incoming.description,
      attributes: mergeByKey([...current.attributes, ...incoming.attributes], (attribute) => attribute.name.toLowerCase(), (first) => first),
      relationships: mergeByKey(
        [...current.relationships, ...incoming.relationships],
        (relationship) => `${relationship.name}|${relationship.type}|${relationship.referencedEntity}|${relationship.referencingEntity}|${relationship.referencingAttribute || ''}`.toLowerCase(),
        (first) => first,
      ),
    }),
  );

  const optionSets = mergeByKey(
    items.flatMap((solution) => solution.optionSets),
    (optionSet) => optionSet.name.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: current.displayName || incoming.displayName,
      options: mergeByKey([...current.options, ...incoming.options], (option) => String(option.value), (first) => first),
    }),
  );

  const forms = mergeByKey(items.flatMap((solution) => solution.forms), (form) => `${form.entityLogicalName}|${form.name}|${form.formType}`.toLowerCase(), (first) => first);
  const views = mergeByKey(items.flatMap((solution) => solution.views), (view) => `${view.entityLogicalName}|${view.name}|${view.viewType}`.toLowerCase(), (first) => first);

  const processes = mergeByKey(
    items.flatMap((solution) => solution.processes),
    (process) => stripTrailingGuid(process.uniqueName || process.displayName || process.name).toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: stripTrailingGuid(current.displayName || incoming.displayName || current.name || incoming.name),
      description: current.description || incoming.description,
      category: current.category === ProcessCategory.Workflow ? incoming.category : current.category,
      primaryEntity: current.primaryEntity || incoming.primaryEntity,
      relatedEntities: uniqueStrings([...(current.relatedEntities ?? []), ...(incoming.relatedEntities ?? []), current.primaryEntity, incoming.primaryEntity]),
      triggerAttributes: uniqueStrings([...(current.triggerAttributes ?? []), ...(incoming.triggerAttributes ?? [])]),
      flowConnectors: uniqueStrings([...(current.flowConnectors ?? []), ...(incoming.flowConnectors ?? [])]),
      flowTrigger: current.flowTrigger || incoming.flowTrigger,
      triggerType: current.triggerType || incoming.triggerType,
      steps: current.steps.length >= incoming.steps.length ? current.steps : incoming.steps,
      flowDefinition: current.flowDefinition || incoming.flowDefinition,
      isActivated: current.isActivated ?? incoming.isActivated,
    }),
  );

  const apps = mergeByKey(
    items.flatMap((solution) => solution.apps).filter(isDocumentedApp),
    (app) => `${app.appType}|${stripTrailingGuid(app.uniqueName || app.name || app.displayName)}`.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: stripTrailingGuid(current.displayName || incoming.displayName || current.name || incoming.name),
      entities: uniqueStrings([...(current.entities ?? []), ...(incoming.entities ?? [])]),
      sitemapAreas: uniqueStrings([...(current.sitemapAreas ?? []), ...(incoming.sitemapAreas ?? [])]),
      siteMap: (current.siteMap?.length ?? 0) >= (incoming.siteMap?.length ?? 0) ? current.siteMap : incoming.siteMap,
      siteMapSettings: {
        showHome: current.siteMapSettings?.showHome ?? incoming.siteMapSettings?.showHome,
        showPinned: current.siteMapSettings?.showPinned ?? incoming.siteMapSettings?.showPinned,
        showRecents: current.siteMapSettings?.showRecents ?? incoming.siteMapSettings?.showRecents,
        enableCollapsibleGroups: current.siteMapSettings?.enableCollapsibleGroups ?? incoming.siteMapSettings?.enableCollapsibleGroups,
      },
      canvasInsights: {
        screenCount: maxDefinedNumber(current.canvasInsights?.screenCount, incoming.canvasInsights?.screenCount),
        controlCount: maxDefinedNumber(current.canvasInsights?.controlCount, incoming.canvasInsights?.controlCount),
        dataSourceCount: maxDefinedNumber(current.canvasInsights?.dataSourceCount, incoming.canvasInsights?.dataSourceCount),
        variableCount: maxDefinedNumber(current.canvasInsights?.variableCount, incoming.canvasInsights?.variableCount),
        resourceCount: maxDefinedNumber(current.canvasInsights?.resourceCount, incoming.canvasInsights?.resourceCount),
        screenNames: uniqueStrings([...(current.canvasInsights?.screenNames ?? []), ...(incoming.canvasInsights?.screenNames ?? [])]),
        dataSources: uniqueStrings([...(current.canvasInsights?.dataSources ?? []), ...(incoming.canvasInsights?.dataSources ?? [])]),
        variables: uniqueStrings([...(current.canvasInsights?.variables ?? []), ...(incoming.canvasInsights?.variables ?? [])]),
        resources: uniqueStrings([...(current.canvasInsights?.resources ?? []), ...(incoming.canvasInsights?.resources ?? [])]),
        screens: (() => {
          const map = new Map<string, Set<string>>();
          [...(current.canvasInsights?.screens ?? []), ...(incoming.canvasInsights?.screens ?? [])].forEach((screen) => {
            if (!map.has(screen.name)) map.set(screen.name, new Set<string>());
            screen.controls.forEach((control) => map.get(screen.name)!.add(control));
          });
          return Array.from(map.entries()).map(([name, controls]) => ({ name, controls: Array.from(controls).sort((a, b) => a.localeCompare(b)) }));
        })(),
        navigation: (() => {
          const set = new Set<string>();
          [...(current.canvasInsights?.navigation ?? []), ...(incoming.canvasInsights?.navigation ?? [])].forEach((link) => {
            set.add(`${link.from}|${link.to}`);
          });
          return Array.from(set).map((item) => {
            const [from, to] = item.split('|');
            return { from, to };
          });
        })(),
      },
      connectors: uniqueStrings([...(current.connectors ?? []), ...(incoming.connectors ?? [])]),
      isEnabled: current.isEnabled !== false || incoming.isEnabled !== false,
      version: current.version || incoming.version,
    }),
  );

  const agents = mergeByKey(
    items.flatMap((solution) => solution.agents ?? []),
    (agent) => agent.sourcePath.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: current.displayName || incoming.displayName,
      agentType: current.agentType || incoming.agentType,
      language: current.language || incoming.language,
      trigger: current.trigger || incoming.trigger,
      connectors: uniqueStrings([...(current.connectors ?? []), ...(incoming.connectors ?? [])]),
    }),
  );

  const aiModels = mergeByKey(
    items.flatMap((solution) => solution.aiModels ?? []),
    (model) => model.sourcePath.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: current.displayName || incoming.displayName,
      modelType: current.modelType || incoming.modelType,
      provider: current.provider || incoming.provider,
      version: current.version || incoming.version,
      endpoint: current.endpoint || incoming.endpoint,
    }),
  );

  const desktopFlows = mergeByKey(
    items.flatMap((solution) => solution.desktopFlows ?? []),
    (flow) => flow.sourcePath.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: current.displayName || incoming.displayName,
      folder: current.folder || incoming.folder,
      isEnabled: current.isEnabled ?? incoming.isEnabled,
      stepCount: current.stepCount ?? incoming.stepCount,
      connectors: uniqueStrings([...(current.connectors ?? []), ...(incoming.connectors ?? [])]),
    }),
  );

  const dataflows = mergeByKey(
    items.flatMap((solution) => solution.dataflows ?? []),
    (flow) => flow.sourcePath.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: current.displayName || incoming.displayName,
      connectors: uniqueStrings([...(current.connectors ?? []), ...(incoming.connectors ?? [])]),
      refreshMode: current.refreshMode || incoming.refreshMode,
    }),
  );

  const customApis = mergeByKey(
    items.flatMap((solution) => solution.customApis ?? []),
    (api) => api.sourcePath.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: current.displayName || incoming.displayName,
      boundEntityLogicalName: current.boundEntityLogicalName || incoming.boundEntityLogicalName,
      isFunction: current.isFunction ?? incoming.isFunction,
    }),
  );

  const offlineProfiles = mergeByKey(
    items.flatMap((solution) => solution.offlineProfiles ?? []),
    (profile) => profile.sourcePath.toLowerCase(),
    (current, incoming) => ({
      ...current,
      displayName: current.displayName || incoming.displayName,
      profileType: current.profileType || incoming.profileType,
      entities: uniqueStrings([...(current.entities ?? []), ...(incoming.entities ?? [])]),
    }),
  );

  const webResources = mergeByKey(items.flatMap((solution) => solution.webResources), (item) => (item.schemaName || item.name).toLowerCase(), (first) => first);
  const securityRoles = mergeByKey(items.flatMap((solution) => solution.securityRoles), (role) => role.name.toLowerCase(), (current, incoming) => ({ ...current, privileges: mergeByKey([...current.privileges, ...incoming.privileges], (priv) => `${priv.privilegeName}|${priv.depth}`.toLowerCase(), (first) => first) }));
  const fieldSecurityProfiles = mergeByKey(items.flatMap((solution) => solution.fieldSecurityProfiles), (profile) => profile.name.toLowerCase(), (current, incoming) => ({ ...current, permissions: mergeByKey([...current.permissions, ...incoming.permissions], (perm) => perm.attributeName.toLowerCase(), (first) => first) }));
  const connectionReferences = mergeByKey(items.flatMap((solution) => solution.connectionReferences), (cr) => cr.name.toLowerCase(), (current, incoming) => ({ ...current, displayName: current.displayName || incoming.displayName, connectorDisplayName: current.connectorDisplayName || incoming.connectorDisplayName, connectorId: current.connectorId || incoming.connectorId, connectionId: current.connectionId || incoming.connectionId }));
  const environmentVariables = mergeByKey(items.flatMap((solution) => solution.environmentVariables), (variable) => variable.schemaName.toLowerCase(), (current, incoming) => ({ ...current, displayName: current.displayName || incoming.displayName, description: current.description || incoming.description, defaultValue: current.defaultValue || incoming.defaultValue, hasCurrentValue: current.hasCurrentValue || incoming.hasCurrentValue, currentValue: current.currentValue || incoming.currentValue }));
  const emailTemplates = mergeByKey(items.flatMap((solution) => solution.emailTemplates), (template) => `${template.name}|${template.subject || ''}`.toLowerCase(), (current, incoming) => ({ ...current, displayName: current.displayName || incoming.displayName, description: current.description || incoming.description, subject: current.subject || incoming.subject, entityLogicalName: current.entityLogicalName || incoming.entityLogicalName, templateType: current.templateType || incoming.templateType, languageCode: current.languageCode || incoming.languageCode }));
  const reports = mergeByKey(items.flatMap((solution) => solution.reports), (report) => `${report.fileName || ''}|${report.name}`.toLowerCase(), (current, incoming) => ({ ...current, displayName: current.displayName || incoming.displayName, relatedEntities: uniqueStrings([...(current.relatedEntities ?? []), ...(incoming.relatedEntities ?? [])]), category: current.category || incoming.category }));
  const dashboards = mergeByKey(items.flatMap((solution) => solution.dashboards), (dashboard) => (dashboard.name || dashboard.displayName || '').toLowerCase(), (current, incoming) => ({ ...current, displayName: current.displayName || incoming.displayName, entityLogicalName: current.entityLogicalName || incoming.entityLogicalName, dashboardType: current.dashboardType || incoming.dashboardType, components: uniqueStrings([...(current.components ?? []), ...(incoming.components ?? [])]) }));
  const pluginAssemblies = mergeByKey(items.flatMap((solution) => solution.pluginAssemblies), (assembly) => assembly.assemblyName.toLowerCase(), (current, incoming) => ({ ...current, displayName: current.displayName || incoming.displayName, version: current.version || incoming.version, culture: current.culture || incoming.culture, publicKeyToken: current.publicKeyToken || incoming.publicKeyToken, sourceType: current.sourceType || incoming.sourceType, steps: mergeByKey([...current.steps, ...incoming.steps], (step) => `${step.name}|${step.message}|${step.primaryEntity || ''}|${step.stage}|${step.mode}|${step.pluginTypeName}`.toLowerCase(), (first) => first) }));

  return {
    metadata: {
      uniqueName: 'consolidated_solutions',
      displayName: 'Consolidated Solution Summary',
      version: new Date().toISOString().slice(0, 10),
      publisherName: 'Multiple Publishers',
      isManaged: false,
      dependencies: [],
      componentInventory: [],
    },
    entities,
    optionSets,
    forms,
    views,
    processes,
    apps,
    agents,
    aiModels,
    desktopFlows,
    dataflows,
    customApis,
    offlineProfiles,
    webResources,
    securityRoles,
    fieldSecurityProfiles,
    connectionReferences,
    environmentVariables,
    emailTemplates,
    reports,
    dashboards,
    pluginAssemblies,
    warnings: uniqueStrings(items.flatMap((solution) => solution.warnings)),
  };
}

function flattenProcessSteps(steps: ProcessStep[], depth = 0): Array<{ step: ProcessStep; depth: number }> {
  const flat: Array<{ step: ProcessStep; depth: number }> = [];
  steps.forEach((step) => {
    flat.push({ step, depth });
    if (step.children && step.children.length > 0) {
      flat.push(...flattenProcessSteps(step.children, depth + 1));
    }
  });
  return flat;
}

function erEntityId(name: string): string {
  const sanitized = name.replace(/[^a-zA-Z0-9_]/g, '_');
  if (!sanitized) return 'entity_unknown';
  return /^\d/.test(sanitized) ? `entity_${sanitized}` : sanitized;
}

function flowchartNodeId(name: string): string {
  const sanitized = name.replace(/[^a-zA-Z0-9_]/g, '_');
  if (!sanitized) return 'node_unknown';
  return /^\d/.test(sanitized) ? `node_${sanitized}` : sanitized;
}

function hasDocumentContext(context: DocumentContext | undefined): boolean {
  if (!context) return false;
  return [
    context.client,
    context.project,
    context.contract,
    context.sow,
    context.sprint,
    context.releaseDate,
  ].some((value) => !!value && value.trim().length > 0);
}

function generateDocumentContextSection(context: DocumentContext | undefined): string {
  if (!hasDocumentContext(context)) return '';

  const lines: string[] = [];

  lines.push('<table>');
  lines.push(`<tr><td><strong>Client</strong></td><td>${htmlEscape(context?.client)}</td></tr>`);
  lines.push(`<tr><td><strong>Contract</strong></td><td>${htmlEscape(context?.contract)}</td></tr>`);
  lines.push(`<tr><td><strong>Contract ID/SoW</strong></td><td>${htmlEscape(context?.sow)}</td></tr>`);
  lines.push(`<tr><td><strong>Project</strong></td><td>${htmlEscape(context?.project)}</td></tr>`);
  lines.push(`<tr><td><strong>Sprint</strong></td><td>${htmlEscape(context?.sprint)}</td></tr>`);
  lines.push(`<tr><td><strong>Release Date</strong></td><td>${htmlEscape(context?.releaseDate)}</td></tr>`);
  lines.push('</table>');
  lines.push('');

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Document header
// ---------------------------------------------------------------------------

/**
 * Generates the document title and metadata table.
 *
 * @param solution - Parsed solution data
 * @returns Markdown string for the document header section
 */
function generateHeader(solution: ParsedSolution, documentContext?: DocumentContext): string {
  const { metadata } = solution;
  const lines: string[] = [];

  const contextSection = generateDocumentContextSection(documentContext);
  if (contextSection) {
    lines.push(contextSection);
  }

  lines.push(
    heading(1, `Power Platform Solution: ${metadata.displayName}`),
    '',
    `> **Generated by PP-MD** — Power Platform Solution Documentation`,
    `> Generated on: ${new Date().toLocaleString()}`,
    '',
    heading(2, 'Solution Overview'),
    '',
    `- **Unique Name:** ${mdEscape(metadata.uniqueName)}`,
    `- **Display Name:** ${mdEscape(metadata.displayName)}`,
    `- **Version:** ${mdEscape(metadata.version)}`,
    `- **Publisher:** ${mdEscape(metadata.publisherName)}`,
  );

  if (metadata.publisherPrefix) {
    lines.push(`- **Publisher Prefix:** ${mdEscape(metadata.publisherPrefix)}`);
  }

  if (metadata.solutionPrefix) {
    lines.push(`- **Solution Prefix:** ${mdEscape(metadata.solutionPrefix)}`);
  }

  lines.push(
    `- **Type:** ${metadata.isManaged ? 'Managed' : 'Unmanaged'}`,
  );

  if (metadata.description) {
    lines.push(`- **Description:** ${mdEscape(metadata.description)}`);
  }

  lines.push('');
  return lines.join('\n');
}

type ComponentGraphEdge = {
  from: string;
  to: string;
  relation: string;
};

type ComponentGraphData = {
  nodes: Map<string, string>;
  edges: ComponentGraphEdge[];
};

function graphLabel(displayName: string | undefined | null, schemaName: string | undefined | null): string {
  const display = stripTrailingGuid(displayName).trim();
  const schema = stripTrailingGuid(schemaName).trim();

  if (display && schema && display.toLowerCase() !== schema.toLowerCase()) {
    return `${display} - ${schema}`;
  }

  return display || schema || 'Unknown';
}

function entityGraphLabel(logicalName: string, entityMap: Map<string, string>): string {
  const display = entityMap.get(logicalName.toLowerCase());
  return stripTrailingGuid(display || logicalName).trim() || logicalName;
}

function collectComponentGraphData(solution: ParsedSolution, entityMap: Map<string, string>): ComponentGraphData {
  const nodes = new Map<string, string>();
  const edges: ComponentGraphEdge[] = [];
  const emittedEdges = new Set<string>();

  const nodeIdFor = (kind: string, rawName: string): string => flowchartNodeId(`${kind}_${rawName}`);
  const ensureNode = (kind: string, rawName: string, label: string): string => {
    const id = nodeIdFor(kind, rawName);
    if (!nodes.has(id)) {
      nodes.set(id, label);
    }
    return id;
  };

  const addEdge = (from: string, to: string, relation: string) => {
    const key = `${from}|${to}|${relation}`;
    if (emittedEdges.has(key)) return;
    emittedEdges.add(key);
    edges.push({ from, to, relation });
  };

  sortByLabel(solution.forms, (form) => form.displayName || form.name).forEach((form) => {
    const formNode = ensureNode('form', `${form.entityLogicalName}_${form.name}`, graphLabel(form.displayName || form.name, undefined));
    const entityNode = ensureNode('entity', form.entityLogicalName, entityGraphLabel(form.entityLogicalName, entityMap));
    addEdge(formNode, entityNode, 'binds');
  });

  sortByLabel(solution.views, (view) => view.displayName || view.name).forEach((view) => {
    const viewNode = ensureNode('view', `${view.entityLogicalName}_${view.name}`, graphLabel(view.displayName || view.name, undefined));
    const entityNode = ensureNode('entity', view.entityLogicalName, entityGraphLabel(view.entityLogicalName, entityMap));
    addEdge(viewNode, entityNode, 'queries');
  });

  sortByLabel(solution.processes, (process) => processTitle(process)).forEach((process) => {
    const processNode = ensureNode('process', process.uniqueName || process.name, graphLabel(process.displayName || process.name, process.uniqueName));

    uniqueStrings([process.primaryEntity, ...(process.relatedEntities ?? [])].filter(Boolean) as string[])
      .forEach((table) => {
        const tableNode = ensureNode('entity', table, entityGraphLabel(table, entityMap));
        addEdge(processNode, tableNode, 'uses');
      });

    sortByLabel(process.flowConnectionReferences ?? [], (ref) => ref).forEach((ref) => {
      const refNode = ensureNode('connection', ref, graphLabel('Connection Ref', ref));
      addEdge(processNode, refNode, 'connects');
    });

    sortByLabel(process.flowEnvironmentVariables ?? [], (env) => env).forEach((env) => {
      const envNode = ensureNode('environment', env, graphLabel('Environment Var', env));
      addEdge(processNode, envNode, 'reads');
    });
  });

  sortByLabel(solution.apps.filter(isDocumentedApp), (app) => appTitle(app)).forEach((app) => {
    const appNode = ensureNode('app', app.uniqueName, graphLabel(app.displayName || app.name, app.uniqueName));
    sortByLabel(app.entities ?? [], (entity) => entity).forEach((entity) => {
      const entityNode = ensureNode('entity', entity, entityGraphLabel(entity, entityMap));
      addEdge(appNode, entityNode, 'surfaces');
    });
  });

  sortByLabel(solution.pluginAssemblies, (assembly) => assembly.name).forEach((assembly) => {
    sortByLabel(assembly.steps ?? [], (step) => step.name).forEach((step) => {
      if (!step.primaryEntity) return;
      const stepNode = ensureNode('plugin-step', `${assembly.name}_${step.name}`, graphLabel(step.name, assembly.name));
      const entityNode = ensureNode('entity', step.primaryEntity, entityGraphLabel(step.primaryEntity, entityMap));
      addEdge(stepNode, entityNode, 'executes');
    });
  });

  sortByLabel(solution.reports, (report) => report.displayName || report.name).forEach((report) => {
    const relatedEntities = uniqueStrings(report.relatedEntities ?? []);
    if (relatedEntities.length === 0) return;
    const reportNode = ensureNode('report', report.name, graphLabel(report.displayName || report.name, report.fileName));
    relatedEntities.forEach((relatedEntity) => {
      const entityNode = ensureNode('entity', relatedEntity, entityGraphLabel(relatedEntity, entityMap));
      addEdge(reportNode, entityNode, 'reports');
    });
  });

  sortByLabel(solution.dashboards, (dashboard) => dashboard.displayName || dashboard.name).forEach((dashboard) => {
    if (!dashboard.entityLogicalName) return;
    const dashboardNode = ensureNode('dashboard', dashboard.name, graphLabel(dashboard.displayName || dashboard.name, dashboard.dashboardType));
    const entityNode = ensureNode('entity', dashboard.entityLogicalName, entityGraphLabel(dashboard.entityLogicalName, entityMap));
    addEdge(dashboardNode, entityNode, 'visualizes');
  });

  return {
    nodes,
    edges: [...edges].sort((a, b) => `${a.from}|${a.to}|${a.relation}`.localeCompare(`${b.from}|${b.to}|${b.relation}`)),
  };
}

function componentRelationshipEdgeCount(solution: ParsedSolution): number {
  const entityMap = buildEntityDisplayMap(solution.entities);
  return collectComponentGraphData(solution, entityMap).edges.length;
}

// ---------------------------------------------------------------------------
// Table of Contents
// ---------------------------------------------------------------------------

/**
 * Generates a Markdown table of contents based on which sections are populated.
 *
 * @param solution - Parsed solution data
 * @returns Markdown ToC string
 */
function generateTableOfContents(solution: ParsedSolution, settings: DocumentationSettings): string {
  const lines: string[] = [heading(2, 'Table of Contents'), ''];
  const documentedApps = solution.apps.filter(isDocumentedApp);
  const includeDetailedSections = settings.detailLevel === 'detailed';
  const securityRoleMatrices = buildSecurityRolePrivilegeMatrices(
    solution.securityRoles,
    solution.entities,
    settings.securityRoleFilters,
  );

  const sections: Array<{ label: string; count: number }> = [
    { label: 'Solution Dependencies',          count: solution.metadata.dependencies.length },
    { label: 'Solution Component Inventory',   count: solution.metadata.componentInventory.length },
    { label: 'Solution Component Relationship Graph', count: includeDetailedSections ? componentRelationshipEdgeCount(solution) : 0 },
    { label: 'Entity Relationship Diagram',    count: includeDetailedSections ? solution.entities.length : 0 },
    { label: 'Tables & Columns',               count: includeDetailedSections ? solution.entities.length : 0 },
    { label: 'Global Option Sets',             count: includeDetailedSections ? solution.optionSets.length : 0 },
    { label: 'Forms & Views',                  count: includeDetailedSections ? solution.forms.length + solution.views.length : 0 },
    { label: 'Processes & Automation',         count: settings.scope.flows ? solution.processes.length : 0 },
    { label: 'Power Apps',                     count: settings.scope.apps ? documentedApps.length : 0 },
    { label: 'Copilot Studio Agents',          count: includeDetailedSections ? solution.agents.length : 0 },
    { label: 'AI Models',                      count: includeDetailedSections ? solution.aiModels.length : 0 },
    { label: 'Desktop Flows',                  count: includeDetailedSections ? solution.desktopFlows.length : 0 },
    { label: 'Dataflows',                      count: includeDetailedSections ? (solution.dataflows?.length ?? 0) : 0 },
    { label: 'Custom APIs',                    count: includeDetailedSections ? (solution.customApis?.length ?? 0) : 0 },
    { label: 'Offline Profiles',               count: includeDetailedSections ? (solution.offlineProfiles?.length ?? 0) : 0 },
    { label: 'Web Resources',                  count: includeDetailedSections ? solution.webResources.length : 0 },
    { label: 'Security Roles',                 count: settings.scope.security ? securityRoleMatrices.length : 0 },
    { label: 'Column Level Security Profiles', count: settings.scope.security ? solution.fieldSecurityProfiles.length : 0 },
    { label: 'Connection References',          count: settings.scope.integration ? solution.connectionReferences.length : 0 },
    { label: 'Environment Variables',          count: settings.scope.integration ? solution.environmentVariables.length : 0 },
    { label: 'Email Templates',                count: settings.scope.integration ? solution.emailTemplates.length : 0 },
    { label: 'Reports & Dashboards',           count: settings.scope.reports ? solution.reports.length + solution.dashboards.length : 0 },
    { label: 'Plugin Assemblies & Steps',      count: settings.scope.plugins ? solution.pluginAssemblies.length : 0 },
  ];

  sections.forEach(({ label, count }) => {
    if (count > 0) {
      lines.push(`- [${label}](${headingAnchor(label)})`);
    }
  });

  lines.push('');
  return lines.join('\n');
}

function generateComponentInventorySection(inventory: SolutionComponentInventoryItem[]): string {
  if (inventory.length === 0) return '';

  const documentedTypes = new Set([
    'Entity',
    'Attribute',
    'Relationship',
    'OptionSet',
    'Form',
    'SavedQuery',
    'SystemForm',
    'Workflow',
    'PowerAutomateFlow',
    'AppModule',
    'CanvasApp',
    'Dataflow',
    'CustomAPI',
    'MobileOfflineProfile',
    'WebResource',
    'Role',
    'FieldSecurityProfile',
    'ConnectionReference',
    'EnvironmentVariableDefinition',
    'EnvironmentVariableValue',
    'Report',
    'Dashboard',
    'PluginAssembly',
    'PluginStep',
    'SdkMessageProcessingStep',
  ]);

  return renderMarkdownDocument([
    createTableSection({
      id: 'solution-component-inventory',
      title: 'Solution Component Inventory',
      headers: ['Component Type', 'Count', 'Covered in PP-MD'],
      rows: sortByLabel(inventory, (item) => item.componentType).map((item) => {
        const covered = documentedTypes.has(item.componentType) ? '✅' : '⚠️ Summary only';
        return [mdEscape(item.componentType), String(item.count), covered];
      }),
    }),
  ]);
}

function generateComponentRelationshipGraphSection(solution: ParsedSolution, entityMap: Map<string, string>): string {
  const graph = collectComponentGraphData(solution, entityMap);
  if (graph.edges.length === 0) return '';

  const lines: string[] = [
    heading(2, 'Solution Component Relationship Graph'),
    '',
    '> Component dependencies across forms, views, processes, apps, plugins, and reporting artifacts.',
    '',
  ];

  const buildDiagram = (edges: ComponentGraphEdge[]): string => {
    const nodeIds = Array.from(new Set(edges.flatMap((edge) => [edge.from, edge.to]))).sort();
    const diagramLines: string[] = ['flowchart LR'];

    nodeIds.forEach((nodeId) => {
      const label = graph.nodes.get(nodeId) ?? nodeId;
      diagramLines.push(`  ${nodeId}["${mermaidLabel(label)}"]`);
    });

    edges.forEach((edge) => {
      diagramLines.push(`  ${edge.from} -->|${mermaidLabel(edge.relation)}| ${edge.to}`);
    });

    return diagramLines.join('\n');
  };

  const MAX_EDGES_PER_DIAGRAM = 80;
  if (graph.edges.length <= MAX_EDGES_PER_DIAGRAM) {
    lines.push(mermaidBlock(buildDiagram(graph.edges), 'Solution Component Relationship Graph'));
    lines.push('');
    return lines.join('\n');
  }

  lines.push(`> Large dependency graph detected (${graph.edges.length} relationships). Diagram is split for readability.`);
  lines.push('');

  const chunks: ComponentGraphEdge[][] = [];
  for (let i = 0; i < graph.edges.length; i += MAX_EDGES_PER_DIAGRAM) {
    chunks.push(graph.edges.slice(i, i + MAX_EDGES_PER_DIAGRAM));
  }

  chunks.forEach((chunk, index) => {
    lines.push(heading(3, `Component Relationship Map (Part ${index + 1})`));
    lines.push('');
    lines.push(mermaidBlock(buildDiagram(chunk), `Component Relationship Graph Part ${index + 1}`));
    lines.push('');
  });

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Solution Dependencies
// ---------------------------------------------------------------------------

/**
 * Documents all solution dependencies required by this solution.
 *
 * @param dependencies - Array of solution dependencies
 * @returns Markdown string
 */
function generateDependenciesSection(dependencies: SolutionDependency[]): string {
  if (dependencies.length === 0) return '';

  return renderMarkdownDocument([
    createTableSection({
      id: 'solution-dependencies',
      title: 'Solution Dependencies',
      introMarkdown:
        '> This solution depends on the following solutions. ' +
        'These must be imported first to ensure all dependencies are satisfied.',
      headers: ['Solution Name', 'Display Name', 'Version', 'Internal'],
      rows: sortByLabel(dependencies, (dep) => dep.displayName || dep.solutionName).map((dep) => [
        `\`${mdEscape(dep.solutionName)}\``,
        mdEscape(dep.displayName || dep.solutionName),
        dep.version || '–',
        dep.isInternal ? '✅' : '',
      ]),
    }),
  ]);
}

// ---------------------------------------------------------------------------
// ERD
// ---------------------------------------------------------------------------

/**
 * Generates a Mermaid erDiagram showing all entities and their relationships.
 * Limits to keep the diagram readable in browsers.
 *
 * @param entities - List of parsed entity definitions
 * @returns Markdown section with embedded Mermaid ERD
 */
function generateERD(
  entities: EntityDefinition[],
  options: MarkdownGenerationOptions,
  settings: DocumentationSettings,
): string {
  if (entities.length === 0) return '';

  const lines: string[] = [
    heading(2, 'Entity Relationship Diagram'),
    '',
    '> ERDs use Mermaid `erDiagram` syntax with top-down layout and crow\'s foot relationship lines.',
    '',
  ];

  const mode = options.erdMode ?? 'detailed-relationships';
  const includedNames = new Set(entities.map((e) => e.logicalName.toLowerCase()));
  const entityMap = new Map(entities.map((entity) => [entity.logicalName.toLowerCase(), entity]));

  interface ErRelationship {
    from: string;
    to: string;
    marker: '||--o{' | '}o--o{';
    label: string;
  }

  const relationships: ErRelationship[] = [];
  const emittedRelationships = new Set<string>();
  const degree = new Map<string, number>();

  const increaseDegree = (name: string) => {
    degree.set(name, (degree.get(name) ?? 0) + 1);
  };

  entities.forEach((entity) => {
    entity.relationships.forEach((rel) => {
      const referenced = rel.referencedEntity.toLowerCase();
      const referencing = rel.referencingEntity.toLowerCase();
      if (!includedNames.has(referenced) || !includedNames.has(referencing)) return;

      const relColumn = rel.referencingAttribute ? ` [${rel.referencingAttribute}]` : '';
      const relLabel = mermaidLabel(`${rel.name || `${referenced}_${referencing}`}${relColumn}`);

      if (rel.type === 'ManyToMany') {
        const [left, right] = [referenced, referencing].sort();
        const key = `${left}|}o--o{|${right}|${relLabel}`;
        if (emittedRelationships.has(key)) return;
        emittedRelationships.add(key);
        relationships.push({ from: left, to: right, marker: '}o--o{', label: relLabel });
        increaseDegree(left);
        increaseDegree(right);
        return;
      }

      // Dataverse relationship metadata consistently exposes referenced (parent/one)
      // and referencing (child/many). Render as one-to-many crow's foot.
      const key = `${referenced}|||--o{|${referencing}|${relLabel}`;
      if (emittedRelationships.has(key)) return;
      emittedRelationships.add(key);
      relationships.push({ from: referenced, to: referencing, marker: '||--o{', label: relLabel });
      increaseDegree(referenced);
      increaseDegree(referencing);
    });
  });

  const buildDiagram = (entityNames: string[], rels: ErRelationship[]): string => {
    const diagram: string[] = ['erDiagram', '  direction TB'];

    entityNames.forEach((entityName) => {
      const entity = entityMap.get(entityName);
      if (!entity) return;
      const entityId = erEntityId(entity.logicalName);
      // Use "Display Name" alias syntax: EntityId["Display Name"]
      const displayLabel = entity.displayName && entity.displayName !== entity.logicalName
        ? mermaidLabel(entity.displayName)
        : undefined;

      if (mode === 'compact') {
        diagram.push(displayLabel ? `  ${entityId}["${displayLabel}"]` : `  ${entityId}`);
        return;
      }

      const relationshipColumns = new Set<string>();
      entity.relationships.forEach((rel) => {
        if (rel.referencingAttribute) relationshipColumns.add(rel.referencingAttribute.toLowerCase());
        if (rel.referencedAttribute) relationshipColumns.add(rel.referencedAttribute.toLowerCase());
      });

      const attrsToShow = entity.attributes.filter((attr) =>
        ((settings.metadata.includeDefaultColumns || attr.isCustom || relationshipColumns.has(attr.name.toLowerCase())) && (
          relationshipColumns.has(attr.name.toLowerCase()) ||
          attr.type === AttributeType.Lookup ||
          attr.type === AttributeType.Owner ||
          attr.type === AttributeType.Customer ||
          attr.type === AttributeType.PartyList ||
          /id$/i.test(attr.name)
        )),
      ).slice(0, 12);

      diagram.push(displayLabel ? `  ${entityId}["${displayLabel}"] {` : `  ${entityId} {`);
      if (attrsToShow.length === 0) {
        diagram.push('    string logical_name PK');
      } else {
        attrsToShow.forEach((attr) => {
          const attrType = attr.type === AttributeType.Unknown ? 'string' : attr.type.toString().toLowerCase();
          const attrName = erEntityId(attr.name);
          const marker = attr.required ? ' PK' : (attr.lookupTarget ? ' FK' : '');
          diagram.push(`    ${attrType} ${attrName}${marker}`);
        });
      }
      diagram.push('  }');
    });

    rels.forEach((rel) => {
      diagram.push(`  ${erEntityId(rel.from)} ${rel.marker} ${erEntityId(rel.to)} : ${rel.label}`);
    });

    return diagram.join('\n');
  };

  const sortedEntityNames = Array.from(includedNames).sort();
  if (relationships.length === 0) {
    lines.push(mermaidBlock(buildDiagram(sortedEntityNames, []), 'Entity Relationship Diagram'));
    lines.push('');
    return lines.join('\n');
  }

  const relationshipDensity = relationships.length / Math.max(1, sortedEntityNames.length);
  let MAX_RELATIONSHIPS_PER_DIAGRAM = 45;
  let HUB_THRESHOLD = 12;

  // Density-aware tuning to avoid clutter in dense models and unnecessary splitting
  // in sparse models.
  if (relationshipDensity < 0.5) {
    MAX_RELATIONSHIPS_PER_DIAGRAM = 120;
    HUB_THRESHOLD = 20;
  } else if (relationshipDensity < 1.25) {
    MAX_RELATIONSHIPS_PER_DIAGRAM = 80;
    HUB_THRESHOLD = 16;
  } else if (relationshipDensity < 2.5) {
    MAX_RELATIONSHIPS_PER_DIAGRAM = 55;
    HUB_THRESHOLD = 12;
  } else {
    MAX_RELATIONSHIPS_PER_DIAGRAM = 35;
    HUB_THRESHOLD = 8;
  }
  const sortedRelationships = [...relationships].sort((a, b) => `${a.from}:${a.to}`.localeCompare(`${b.from}:${b.to}`));

  if (sortedRelationships.length <= MAX_RELATIONSHIPS_PER_DIAGRAM) {
    lines.push(mermaidBlock(buildDiagram(sortedEntityNames, sortedRelationships), 'Entity Relationship Diagram'));
    lines.push('');
    return lines.join('\n');
  }

  lines.push(`> Large model detected (${sortedRelationships.length} relationships). Diagram is split for readability.`);
  lines.push('');

  const chunks: ErRelationship[][] = [];
  for (let i = 0; i < sortedRelationships.length; i += MAX_RELATIONSHIPS_PER_DIAGRAM) {
    chunks.push(sortedRelationships.slice(i, i + MAX_RELATIONSHIPS_PER_DIAGRAM));
  }

  chunks.forEach((chunk, index) => {
    const chunkEntities = Array.from(new Set(chunk.flatMap((rel) => [rel.from, rel.to]))).sort();
    lines.push(heading(3, `ERD Relationship Map (Part ${index + 1})`));
    lines.push('');
    lines.push(mermaidBlock(buildDiagram(chunkEntities, chunk), `Entity Relationship Diagram Part ${index + 1}`));
    lines.push('');
  });

  const hubs = Array.from(degree.entries())
    .filter(([, relCount]) => relCount >= HUB_THRESHOLD)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  hubs.forEach(([hubName, relCount]) => {
    const directEdges = sortedRelationships.filter((rel) => rel.from === hubName || rel.to === hubName);
    if (directEdges.length === 0) return;

    const neighbors = Array.from(new Set(directEdges.flatMap((rel) => [rel.from, rel.to])));
    const focusedEntities = neighbors.slice(0, 22);
    if (!focusedEntities.includes(hubName)) focusedEntities.unshift(hubName);
    const focusedSet = new Set(focusedEntities);

    const focusedEdges = sortedRelationships
      .filter((rel) => focusedSet.has(rel.from) && focusedSet.has(rel.to))
      .slice(0, MAX_RELATIONSHIPS_PER_DIAGRAM);

    const hubLabel = entityMap.get(hubName)?.displayName || hubName;
    lines.push(heading(3, `Focused ERD: ${hubLabel} (${relCount} relationships)`));
    lines.push('');
    lines.push(mermaidBlock(buildDiagram(focusedEntities.sort(), focusedEdges), `Focused ERD for ${hubLabel}`));
    lines.push('');
  });

  lines.push('');
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Tables & Columns detail
// ---------------------------------------------------------------------------

/**
 * Generates detailed documentation for every entity including all columns,
 * option sets, and relationship summaries.
 *
 * @param entities   - List of entity definitions
 * @param optionSets - Global option sets (for cross-referencing)
 * @returns Markdown string
 */
function generateEntitiesSection(
  entities: EntityDefinition[],
  optionSets: OptionSetDefinition[],
  entityMap: Map<string, string>,
  forms: FormDefinition[],
  fieldSecurityProfiles: FieldSecurityProfileDefinition[],
  settings: DocumentationSettings,
): string {
  if (entities.length === 0) return '';

  const lines: string[] = [heading(2, 'Tables & Columns'), ''];
  const metadataSettings = settings.metadata;

  const securedAttributeNames = new Set(
    fieldSecurityProfiles
      .flatMap((profile) => profile.permissions)
      .map((permission) => permission.attributeName.toLowerCase()),
  );

  const filterAttributesForEntity = (entity: EntityDefinition) => {
    const formFieldNames = new Set(
      forms
        .filter((form) => form.entityLogicalName.toLowerCase() === entity.logicalName.toLowerCase())
        .flatMap((form) => form.fields)
        .map((field) => field.attributeName.toLowerCase()),
    );

    const withoutVirtual = metadataSettings.excludeVirtualAttributes
      ? entity.attributes.filter((attr) => attr.type !== AttributeType.Virtual)
      : entity.attributes;

    const withoutDefaultColumns = metadataSettings.includeDefaultColumns
      ? withoutVirtual
      : withoutVirtual.filter((attr) => attr.isCustom);

    switch (metadataSettings.attributeSelectionMode) {
      case 'attributes-on-form':
        return withoutDefaultColumns.filter((attr) => formFieldNames.has(attr.name.toLowerCase()));
      case 'attributes-not-on-form':
        return withoutDefaultColumns.filter((attr) => !formFieldNames.has(attr.name.toLowerCase()));
      case 'custom-only':
        return withoutDefaultColumns.filter((attr) => attr.isCustom);
      case 'option-set-focused':
        return withoutDefaultColumns.filter((attr) => attr.type === AttributeType.OptionSet || attr.type === AttributeType.MultiSelectOptionSet);
      case 'manually-selected': {
        const selected = new Set(metadataSettings.manuallySelectedAttributes);
        return withoutDefaultColumns.filter((attr) => selected.has(attr.name.toLowerCase()));
      }
      case 'unmanaged-only':
        return withoutDefaultColumns.filter((attr) => attr.isManaged === undefined ? attr.isCustom : !attr.isManaged);
      case 'all':
      default:
        return withoutDefaultColumns;
    }
  };

  sortByLabel(entities, (entity) => labelWithSchema(entity.displayName || entity.logicalName, entity.logicalName)).forEach((entity) => {
    const entityDisplayName = entity.displayName || entity.logicalName;
    const selectedAttributes = filterAttributesForEntity(entity);

    lines.push(heading(3, `${entityDisplayName} (${entity.logicalName})`));
    lines.push('');

    // Entity metadata
    lines.push('| Property | Value |');
    lines.push('|----------|-------|');
    lines.push(`| **Logical Name** | \`${entity.logicalName}\` |`);
    if (entity.entitySetName) {
      lines.push(`| **Entity Set Name** | \`${entity.entitySetName}\` |`);
    }
    if (entity.objectTypeCode) {
      lines.push(`| **Object Type Code** | ${entity.objectTypeCode} |`);
    }
    lines.push(`| **Ownership** | ${entity.ownershipType ?? 'User'} |`);
    lines.push(`| **Custom** | ${entity.isCustom ? 'Yes' : 'No'} |`);
    if (entity.primaryAttributeName) {
      lines.push(`| **Primary Attribute** | \`${entity.primaryAttributeName}\` |`);
    }
    if (entity.isActivity) lines.push(`| **Activity** | Yes |`);
    if (entity.changeTracking) lines.push(`| **Change Tracking** | Enabled |`);
    if (entity.description) lines.push(`| **Description** | ${mdEscape(entity.description)} |`);
    lines.push('');

    // Columns
    if (selectedAttributes.length > 0) {
      lines.push(heading(4, 'Columns'));
      lines.push('');

      const hasDescriptions = selectedAttributes.some((a) => !!a.description);
      const includeAuditInfo = metadataSettings.includeAuditInfo;
      const includeRequiredInfo = metadataSettings.includeRequiredLevelInfo;
      const includeFieldSecurityFlags = metadataSettings.includeFieldSecurityFlags;
      const includeAdvancedFind = metadataSettings.includeValidForAdvancedFindInfo;
      const includeMetadataDiagnostics = metadataSettings.includeMetadataDiagnosticInfo;

      const headers = ['Display Name', 'Schema Name', 'Type'];
      if (includeRequiredInfo) headers.push('Required');
      headers.push('Custom');
      if (includeAuditInfo) headers.push('Audited');
      if (includeFieldSecurityFlags) headers.push('Field Security');
      if (includeAdvancedFind) headers.push('Advanced Find');
      if (includeMetadataDiagnostics) headers.push('Metadata Source');
      headers.push('Notes');
      if (hasDescriptions) headers.push('Description');

      lines.push(`| ${headers.join(' | ')} |`);
      lines.push(`| ${headers.map(() => '---').join(' | ')} |`);

      sortByLabel(selectedAttributes, (attr) => labelWithSchema(attr.displayName || attr.name, attr.name)).forEach((attr) => {
        const notes: string[] = [];
        if (attr.isPrimaryName) notes.push('🔑 Primary Name');
        if ((attr.lookupTargets?.length ?? 0) > 0) {
          notes.push(`→ ${sortByLabel(attr.lookupTargets ?? [], (item) => item).map((target) => entityDisplayLabel(target, entityMap)).join(', ')}`);
        } else if (attr.lookupTarget) {
          notes.push(`→ ${entityDisplayLabel(attr.lookupTarget, entityMap)}`);
        }
        if (attr.optionSetName) notes.push(`OptionSet: ${attr.optionSetName}`);
        if (attr.maxLength) notes.push(`Max Length: ${attr.maxLength}`);
        if (attr.precision) notes.push(`Precision: ${attr.precision}`);
        if (attr.minValue !== undefined) notes.push(`Min: ${attr.minValue}`);
        if (attr.maxValue !== undefined) notes.push(`Max Value: ${attr.maxValue}`);
        if (attr.format) notes.push(`Format: ${attr.format}`);
        if (attr.defaultValue !== undefined && attr.defaultValue !== '') notes.push(`Default: ${attr.defaultValue}`);

        const rowCells: string[] = [
          mdEscape(attr.displayName || attr.name),
          `\`${mdEscape(attr.name)}\``,
          attr.type,
        ];

        if (includeRequiredInfo) {
          const requiredLevel = attr.requiredLevel ? ` (${attr.requiredLevel})` : '';
          const isRequired = attr.required || ['required', 'systemrequired', 'applicationrequired'].includes((attr.requiredLevel || '').toLowerCase());
          rowCells.push(`${isRequired ? '✅ Yes' : 'No'}${requiredLevel}`);
        }

        rowCells.push(attr.isCustom ? '✳️ Yes' : 'No');

        if (includeAuditInfo) {
          rowCells.push(attr.isAuditEnabled === undefined ? '–' : attr.isAuditEnabled ? '🔍 Yes' : 'No');
        }

        if (includeFieldSecurityFlags) {
          const isSecured = attr.isSecured ?? securedAttributeNames.has(attr.name.toLowerCase());
          rowCells.push(isSecured ? '🔐 Yes' : 'No');
        }

        if (includeAdvancedFind) {
          rowCells.push(
            attr.isValidForAdvancedFind === undefined
              ? '–'
              : attr.isValidForAdvancedFind
                ? '✅ Yes'
                : 'No',
          );
        }

        if (includeMetadataDiagnostics) {
          const diagParts: string[] = [];
          if (attr.metadataSources?.isCustom) {
            diagParts.push(`Custom<=${attr.metadataSources.isCustom}`);
          }
          if (attr.metadataSources?.isValidForAdvancedFind) {
            diagParts.push(`AdvancedFind<=${attr.metadataSources.isValidForAdvancedFind}`);
          }
          rowCells.push(mdEscape(diagParts.join(', ')) || '–');
        }

        rowCells.push(mdEscape(notes.join(', ')));
        if (hasDescriptions) {
          rowCells.push(mdEscape(attr.description));
        }

        lines.push(`| ${rowCells.join(' | ')} |`);
      });
      lines.push('');
    } else {
      lines.push('_No attributes matched the active attribute selection settings for this table._');
      lines.push('');
    }

    // Relationships
    if (entity.relationships.length > 0) {
      lines.push(heading(4, 'Relationships'));
      lines.push('');

      const hasRelDescriptions = entity.relationships.some((r) => !!r.relationshipDescription);
      const hasCascade = entity.relationships.some((r) => !!r.cascadeDelete);
      const hasParentFK = entity.relationships.some((r) => !!r.referencedAttribute);

      // Build header
      let relHeader = '| Relationship Name | Type | Related Entity | FK Column |';
      let relSep = '|-------------------|------|----------------|-----------|';
      if (hasParentFK) { relHeader += ' Parent Column |'; relSep += '---------------|'; }
      if (hasCascade)  { relHeader += ' On Delete |'; relSep += '-----------|'; }
      if (hasRelDescriptions) { relHeader += ' Description |'; relSep += '-------------|'; }
      lines.push(relHeader);
      lines.push(relSep);

      sortByLabel(entity.relationships, (rel) => rel.name).forEach((rel) => {
        const relatedEntity = rel.type === 'OneToMany' ? rel.referencingEntity : rel.referencedEntity;
        let row = `| \`${mdEscape(rel.name)}\` ` +
          `| ${rel.type} ` +
          `| ${entityDisplayLabel(relatedEntity, entityMap)} ` +
          `| ${rel.referencingAttribute ? `\`${rel.referencingAttribute}\`` : '–'} |`;
        if (hasParentFK) row += ` ${rel.referencedAttribute ? `\`${rel.referencedAttribute}\`` : '–'} |`;
        if (hasCascade)  row += ` ${rel.cascadeDelete || '–'} |`;
        if (hasRelDescriptions) row += ` ${mdEscape(rel.relationshipDescription)} |`;
        lines.push(row);
      });
      lines.push('');
    }

    // Inline OptionSet options (for local optionsets)
    const localOSAttrs = selectedAttributes.filter(
      (a) => (a.type === AttributeType.OptionSet || a.type === AttributeType.MultiSelectOptionSet) && a.options,
    );
    if (localOSAttrs.length > 0) {
      lines.push(heading(4, 'Choice Column Values'));
      sortByLabel(localOSAttrs, (attr) => labelWithSchema(attr.displayName || attr.name, attr.name)).forEach((attr) => {
        if (!attr.options) return;
        lines.push(`**${attr.displayName || attr.name} (\`${attr.name}\`)**:`);
        lines.push('');
        lines.push('| Label | Value | Default | Color | Description |');
        lines.push('|-------|-------|---------|-------|-------------|');
        sortByLabel(attr.options, (opt) => opt.label || String(opt.value)).forEach((opt) => {
          lines.push(`| ${mdEscape(opt.label)} | ${opt.value} | ${opt.isDefault ? '✅' : '–'} | ${mdEscape(opt.color)} | ${mdEscape(opt.description)} |`);
        });
        lines.push('');
      });
    }

    // Cross-reference global option sets used
    const globalOSRefs = selectedAttributes
      .filter((a) => a.optionSetName)
      .map((a) => a.optionSetName as string);
    if (globalOSRefs.length > 0) {
      const globalDefs = optionSets.filter((os) => globalOSRefs.includes(os.name));
      if (globalDefs.length > 0) {
        lines.push(heading(4, 'Global Choice References'));
        sortByLabel(globalDefs, (os) => labelWithSchema(os.displayName || os.name, os.name)).forEach((os) => {
          lines.push(`**${os.displayName || os.name} (\`${os.name}\`)**:`);
          lines.push('');
          lines.push('| Label | Value | Default | Color | Description |');
          lines.push('|-------|-------|---------|-------|-------------|');
          sortByLabel(os.options, (opt) => opt.label || String(opt.value)).forEach((opt) => {
            lines.push(`| ${mdEscape(opt.label)} | ${opt.value} | ${opt.isDefault ? '✅' : '–'} | ${mdEscape(opt.color)} | ${mdEscape(opt.description)} |`);
          });
          lines.push('');
        });
      }
    }

    lines.push('---');
    lines.push('');
  });

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Global Option Sets
// ---------------------------------------------------------------------------

/**
 * Documents all global option sets defined in the solution.
 *
 * @param optionSets - Array of global option set definitions
 * @returns Markdown string
 */
function generateOptionSetsSection(optionSets: OptionSetDefinition[]): string {
  if (optionSets.length === 0) return '';

  const lines: string[] = [heading(2, 'Global Option Sets'), '', ''];
  sortByLabel(optionSets, (os) => labelWithSchema(os.displayName || os.name, os.name)).forEach((os) => {
    lines.push(heading(3, labelWithSchema(os.displayName || os.name, os.name)));
    lines.push('');
    if (os.description) {
      lines.push(`> ${os.description}`);
      lines.push('');
    }
    lines.push('| Label | Value | Default | Color | Description |');
    lines.push('|-------|-------|---------|-------|-------------|');
    sortByLabel(os.options, (opt) => opt.label || String(opt.value)).forEach((opt) => {
      lines.push(`| ${mdEscape(opt.label || String(opt.value))} | ${opt.value} | ${opt.isDefault ? '✅' : '–'} | ${mdEscape(opt.color)} | ${mdEscape(opt.description)} |`);
    });
    lines.push('');
  });

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Forms & Views
// ---------------------------------------------------------------------------

/**
 * Documents all forms and views.
 *
 * @param solution - Parsed solution data
 * @returns Markdown string
 */
function generateFormsViewsSection(solution: ParsedSolution, entityMap: Map<string, string>): string {
  if (solution.forms.length === 0 && solution.views.length === 0) return '';

  const lines: string[] = [heading(2, 'Forms & Views'), ''];

  // Forms
  if (solution.forms.length > 0) {
    lines.push(heading(3, 'Forms'));
    lines.push('');
    lines.push('| Form Name | Entity | Form Type | Fields |');
    lines.push('|-----------|--------|-----------|--------|');

    sortByLabel(solution.forms, (form) => form.displayName || form.name).forEach((form) => {
      lines.push(
        `| ${mdEscape(form.displayName || form.name)} ` +
        `| ${entityDisplayLabel(form.entityLogicalName, entityMap)} ` +
        `| ${form.formType} ` +
        `| ${form.fields.length} |`,
      );
    });
    lines.push('');

    // Detailed form field lists
    sortByLabel(solution.forms, (form) => form.displayName || form.name).forEach((form) => {
      if (form.fields.length > 0) {
        lines.push(heading(4, `${form.displayName || form.name} (${entityDisplayLabel(form.entityLogicalName, entityMap)} — ${form.formType})`));
        lines.push('');
        lines.push('| Field (attribute) | Label | Required | Region | Tab | Section |');
        lines.push('|-------------------|-------|----------|--------|-----|---------|');
        form.fields.forEach((f) => {
          const region = f.location ? `${f.location[0].toUpperCase()}${f.location.slice(1)}` : 'Body';
          lines.push(
            `| \`${mdEscape(f.attributeName)}\` ` +
            `| ${mdEscape(f.label)} ` +
            `| ${f.required ? '✅' : ''} ` +
            `| ${region} ` +
            `| ${mdEscape(f.tabName)} ` +
            `| ${mdEscape(f.sectionName)} |`,
          );
        });
        lines.push('');
      }
    });
  }

  // Views
  if (solution.views.length > 0) {
    lines.push(heading(3, 'Views'));
    lines.push('');
    lines.push('| View Name | Entity | Type | Columns |');
    lines.push('|-----------|--------|------|---------|');

    sortByLabel(solution.views, (view) => view.displayName || view.name).forEach((view) => {
      lines.push(
        `| ${mdEscape(view.displayName || view.name)} ` +
        `| ${entityDisplayLabel(view.entityLogicalName, entityMap)} ` +
        `| ${view.viewType} ` +
        `| ${view.columns.join(', ') || '–'} |`,
      );
    });
    lines.push('');
  }

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Processes
// ---------------------------------------------------------------------------

/**
 * Generates the Processes & Automation section with summary and detailed
 * step tables plus per-process Mermaid flow diagrams when step structure exists.
 *
 * @param processes - Array of process definitions
 * @returns Markdown string
 */
function generateProcessesSection(
  processes: ProcessDefinition[],
  entityMap: Map<string, string>,
  connectionRefs: ConnectionReferenceDefinition[],
  envVars: EnvironmentVariableDefinition[],
): string {
  if (processes.length === 0) return '';

  const lines: string[] = [heading(2, 'Processes & Automation'), ''];
  const statusRank = (status: boolean | undefined): number => {
    if (status === true) return 0;
    if (status === false) return 1;
    return 2;
  };
  const sortedProcesses = [...processes].sort((a, b) => {
    const rankDiff = statusRank(a.isActivated) - statusRank(b.isActivated);
    if (rankDiff !== 0) return rankDiff;
    return processTitle(a).localeCompare(processTitle(b), undefined, { sensitivity: 'base' });
  });

  const refCandidates = connectionRefs.flatMap((ref) => [ref.name, ref.displayName]).filter((v): v is string => !!v);
  const envCandidates = envVars.flatMap((env) => [env.schemaName, env.displayName]).filter((v): v is string => !!v);
  const flowMatches = (flow: object | undefined, candidates: string[]) => {
    if (!flow || candidates.length === 0) return [];
    const serialized = JSON.stringify(flow).toLowerCase();
    return Array.from(new Set(candidates.filter((candidate) => serialized.includes(candidate.toLowerCase()))));
  };

  // Summary table
  lines.push('| Name | Category | Primary Table | Referenced Tables | Connectors | Connection References | Environment Variables | Trigger | Status |');
  lines.push('|------|----------|---------------|-------------------|------------|------------------------|-----------------------|---------|--------|');

  sortedProcesses.forEach((proc) => {
    const relatedTables = uniqueStrings([...(proc.relatedEntities ?? []), proc.primaryEntity].filter(Boolean) as string[]);
    const usedRefs = (proc.flowConnectionReferences?.length ?? 0) > 0 ? proc.flowConnectionReferences! : flowMatches(proc.flowDefinition as object | undefined, refCandidates);
    const usedEnvVars = (proc.flowEnvironmentVariables?.length ?? 0) > 0 ? proc.flowEnvironmentVariables! : flowMatches(proc.flowDefinition as object | undefined, envCandidates);
    lines.push(
      `| ${mdEscape(processTitle(proc))} ` +
      `| ${processCategoryLabel(proc.category)} ` +
      `| ${entityDisplayLabel(proc.primaryEntity, entityMap)} ` +
      `| ${relatedTables.length > 0 ? relatedTables.map((table) => entityDisplayLabel(table, entityMap)).join(', ') : '–'} ` +
      `| ${mdEscape(sortByLabel(proc.flowConnectors ?? [], (c) => c).join(', ')) || '–'} ` +
      `| ${mdEscape(sortByLabel(usedRefs, (r) => r).join(', ')) || '–'} ` +
      `| ${mdEscape(sortByLabel(usedEnvVars, (e) => e).join(', ')) || '–'} ` +
      `| ${mdEscape(proc.triggerType || proc.flowTrigger || '–')} ` +
      `| ${processStatusLabel(proc.isActivated, true)} |`,
    );
  });
  lines.push('');

  // Detail per process
  sortedProcesses.forEach((proc) => {
    const relatedTables = uniqueStrings([...(proc.relatedEntities ?? []), proc.primaryEntity].filter(Boolean) as string[]);
    const usedRefs = (proc.flowConnectionReferences?.length ?? 0) > 0 ? proc.flowConnectionReferences! : flowMatches(proc.flowDefinition as object | undefined, refCandidates);
    const usedEnvVars = (proc.flowEnvironmentVariables?.length ?? 0) > 0 ? proc.flowEnvironmentVariables! : flowMatches(proc.flowDefinition as object | undefined, envCandidates);

    lines.push(heading(3, processTitle(proc)));
    lines.push('');

    lines.push('| Property | Value |');
    lines.push('|----------|-------|');
    lines.push(`| **Category** | ${processCategoryLabel(proc.category)} |`);
    if (proc.primaryEntity) lines.push(`| **Primary Table** | ${entityDisplayLabel(proc.primaryEntity, entityMap)} |`);
    if (relatedTables.length > 0) lines.push(`| **Referenced Tables** | ${relatedTables.map((table) => entityDisplayLabel(table, entityMap)).join(', ')} |`);
    if (proc.triggerType)   lines.push(`| **Trigger** | ${proc.triggerType} |`);
    if (proc.flowTrigger)   lines.push(`| **Flow Trigger** | ${mdEscape(proc.flowTrigger)} |`);
    lines.push(`| **Status** | ${processStatusLabel(proc.isActivated)} |`);
    if (proc.runAs)         lines.push(`| **Run As** | ${proc.runAs} |`);
    if (proc.scope)         lines.push(`| **Scope** | ${proc.scope} |`);

    if (proc.flowConnectors && proc.flowConnectors.length > 0) {
      lines.push(`| **Connectors Used** | ${sortByLabel(proc.flowConnectors, (c) => c).map((c) => mdEscape(c)).join(', ')} |`);
    }
    if (usedRefs.length > 0) {
      lines.push(`| **Connection References Used** | ${sortByLabel(usedRefs, (r) => r).map((r) => mdEscape(r)).join(', ')} |`);
    }
    if (usedEnvVars.length > 0) {
      lines.push(`| **Environment Variables Used** | ${sortByLabel(usedEnvVars, (e) => e).map((e) => mdEscape(e)).join(', ')} |`);
    }
    lines.push('');

    if (proc.description) {
      lines.push(`> ${mdEscape(proc.description)}`);
      lines.push('');
    }

    const graphTargets = uniqueStrings([
      proc.primaryEntity,
      ...(proc.relatedEntities ?? []),
      ...usedRefs,
      ...usedEnvVars,
    ].filter(Boolean) as string[]);

    if (graphTargets.length > 0) {
      const graphNodeId = flowchartNodeId(`process_${proc.uniqueName || proc.name}`);
      const graphLines: string[] = [
        'flowchart LR',
        `  ${graphNodeId}["${mermaidLabel(graphLabel(proc.displayName || proc.name, proc.uniqueName))}"]`,
      ];
      const emittedNodes = new Set<string>([graphNodeId]);

      const addTarget = (kind: 'entity' | 'reference' | 'variable', targetName: string, label: string) => {
        const nodeId = flowchartNodeId(`${kind}_${targetName}`);
        if (!emittedNodes.has(nodeId)) {
          graphLines.push(`  ${nodeId}["${mermaidLabel(label)}"]`);
          emittedNodes.add(nodeId);
        }
        graphLines.push(`  ${graphNodeId} --> ${nodeId}`);
      };

      uniqueStrings([proc.primaryEntity, ...(proc.relatedEntities ?? [])].filter(Boolean) as string[])
        .sort((a, b) => entityDisplayLabel(a, entityMap).localeCompare(entityDisplayLabel(b, entityMap), undefined, { sensitivity: 'base' }))
        .forEach((tableName) => {
          addTarget('entity', tableName, entityGraphLabel(tableName, entityMap));
        });

      sortByLabel(usedRefs, (ref) => ref).forEach((ref) => {
        addTarget('reference', ref, `Connection Reference: ${ref}`);
      });

      sortByLabel(usedEnvVars, (env) => env).forEach((env) => {
        addTarget('variable', env, `Environment Variable: ${env}`);
      });

      lines.push(heading(4, 'Relationship Diagram'));
      lines.push('');
      lines.push('> This diagram shows the process and the solution components it depends on.');
      lines.push('');
      lines.push(mermaidBlock(graphLines.join('\n'), `Process dependencies for ${processTitle(proc)}`));
      lines.push('');
    }

    if (proc.steps.length > 0) {
      type ProcessFlowEdge = { from: string; to: string; relation?: string };
      const processNodeId = flowchartNodeId(`process_flow_${proc.uniqueName || proc.name}`);
      const nodeLabels = new Map<string, string>([[processNodeId, graphLabel(proc.displayName || proc.name, proc.uniqueName)]]);
      const stepNodeIds: string[] = [];
      const stepNodeSet = new Set<string>();
      const entityNodeSet = new Set<string>();
      const flowEdges: ProcessFlowEdge[] = [];
      const emittedFlowEdges = new Set<string>();

      const addEdge = (from: string, to: string, relation?: string) => {
        const key = `${from}|${to}|${relation || ''}`;
        if (emittedFlowEdges.has(key)) return;
        emittedFlowEdges.add(key);
        flowEdges.push({ from, to, relation });
      };

      const walkSteps = (steps: ProcessStep[], parentNodeId: string, pathPrefix: string) => {
        steps.forEach((step, index) => {
          const stepNodeId = flowchartNodeId(`step_${proc.uniqueName || proc.name}_${pathPrefix}_${index}_${step.id || step.name}`);
          const stepLabel = graphLabel(step.name || step.stepType, step.stepType);

          if (!nodeLabels.has(stepNodeId)) {
            nodeLabels.set(stepNodeId, stepLabel);
            stepNodeIds.push(stepNodeId);
            stepNodeSet.add(stepNodeId);
          }

          addEdge(parentNodeId, stepNodeId, parentNodeId === processNodeId ? 'starts' : 'then');

          sortByLabel(step.referencedEntities ?? [], (entity) => entity).forEach((entity) => {
            const entityNodeId = flowchartNodeId(`step_entity_${entity}`);
            if (!nodeLabels.has(entityNodeId)) {
              nodeLabels.set(entityNodeId, entityGraphLabel(entity, entityMap));
              entityNodeSet.add(entityNodeId);
            }
            addEdge(stepNodeId, entityNodeId, 'touches');
          });

          if (step.children && step.children.length > 0) {
            walkSteps(step.children, stepNodeId, `${pathPrefix}_${index}`);
          }
        });
      };

      walkSteps(proc.steps, processNodeId, 'root');

      const buildFlowDiagram = (includedSteps: Set<string>) => {
        const includeNodes = new Set<string>([processNodeId, ...includedSteps]);
        const includedEdges = flowEdges.filter((edge) => {
          if (edge.from === processNodeId && includedSteps.has(edge.to)) return true;
          if (includedSteps.has(edge.from) && includedSteps.has(edge.to)) return true;
          if (includedSteps.has(edge.from) && entityNodeSet.has(edge.to)) {
            includeNodes.add(edge.to);
            return true;
          }
          return false;
        });

        const diagramLines: string[] = ['flowchart TD'];
        diagramLines.push(`  ${processNodeId}["${mermaidLabel(nodeLabels.get(processNodeId))}"]`);

        stepNodeIds.filter((nodeId) => includeNodes.has(nodeId)).forEach((nodeId) => {
          diagramLines.push(`  ${nodeId}["${mermaidLabel(nodeLabels.get(nodeId))}"]`);
        });

        Array.from(entityNodeSet)
          .filter((nodeId) => includeNodes.has(nodeId))
          .sort()
          .forEach((nodeId) => {
            diagramLines.push(`  ${nodeId}["${mermaidLabel(nodeLabels.get(nodeId))}"]`);
          });

        includedEdges.forEach((edge) => {
          if (edge.relation) {
            diagramLines.push(`  ${edge.from} -->|${mermaidLabel(edge.relation)}| ${edge.to}`);
          } else {
            diagramLines.push(`  ${edge.from} --> ${edge.to}`);
          }
        });

        return diagramLines.join('\n');
      };

      lines.push(heading(4, 'Process Flow Diagram'));
      lines.push('');

      const MAX_STEPS_PER_DIAGRAM = 45;
      if (stepNodeIds.length <= MAX_STEPS_PER_DIAGRAM) {
        lines.push(mermaidBlock(buildFlowDiagram(new Set(stepNodeIds)), `Process flow for ${processTitle(proc)}`));
        lines.push('');
      } else {
        lines.push(`> Large process detected (${stepNodeIds.length} steps). Flow diagram is split for readability.`);
        lines.push('');
        for (let i = 0; i < stepNodeIds.length; i += MAX_STEPS_PER_DIAGRAM) {
          const part = (i / MAX_STEPS_PER_DIAGRAM) + 1;
          const chunk = new Set(stepNodeIds.slice(i, i + MAX_STEPS_PER_DIAGRAM));
          lines.push(heading(5, `Flow Segment ${part}`));
          lines.push('');
          lines.push(mermaidBlock(buildFlowDiagram(chunk), `Process flow segment ${part} for ${processTitle(proc)}`));
          lines.push('');
        }
      }
    }

    // Trigger attributes
    if (proc.triggerAttributes && proc.triggerAttributes.length > 0) {
      lines.push(`**Triggered on changes to:** ${proc.triggerAttributes.map((a) => `\`${a}\``).join(', ')}`);
      lines.push('');
    }

    // Step table (if any)
    if (proc.steps.length > 0) {
      const flattened = flattenProcessSteps(proc.steps);
      const showDescription = flattened.some(({ step }) => processStepDescription(step).length > 0);
      const showReferencedEntities = flattened.some(({ step }) => (step.referencedEntities?.length ?? 0) > 0);
      lines.push(heading(4, 'Steps'));
      lines.push('');
      let header = '| # | Step | Type |';
      let sep    = '|---|------|------|';
      if (showDescription) { header += ' Description |'; sep += '-------------|'; }
      if (showReferencedEntities) { header += ' Referenced Tables |'; sep += '-------------------|'; }
      lines.push(header);
      lines.push(sep);
      flattened.forEach(({ step, depth }, index) => {
        const depthPrefix = '&nbsp;'.repeat(depth * 2);
        const stepDescription = processStepDescription(step);
        let row = `| ${index + 1} ` +
          `| ${depthPrefix}${mdEscape(step.name)} ` +
          `| ${mdEscape(step.stepType)} |`;
        if (showDescription) row += ` ${mdEscape(stepDescription) || '–'} |`;
        if (showReferencedEntities) {
          const refs = (step.referencedEntities ?? []).map((e) => entityDisplayLabel(e, entityMap)).join(', ');
          row += ` ${refs || '–'} |`;
        }
        lines.push(row);
      });
      lines.push('');
    }

    lines.push('---');
    lines.push('');
  });

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Power Apps
// ---------------------------------------------------------------------------

/**
 * Documents all Power Apps included in the solution.
 *
 * @param apps - Array of app definitions
 * @returns Markdown string
 */
function generateAppsSection(apps: AppDefinition[], entityMap: Map<string, string>): string {
  const documentedApps = apps.filter(isDocumentedApp);
  if (documentedApps.length === 0) return '';

  const lines: string[] = [heading(2, 'Power Apps'), ''];

  const appDetailBullets = (app: AppDefinition): string => {
    const details: string[] = [];
    if (app.appType === AppType.ModelDriven) {
      const areaCount = app.siteMap?.length ?? 0;
      const groupCount = (app.siteMap ?? []).reduce((sum, area) => sum + (area.groups?.length ?? 0), 0);
      const subAreaCount = (app.siteMap ?? []).reduce(
        (sum, area) => sum + (area.groups ?? []).reduce((groupSum, group) => groupSum + (group.subAreas?.length ?? 0), 0),
        0,
      );
      if (areaCount > 0) details.push(`Areas ${areaCount}`);
      if (groupCount > 0) details.push(`Groups ${groupCount}`);
      if (subAreaCount > 0) details.push(`Subareas ${subAreaCount}`);
    }
    const screenCount = app.canvasInsights?.screenCount ?? app.canvasInsights?.screenNames?.length;
    if (screenCount) details.push(`Screens ${screenCount}`);
    if (app.canvasInsights?.controlCount) details.push(`Controls ${app.canvasInsights.controlCount}`);
    if (app.canvasInsights?.dataSourceCount) details.push(`Data Sources ${app.canvasInsights.dataSourceCount}`);
    if (app.canvasInsights?.variableCount) details.push(`Variables ${app.canvasInsights.variableCount}`);
    if (app.canvasInsights?.resourceCount) details.push(`Resources ${app.canvasInsights.resourceCount}`);
    return tableBulletedList(details);
  };

  lines.push('| App Name | Unique Name | Type | Tables | Connectors | Details | Version | Status |');
  lines.push('|----------|-------------|------|--------|------------|---------|---------|--------|');
  sortByLabel(documentedApps, (app) => appTitle(app)).forEach((app) => {
    const tableLabels = sortByLabel(app.entities ?? [], (name) => name).map((name) => entityDisplayLabel(name, entityMap));
    const connectorLabels = sortByLabel(app.connectors ?? [], (name) => name);
    lines.push(
      `| ${mdEscape(appTitle(app))} ` +
      `| \`${mdEscape(app.uniqueName)}\` ` +
      `| ${appTypeLabel(app.appType)} ` +
      `| ${tableBulletedList(tableLabels)} ` +
      `| ${tableBulletedList(connectorLabels)} ` +
      `| ${appDetailBullets(app)} ` +
      `| ${app.version || '–'} ` +
      `| ${app.isEnabled !== false ? '✅ Enabled' : '⛔ Disabled'} |`,
    );
  });
  lines.push('');

  const modelDrivenApps = sortByLabel(
    documentedApps.filter((app) => app.appType === AppType.ModelDriven),
    (app) => appTitle(app),
  );
  const canvasLikeApps = sortByLabel(
    documentedApps.filter((app) => app.appType === AppType.Canvas || app.appType === AppType.CustomPage),
    (app) => appTitle(app),
  );

  if (modelDrivenApps.length > 0) {
    lines.push(heading(3, 'Site Map'));
    lines.push('');

    modelDrivenApps.forEach((app) => {
      lines.push(heading(4, appTitle(app)));
      lines.push('');

      const settings = app.siteMapSettings;
      if (settings) {
        lines.push('| Setting | Value |');
        lines.push('|---------|-------|');
        lines.push(`| Show Home | ${settings.showHome === undefined ? '–' : settings.showHome ? 'Yes' : 'No'} |`);
        lines.push(`| Show Pinned | ${settings.showPinned === undefined ? '–' : settings.showPinned ? 'Yes' : 'No'} |`);
        lines.push(`| Show Recents | ${settings.showRecents === undefined ? '–' : settings.showRecents ? 'Yes' : 'No'} |`);
        lines.push(`| Collapsible Groups | ${settings.enableCollapsibleGroups === undefined ? '–' : settings.enableCollapsibleGroups ? 'Yes' : 'No'} |`);
        lines.push('');
      }

      if ((app.siteMap?.length ?? 0) === 0) {
        lines.push('No sitemap structure was exported for this app.');
        lines.push('');
        return;
      }

      app.siteMap!.forEach((area) => {
        const areaTitle = mdEscape(area.title || area.id || 'Unnamed Area');
        lines.push(`- Area: ${areaTitle}`);

        if ((area.groups?.length ?? 0) === 0) {
          lines.push('  - Group: (none)');
          return;
        }

        area.groups.forEach((group) => {
          const groupTitle = mdEscape(group.title || group.id || 'Unnamed Group');
          lines.push(`  - Group: ${groupTitle}`);

          if ((group.subAreas?.length ?? 0) === 0) {
            lines.push('    - Subarea: (none)');
            return;
          }

          group.subAreas.forEach((subArea) => {
            const subAreaLabel = mdEscape(subArea.title || subArea.id || subArea.entity || subArea.url || 'Subarea');
            const details: string[] = [];
            if (subArea.entity) details.push(`Table ${entityDisplayLabel(subArea.entity, entityMap)}`);
            if (subArea.url) details.push(`URL ${mdEscape(subArea.url)}`);
            lines.push(`    - Subarea: ${subAreaLabel}${details.length > 0 ? ` (${details.join(' | ')})` : ''}`);
          });
        });
      });

      lines.push('');
    });
  }

  if (canvasLikeApps.length > 0) {
    lines.push(heading(3, 'Canvas App Insights'));
    lines.push('');

    canvasLikeApps.forEach((app) => {
      lines.push(heading(4, appTitle(app)));
      lines.push('');

      const insights = app.canvasInsights;
      if (!insights) {
        lines.push('No canvas insight metadata was exported for this app.');
        lines.push('');
        return;
      }

      lines.push('| Metric | Value |');
      lines.push('|--------|-------|');
      lines.push(`| Screens | ${insights.screenCount ?? insights.screenNames?.length ?? '–'} |`);
      lines.push(`| Controls | ${insights.controlCount ?? '–'} |`);
      lines.push(`| Data Sources | ${insights.dataSourceCount ?? '–'} |`);
      lines.push(`| Variables | ${insights.variableCount ?? '–'} |`);
      lines.push(`| Resources | ${insights.resourceCount ?? '–'} |`);
      lines.push('');

      if ((insights.screenNames?.length ?? 0) > 0) {
        lines.push(`- Screens: ${insights.screenNames!.map((name) => mdEscape(name)).join(', ')}`);
      }
      if ((insights.dataSources?.length ?? 0) > 0) {
        lines.push(`- Data Sources: ${insights.dataSources!.map((name) => mdEscape(name)).join(', ')}`);
      }
      if ((insights.variables?.length ?? 0) > 0) {
        lines.push(`- Variables: ${insights.variables!.map((name) => mdEscape(name)).join(', ')}`);
      }
      if ((insights.resources?.length ?? 0) > 0) {
        lines.push(`- Resources: ${insights.resources!.map((name) => mdEscape(name)).join(', ')}`);
      }

      if ((insights.screens?.length ?? 0) > 0) {
        lines.push('');
        lines.push('| Screen | Controls |');
        lines.push('|--------|----------|');
        insights.screens!.forEach((screen) => {
          const controls = screen.controls.length > 0
            ? screen.controls.map((control) => mdEscape(control)).join(', ')
            : '–';
          lines.push(`| ${mdEscape(screen.name)} | ${controls} |`);
        });
      }

      if ((insights.navigation?.length ?? 0) > 0) {
        lines.push('');
        lines.push('| From Screen | To Screen |');
        lines.push('|-------------|-----------|');
        insights.navigation!.forEach((link) => {
          lines.push(`| ${mdEscape(link.from)} | ${mdEscape(link.to)} |`);
        });
      }

      lines.push('');
    });
  }

  return lines.join('\n');
}

function generateAgentsSection(agents: AgentDefinition[]): string {
  if (agents.length === 0) return '';

  const lines: string[] = [heading(2, 'Copilot Studio Agents'), ''];
  lines.push('| Agent | Type | Language | Trigger/Channel | Connectors | Source |');
  lines.push('|-------|------|----------|-----------------|------------|--------|');

  sortByLabel(agents, (agent) => agent.displayName || agent.name).forEach((agent) => {
    lines.push(
      `| ${mdEscape(agent.displayName || agent.name)} ` +
      `| ${mdEscape(agent.agentType) || '–'} ` +
      `| ${mdEscape(agent.language) || '–'} ` +
      `| ${mdEscape(agent.trigger) || '–'} ` +
      `| ${tableBulletedList(sortByLabel(agent.connectors ?? [], (item) => item))} ` +
      `| \`${mdEscape(agent.sourcePath)}\` |`,
    );
  });

  lines.push('');
  return lines.join('\n');
}

function generateAIModelsSection(aiModels: AIModelDefinition[]): string {
  if (aiModels.length === 0) return '';

  const lines: string[] = [heading(2, 'AI Models'), ''];
  lines.push('| Model | Type | Provider | Version | Endpoint/Deployment | Source |');
  lines.push('|-------|------|----------|---------|---------------------|--------|');

  sortByLabel(aiModels, (model) => model.displayName || model.name).forEach((model) => {
    lines.push(
      `| ${mdEscape(model.displayName || model.name)} ` +
      `| ${mdEscape(model.modelType) || '–'} ` +
      `| ${mdEscape(model.provider) || '–'} ` +
      `| ${mdEscape(model.version) || '–'} ` +
      `| ${mdEscape(model.endpoint) || '–'} ` +
      `| \`${mdEscape(model.sourcePath)}\` |`,
    );
  });

  lines.push('');
  return lines.join('\n');
}

function generateDesktopFlowsSection(desktopFlows: DesktopFlowDefinition[]): string {
  if (desktopFlows.length === 0) return '';

  const lines: string[] = [heading(2, 'Desktop Flows'), ''];
  lines.push('| Desktop Flow | Folder | Status | Steps | Connectors | Source |');
  lines.push('|--------------|--------|--------|-------|------------|--------|');

  sortByLabel(desktopFlows, (flow) => flow.displayName || flow.name).forEach((flow) => {
    const status = flow.isEnabled === undefined ? '–' : flow.isEnabled ? '✅ Enabled' : '⛔ Disabled';
    lines.push(
      `| ${mdEscape(flow.displayName || flow.name)} ` +
      `| ${mdEscape(flow.folder) || '–'} ` +
      `| ${status} ` +
      `| ${flow.stepCount ?? '–'} ` +
      `| ${tableBulletedList(sortByLabel(flow.connectors ?? [], (item) => item))} ` +
      `| \`${mdEscape(flow.sourcePath)}\` |`,
    );
  });

  lines.push('');
  return lines.join('\n');
}

function generateDataflowsSection(dataflows: DataflowDefinition[]): string {
  if (dataflows.length === 0) return '';

  const lines: string[] = [heading(2, 'Dataflows'), ''];
  lines.push('| Dataflow | Refresh | Connectors | Source |');
  lines.push('|----------|---------|------------|--------|');

  sortByLabel(dataflows, (flow) => flow.displayName || flow.name).forEach((flow) => {
    lines.push(
      `| ${mdEscape(flow.displayName || flow.name)} ` +
      `| ${mdEscape(flow.refreshMode) || '–'} ` +
      `| ${tableBulletedList(sortByLabel(flow.connectors ?? [], (item) => item))} ` +
      `| \`${mdEscape(flow.sourcePath)}\` |`,
    );
  });

  lines.push('');
  return lines.join('\n');
}

function generateCustomApisSection(customApis: CustomAPIDefinition[]): string {
  if (customApis.length === 0) return '';

  const lines: string[] = [heading(2, 'Custom APIs'), ''];
  lines.push('| Custom API | Bound Table | Function | Source |');
  lines.push('|------------|-------------|----------|--------|');

  sortByLabel(customApis, (api) => api.displayName || api.name).forEach((api) => {
    lines.push(
      `| ${mdEscape(api.displayName || api.name)} ` +
      `| ${mdEscape(api.boundEntityLogicalName) || '–'} ` +
      `| ${api.isFunction === undefined ? '–' : api.isFunction ? 'Yes' : 'No'} ` +
      `| \`${mdEscape(api.sourcePath)}\` |`,
    );
  });

  lines.push('');
  return lines.join('\n');
}

function generateOfflineProfilesSection(offlineProfiles: OfflineProfileDefinition[]): string {
  if (offlineProfiles.length === 0) return '';

  const lines: string[] = [heading(2, 'Offline Profiles'), ''];
  lines.push('| Profile | Type | Tables | Source |');
  lines.push('|---------|------|--------|--------|');

  sortByLabel(offlineProfiles, (profile) => profile.displayName || profile.name).forEach((profile) => {
    lines.push(
      `| ${mdEscape(profile.displayName || profile.name)} ` +
      `| ${mdEscape(profile.profileType) || '–'} ` +
      `| ${tableBulletedList(sortByLabel(profile.entities ?? [], (item) => item))} ` +
      `| \`${mdEscape(profile.sourcePath)}\` |`,
    );
  });

  lines.push('');
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Web Resources
// ---------------------------------------------------------------------------

/**
 * Documents all web resources, categorised by type.
 *
 * @param webResources - Array of web resource definitions
 * @returns Markdown string
 */
function generateWebResourcesSection(webResources: WebResourceDefinition[]): string {
  if (webResources.length === 0) return '';

  const lines: string[] = [heading(2, 'Web Resources'), ''];

  // Group by type
  const grouped = new Map<WebResourceType, WebResourceDefinition[]>();
  webResources.forEach((wr) => {
    if (!grouped.has(wr.resourceType)) grouped.set(wr.resourceType, []);
    grouped.get(wr.resourceType)!.push(wr);
  });

  // Summary
  lines.push('| Type | Count |');
  lines.push('|------|-------|');
  grouped.forEach((items, type) => {
    lines.push(`| ${type} | ${items.length} |`);
  });
  lines.push('');

  // Detail table
  lines.push('| Name | Logical/Schema Name | Type | Enabled For Mobile | Offline |');
  lines.push('|------|---------------------|------|--------------------|---------|');
  webResources.forEach((wr) => {
    const mobileLabel = wr.enabledForMobile === undefined ? '–' : wr.enabledForMobile ? '✅ Yes' : 'No';
    const offlineLabel = wr.availableOffline === undefined ? '–' : wr.availableOffline ? '✅ Yes' : 'No';
    lines.push(
      `| ${mdEscape(wr.displayName || wr.name)} ` +
      `| \`${mdEscape(wr.schemaName || wr.name)}\` ` +
      `| ${wr.resourceType} ` +
      `| ${mobileLabel} ` +
      `| ${offlineLabel} |`,
    );
  });
  lines.push('');

  // Expand JavaScript content (first 3kb) for documentation purposes
  const jsResources = webResources.filter(
    (wr) => (wr.resourceType === WebResourceType.JavaScript || wr.resourceType === WebResourceType.TypeScript)
      && wr.content,
  );
  if (jsResources.length > 0) {
    lines.push(heading(3, 'JavaScript / TypeScript Resources'));
    lines.push('');
    jsResources.forEach((wr) => {
      lines.push(heading(4, labelWithSchema(wr.displayName || wr.name, wr.schemaName || wr.name)));
      lines.push('');
      const snippet = (wr.content ?? '').substring(0, 3000);
      const lang    = wr.resourceType === WebResourceType.TypeScript ? 'typescript' : 'javascript';
      lines.push(`\`\`\`${lang}`);
      lines.push(snippet);
      if ((wr.content?.length ?? 0) > 3000) {
        lines.push(`// ... (${formatBytes(wr.contentLength)} total — truncated for documentation)`);
      }
      lines.push('```');
      lines.push('');
    });
  }

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Security Roles
// ---------------------------------------------------------------------------

/**
 * Documents security roles and column-level security profiles.
 *
 * @param roles   - Security role definitions
 * @param profiles - Column-level security profile definitions
 * @returns Markdown string
 */
function generateSecuritySection(
  roles: SecurityRoleDefinition[],
  profiles: FieldSecurityProfileDefinition[],
  entityMap: Map<string, string>,
  entities: EntityDefinition[],
  attributeDisplayMap: Map<string, string>,
  securityRoleFilters: DocumentationSecurityRoleFilters,
): string {
  if (roles.length === 0 && profiles.length === 0) return '';

  const lines: string[] = [];
  const roleMatrices = buildSecurityRolePrivilegeMatrices(roles, entities, securityRoleFilters);

  if (roleMatrices.length > 0) {
    lines.push(heading(2, 'Security Roles'));
    lines.push('');
    lines.push('| Role Name | Privileges |');
    lines.push('|-----------|-----------|');
    roleMatrices.forEach(({ role, privilegeCount }) => {
      lines.push(
        `| ${mdEscape(role.displayName || role.name)} ` +
        `| ${privilegeCount} |`,
      );
    });
    lines.push('');

    roleMatrices.forEach(({ role, matrix }) => {
      lines.push(heading(3, role.displayName || role.name));
      lines.push('');

      lines.push(heading(4, 'Privilege Matrix'));
      lines.push('');
      lines.push('| Table | Logical Name | Create | Read | Write | Delete | Append | Append To | Assign | Share | Unshare |');
      lines.push('|-------|--------------|--------|------|-------|--------|--------|-----------|--------|-------|---------|');

      const matrixRows = sortByLabel(Array.from(matrix.entries()), ([table]) => table);
      if (matrixRows.length === 0) {
        lines.push('| _No table privileges matched the active filters for this role._ | – | – | – | – | – | – | – | – | – | – |');
      }

      matrixRows.forEach(([table, ops]) => {
        const tableDisplayName = entityMap.get(table.toLowerCase()) || humanizeEntityName(table);
        lines.push(
          `| ${mdEscape(tableDisplayName)} ` +
          `| \`${mdEscape(table)}\` ` +
          `| ${accessDepthBadge(ops.Create)} ` +
          `| ${accessDepthBadge(ops.Read)} ` +
          `| ${accessDepthBadge(ops.Write)} ` +
          `| ${accessDepthBadge(ops.Delete)} ` +
          `| ${accessDepthBadge(ops.Append)} ` +
          `| ${accessDepthBadge(ops.AppendTo)} ` +
          `| ${accessDepthBadge(ops.Assign)} ` +
          `| ${accessDepthBadge(ops.Share)} ` +
          `| ${accessDepthBadge(ops.Unshare)} |`,
        );
      });
      lines.push('');
    });
  }

  if (profiles.length > 0) {
    lines.push(heading(2, 'Column Level Security Profiles'));
    lines.push('');

    lines.push('| Display Name | Logical Name | Columns Secured |');
    lines.push('|--------------|--------------|----------------|');
    sortByLabel(profiles, (profile) => profile.displayName || profile.name).forEach((profile) => {
      lines.push(
        `| ${mdEscape(profile.displayName || profile.name)} ` +
        `| \`${mdEscape(profile.name)}\` ` +
        `| ${profile.permissions.length} |`,
      );
    });
    lines.push('');

    sortByLabel(profiles, (profile) => profile.displayName || profile.name).forEach((profile) => {
      lines.push(heading(3, labelWithSchema(profile.displayName || profile.name, profile.name)));
      lines.push('');
      if (profile.permissions.length > 0) {
        lines.push('| Display Name | Logical/Schema Name | Read | Update | Create |');
        lines.push('|--------------|---------------------|------|--------|--------|');
        sortByLabel(profile.permissions, (perm) => perm.attributeName).forEach((perm) => {
          const displayName = resolveAttributeDisplayName(perm.attributeName, attributeDisplayMap);
          lines.push(
            `| ${mdEscape(displayName || '–')} ` +
            `| \`${mdEscape(perm.attributeName)}\` ` +
            `| ${allowedBadge(perm.canRead)} ` +
            `| ${allowedBadge(perm.canUpdate)} ` +
            `| ${allowedBadge(perm.canCreate)} |`,
          );
        });
        lines.push('');
      }
    });
  }

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Connection References & Environment Variables
// ---------------------------------------------------------------------------

/**
 * Documents connection references and environment variable definitions.
 *
 * @param connectionRefs - Connection reference definitions
 * @param envVars        - Environment variable definitions
 * @returns Markdown string
 */
function generateIntegrationSection(
  connectionRefs: ConnectionReferenceDefinition[],
  envVars: EnvironmentVariableDefinition[],
  emailTemplates: EmailTemplateDefinition[],
): string {
  if (connectionRefs.length === 0 && envVars.length === 0 && emailTemplates.length === 0) return '';

  const lines: string[] = [];

  if (connectionRefs.length > 0) {
    lines.push(heading(2, 'Connection References'));
    lines.push('');
    lines.push('| Display Name | Logical Name | Connector (Friendly) | Connector ID | Connection ID |');
    lines.push('|--------------|--------------|----------------------|--------------|---------------|');
    sortByLabel(connectionRefs, (cr) => cr.displayName || cr.name).forEach((cr) => {
      lines.push(
        `| ${mdEscape(cr.displayName || cr.name)} ` +
        `| \`${mdEscape(cr.name)}\` ` +
        `| ${mdEscape(cr.connectorDisplayName || '–')} ` +
        `| \`${mdEscape(cr.connectorId)}\` ` +
        `| ${cr.connectionId ? `\`${mdEscape(cr.connectionId)}\`` : '–'} |`,
      );
    });
    lines.push('');

    lines.push(heading(2, 'Connections'));
    lines.push('');
    lines.push('| Display Name | Connection ID | Connector | Reference Logical Name |');
    lines.push('|--------------|---------------|-----------|-------------------------|');
    sortByLabel(connectionRefs, (cr) => cr.displayName || cr.name).forEach((cr) => {
      lines.push(
        `| ${mdEscape(cr.displayName || cr.name)} ` +
        `| ${cr.connectionId ? `\`${mdEscape(cr.connectionId)}\`` : '–'} ` +
        `| ${mdEscape(cr.connectorDisplayName || cr.connectorId || '–')} ` +
        `| \`${mdEscape(cr.name)}\` |`,
      );
    });
    lines.push('');
  }

  if (envVars.length > 0) {
    lines.push(heading(2, 'Environment Variables'));
    lines.push('');
    lines.push('| Display Name | Schema Name | Type | Has Value | Current Value | Default Value |');
    lines.push('|--------------|-------------|------|-----------|---------------|---------------|');
    sortByLabel(envVars, (ev) => ev.displayName || ev.schemaName).forEach((ev) => {
      const hasAnyValue = !!(ev.hasCurrentValue || (ev.currentValue && ev.currentValue.trim()) || (ev.defaultValue && ev.defaultValue.trim()));
      lines.push(
        `| ${mdEscape(ev.displayName || ev.schemaName)} ` +
        `| \`${mdEscape(ev.schemaName)}\` ` +
        `| ${ev.type} ` +
        `| ${hasAnyValue ? '✅' : '⚠️ Not set'} ` +
        `| ${ev.hasCurrentValue ? mdEscape(ev.currentValue) || 'Set' : '–'} ` +
        `| ${mdEscape(ev.defaultValue) || '–'} |`,
      );
    });
    lines.push('');
  }

  if (emailTemplates.length > 0) {
    lines.push(heading(2, 'Email Templates'));
    lines.push('');
    lines.push('| Display Name | Logical Name | Subject | Related Table | Type | Language |');
    lines.push('|--------------|--------------|---------|---------------|------|----------|');
    sortByLabel(emailTemplates, (template) => template.displayName || template.name).forEach((template) => {
      lines.push(
        `| ${mdEscape(template.displayName || template.name)} ` +
        `| \`${mdEscape(template.name)}\` ` +
        `| ${mdEscape(template.subject) || '–'} ` +
        `| ${template.entityLogicalName ? `\`${mdEscape(template.entityLogicalName)}\`` : '–'} ` +
        `| ${mdEscape(template.templateType) || '–'} ` +
        `| ${mdEscape(template.languageCode) || '–'} |`,
      );
    });
    lines.push('');
  }

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Reports & Dashboards
// ---------------------------------------------------------------------------

/**
 * Documents reports and dashboards.
 *
 * @param reports    - Report definitions
 * @param dashboards - Dashboard definitions
 * @returns Markdown string
 */
function generateReportsSection(
  reports: ReportDefinition[],
  dashboards: DashboardDefinition[],
): string {
  if (reports.length === 0 && dashboards.length === 0) return '';

  const lines: string[] = [heading(2, 'Reports & Dashboards'), ''];

  if (reports.length > 0) {
    lines.push(heading(3, 'Reports'));
    lines.push('');
    lines.push('| Report Name | Category | File |');
    lines.push('|-------------|----------|------|');
    reports.forEach((rpt) => {
      lines.push(
        `| ${mdEscape(rpt.displayName || rpt.name)} ` +
        `| ${rpt.category || '–'} ` +
        `| ${rpt.fileName || '–'} |`,
      );
    });
    lines.push('');
  }

  if (dashboards.length > 0) {
    lines.push(heading(3, 'Dashboards'));
    lines.push('');
    lines.push('| Dashboard Name | Type | Components |');
    lines.push('|----------------|------|-----------|');
    dashboards.forEach((db) => {
      lines.push(
        `| ${mdEscape(db.displayName || db.name)} ` +
        `| ${db.dashboardType || 'Standard'} ` +
        `| ${db.components.length} |`,
      );
    });
    lines.push('');
  }

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Plugin Assemblies
// ---------------------------------------------------------------------------

/**
 * Documents plugin assemblies and their registered steps.
 *
 * @param assemblies - Plugin assembly definitions
 * @returns Markdown string
 */
function generatePluginsSection(assemblies: PluginAssemblyDefinition[]): string {
  if (assemblies.length === 0) return '';

  const lines: string[] = [heading(2, 'Plugin Assemblies & Steps'), ''];

  // Tabular detail per assembly
  assemblies.forEach((asm) => {
    lines.push(heading(3, labelWithSchema(asm.displayName || asm.name, asm.assemblyName)));
    lines.push('');

    lines.push('| Property | Value |');
    lines.push('|----------|-------|');
    lines.push(`| **Assembly Name** | \`${asm.assemblyName}\` |`);
    if (asm.version)        lines.push(`| **Version** | ${asm.version} |`);
    if (asm.culture)        lines.push(`| **Culture** | ${asm.culture} |`);
    if (asm.publicKeyToken) lines.push(`| **Public Key Token** | \`${asm.publicKeyToken}\` |`);
    lines.push(`| **Isolation Mode** | ${asm.isolationMode === 2 ? 'Sandbox' : 'None'} |`);
    if (asm.sourceType)     lines.push(`| **Source** | ${asm.sourceType} |`);
    lines.push('');

    if (asm.steps.length > 0) {
      lines.push(heading(4, 'Registered Steps'));
      lines.push('');
      lines.push('| Step Name | Message | Entity | Stage | Mode | Filtering Attributes |');
      lines.push('|-----------|---------|--------|-------|------|----------------------|');

      asm.steps.forEach((step: PluginStepDefinition) => {
        lines.push(
          `| ${mdEscape(step.name)} ` +
          `| ${mdEscape(step.message)} ` +
          `| \`${step.primaryEntity || '–'}\` ` +
          `| ${stageLabel(step.stage)} ` +
          `| ${modeLabel(step.mode)} ` +
          `| ${mdEscape(step.filteringAttributes) || '–'} |`,
        );
      });
      lines.push('');
    }

    lines.push('---');
    lines.push('');
  });

  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Parse warnings
// ---------------------------------------------------------------------------

/**
 * Appends a warnings section if there were any non-fatal parse issues.
 *
 * @param warnings - Array of warning strings from the parser
 * @returns Markdown string or empty string
 */
function generateWarningsSection(warnings: string[]): string {
  if (warnings.length === 0) return '';
  const lines: string[] = [
    heading(2, '⚠️ Parse Warnings'),
    '',
    '_The following non-fatal issues were encountered during parsing:_',
    '',
  ];
  warnings.forEach((w) => lines.push(`- ${mdEscape(w)}`));
  lines.push('');
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Generates complete Markdown documentation for a parsed Power Platform solution.
 * This is the primary public function of this module.
 *
 * @param solution - A fully-populated {@link ParsedSolution} from the parser
 * @returns A single Markdown string representing all documentation sections
 */
export function generateMarkdown(
  solution: ParsedSolution,
  options: MarkdownGenerationOptions = {},
): string {
  return renderMarkdownDocument(buildSolutionOutputSections(solution, options), { addBackToTopLinks: true });
}

function splitMarkdownSections(markdown: string): string[] {
  const lines = markdown.split('\n');
  const sections: string[] = [];
  let current: string[] = [];
  let inCodeFence = false;

  lines.forEach((line) => {
    if (line.startsWith('```')) {
      inCodeFence = !inCodeFence;
    }

    if (!inCodeFence && /^##\s+/.test(line) && current.length > 0) {
      sections.push(current.join('\n').trimEnd());
      current = [line];
      return;
    }

    current.push(line);
  });

  if (current.length > 0) sections.push(current.join('\n').trimEnd());
  return sections;
}

function removeMermaidBlocks(markdown: string): string {
  const lines = markdown.split('\n');
  const output: string[] = [];

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    if (!line.startsWith('```mermaid')) {
      output.push(line);
      continue;
    }

    // Remove diagram-specific heading directly above a Mermaid block.
    let scan = output.length - 1;
    while (scan >= 0 && output[scan].trim() === '') scan -= 1;
    if (scan >= 0 && /^#{3,6}\s+.*diagram/i.test(output[scan])) {
      output.splice(scan, output.length - scan);
      while (output.length > 0 && output[output.length - 1].trim() === '') {
        output.pop();
      }
      output.push('');
    }

    i += 1;
    while (i < lines.length && !lines[i].startsWith('```')) i += 1;
  }

  return output.join('\n').replace(/\n{3,}/g, '\n\n').trimEnd();
}

function insertCompanionNotice(markdown: string): string {
  const lines = markdown.split('\n');
  const tocIndex = lines.findIndex((line) => /^##\s+Table of Contents\s*$/.test(line));
  if (tocIndex === -1) {
    return [
      '> Diagram content has been moved to a companion diagrams document.',
      '',
      markdown,
    ].join('\n').trimEnd();
  }

  let insertAt = tocIndex + 1;
  while (insertAt < lines.length && lines[insertAt].trim() !== '') insertAt += 1;
  while (insertAt < lines.length && lines[insertAt].trim() === '') insertAt += 1;

  const notice = [
    '> Diagram content has been moved to a companion diagrams document generated alongside this file.',
    '',
  ];

  lines.splice(insertAt, 0, ...notice);
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trimEnd();
}

function appendBackToTopLinksForRawMarkdown(markdown: string): string {
  const lines = markdown.split('\n');
  const sectionStarts: number[] = [];
  let inCodeFence = false;

  lines.forEach((line, index) => {
    if (line.startsWith('```')) {
      inCodeFence = !inCodeFence;
      return;
    }
    if (!inCodeFence && /^##\s+/.test(line)) {
      sectionStarts.push(index);
    }
  });

  const inserts: Array<{ index: number }> = [];

  sectionStarts.forEach((start, idx) => {
    const title = lines[start].replace(/^##\s+/, '').trim();
    if (title.toLowerCase() === 'table of contents') return;

    const nextStart = idx < sectionStarts.length - 1 ? sectionStarts[idx + 1] : lines.length;
    const sectionText = lines.slice(start, nextStart).join('\n');
    if (sectionText.includes('[Back to Top](#table-of-contents)')) return;

    let insertAt = nextStart;
    while (insertAt > start + 1 && lines[insertAt - 1].trim() === '') {
      insertAt -= 1;
    }

    inserts.push({ index: insertAt });
  });

  inserts
    .sort((a, b) => b.index - a.index)
    .forEach(({ index }) => {
      lines.splice(index, 0, '', '[Back to Top](#table-of-contents)', '');
    });

  return lines.join('\n');
}

export function splitMarkdownForDiagramCompanion(markdown: string): { mainMarkdown: string; companionMarkdown: string } {
  const hasDiagrams = /```mermaid[\s\S]*?```/m.test(markdown);
  if (!hasDiagrams) {
    return { mainMarkdown: markdown, companionMarkdown: '' };
  }

  const sections = splitMarkdownSections(markdown);
  const headerSection = sections.find((section) => /^#\s+/.test(section.trimStart())) ?? '';
  const diagramSections = sections.filter((section) => /```mermaid[\s\S]*?```/m.test(section));

  const companionLines: string[] = [];
  if (headerSection) {
    companionLines.push(headerSection, '');
  }
  companionLines.push(
    '## Table of Contents',
    '',
  );

  diagramSections.forEach((section) => {
    const titleMatch = section.match(/^##\s+(.+)$/m);
    const title = titleMatch?.[1]?.trim() || 'Diagrams';
    companionLines.push(`- [${title}](${headingAnchor(title)})`);
  });

  companionLines.push('');
  companionLines.push('> Companion diagrams document containing Mermaid visuals extracted from the main documentation.', '');
  diagramSections.forEach((section) => {
    companionLines.push(section.trimEnd(), '');
  });

  const companionMarkdown = appendBackToTopLinksForRawMarkdown(companionLines.join('\n').replace(/\n{3,}/g, '\n\n').trimEnd());
  const mainWithoutDiagrams = removeMermaidBlocks(markdown);
  const mainMarkdown = appendBackToTopLinksForRawMarkdown(insertCompanionNotice(mainWithoutDiagrams));

  return { mainMarkdown, companionMarkdown };
}

function buildSolutionOutputSections(
  solution: ParsedSolution,
  options: MarkdownGenerationOptions,
): OutputSection[] {
  const documentationSettings = normalizeDocumentationSettings(options.documentationSettings);
  const includeDetailedSections = documentationSettings.detailLevel === 'detailed';
  const entityMap = buildEntityDisplayMap(solution.entities);
  const attributeDisplayMap = buildAttributeDisplayMap(solution.entities);

  return [
    createOutputSection('header', generateHeader(solution, options.documentContext)),
    createOutputSection('table-of-contents', generateTableOfContents(solution, documentationSettings)),
    createOutputSection('dependencies', generateDependenciesSection(solution.metadata.dependencies)),
    createOutputSection('component-inventory', generateComponentInventorySection(solution.metadata.componentInventory)),
    createOutputSection(
      'component-relationship-graph',
      includeDetailedSections ? generateComponentRelationshipGraphSection(solution, entityMap) : '',
    ),
    createOutputSection(
      'entity-relationship-diagram',
      includeDetailedSections ? generateERD(solution.entities, options, documentationSettings) : '',
    ),
    createOutputSection(
      'tables-columns',
      includeDetailedSections
        ? generateEntitiesSection(
          solution.entities,
          solution.optionSets,
          entityMap,
          solution.forms,
          solution.fieldSecurityProfiles,
          documentationSettings,
        )
        : '',
    ),
    createOutputSection('global-option-sets', includeDetailedSections ? generateOptionSetsSection(solution.optionSets) : ''),
    createOutputSection('forms-views', includeDetailedSections ? generateFormsViewsSection(solution, entityMap) : ''),
    createOutputSection(
      'processes-automation',
      documentationSettings.scope.flows
        ? generateProcessesSection(solution.processes, entityMap, solution.connectionReferences, solution.environmentVariables)
        : '',
    ),
    createOutputSection('power-apps', documentationSettings.scope.apps ? generateAppsSection(solution.apps, entityMap) : ''),
    createOutputSection('copilot-studio-agents', includeDetailedSections ? generateAgentsSection(solution.agents) : ''),
    createOutputSection('ai-models', includeDetailedSections ? generateAIModelsSection(solution.aiModels) : ''),
    createOutputSection('desktop-flows', includeDetailedSections ? generateDesktopFlowsSection(solution.desktopFlows) : ''),
    createOutputSection('dataflows', includeDetailedSections ? generateDataflowsSection(solution.dataflows ?? []) : ''),
    createOutputSection('custom-apis', includeDetailedSections ? generateCustomApisSection(solution.customApis ?? []) : ''),
    createOutputSection('offline-profiles', includeDetailedSections ? generateOfflineProfilesSection(solution.offlineProfiles ?? []) : ''),
    createOutputSection('web-resources', includeDetailedSections ? generateWebResourcesSection(solution.webResources) : ''),
    createOutputSection(
      'security-roles',
      documentationSettings.scope.security
        ? generateSecuritySection(
          solution.securityRoles,
          solution.fieldSecurityProfiles,
          entityMap,
          solution.entities,
          attributeDisplayMap,
          documentationSettings.securityRoleFilters,
        )
        : '',
    ),
    createOutputSection(
      'integration',
      documentationSettings.scope.integration
        ? generateIntegrationSection(solution.connectionReferences, solution.environmentVariables, solution.emailTemplates)
        : '',
    ),
    createOutputSection('reports-dashboards', documentationSettings.scope.reports ? generateReportsSection(solution.reports, solution.dashboards) : ''),
    createOutputSection('plugin-assemblies-steps', documentationSettings.scope.plugins ? generatePluginsSection(solution.pluginAssemblies) : ''),
    createOutputSection('warnings', generateWarningsSection(solution.warnings)),
  ];
}

export function buildMetadataGridRows(solution: ParsedSolution): OutputMetadataGridRow[] {
  const rows: OutputMetadataGridRow[] = [];

  solution.entities.forEach((entity) => {
    const entityLabel = entity.logicalName;
    rows.push({
      entity: entityLabel,
      value: `Entity: ${entity.displayName || entity.logicalName}`,
    });

    entity.attributes.forEach((attribute) => {
      rows.push({
        entity: entityLabel,
        attribute: attribute.name,
        value: `${attribute.type}${attribute.required ? ' (required)' : ''}`,
      });
    });

    entity.relationships.forEach((relationship) => {
      rows.push({
        entity: entityLabel,
        relationship: relationship.name,
        value: `${relationship.type}: ${relationship.referencingEntity} -> ${relationship.referencedEntity}`,
      });
    });
  });

  return rows;
}

export function generateMetadataGridMarkdown(solution: ParsedSolution): string {
  const rows = buildMetadataGridRows(solution);
  if (rows.length === 0) return '';

  return renderMarkdownDocument([
    createMetadataGridSection({
      id: 'solution-metadata-grid',
      title: 'Solution Metadata Grid',
      introMarkdown:
        '> Intermediate metadata-grid export containing entity, attribute, and relationship rows.',
      rows,
    }),
  ]);
}

/**
 * Generates a consolidated markdown summary across multiple parsed solutions
 * with shared inventory, ERD maps, and component sections.
 */
export function generateConsolidatedMarkdown(
  solutions: ParsedSolution[],
  options: MarkdownGenerationOptions = {},
): string {
  const items = solutions.filter((s) => !!s.metadata.uniqueName);
  if (items.length === 0) return '';
  if (items.length === 1) return generateMarkdown(items[0]);

  const consolidated = consolidateSolutions(items);
  const consolidatedEntityMap = buildEntityDisplayMap(consolidated.entities);
  const allConnectors = uniqueStrings([
    ...consolidated.connectionReferences.map((connection) => connection.connectorDisplayName || connection.connectorId),
    ...consolidated.processes.flatMap((process) => process.flowConnectors ?? []),
  ]);

  const lines: string[] = [];
  const contextSection = generateDocumentContextSection(options.documentContext);
  if (contextSection) {
    lines.push(contextSection);
  }

  lines.push(
    heading(1, 'Power Platform Solutions: Consolidated Summary'),
    '',
    `> Generated on: ${new Date().toLocaleString()}`,
    '',
    heading(2, 'Included Solutions'),
    '',
    '| Solution | Unique Name | Version | Tables | Processes | Apps | Other |',
    '|----------|-------------|---------|--------|-----------|------|-------|',
  );

  sortByLabel(items, (sol) => sol.metadata.displayName || sol.metadata.uniqueName).forEach((sol) => {
    const normalized = consolidateSolutions([sol]);
    const documentedApps = normalized.apps.filter(isDocumentedApp).length;
    const other =
      normalized.optionSets.length +
      normalized.forms.length +
      normalized.views.length +
      normalized.webResources.length +
      normalized.securityRoles.length +
      normalized.fieldSecurityProfiles.length +
      normalized.connectionReferences.length +
      normalized.environmentVariables.length +
      normalized.emailTemplates.length +
      normalized.reports.length +
      normalized.dashboards.length +
      normalized.pluginAssemblies.length;
    lines.push(
      `| ${mdEscape(sol.metadata.displayName)} ` +
      `| \`${mdEscape(sol.metadata.uniqueName)}\` ` +
      `| ${mdEscape(sol.metadata.version)} ` +
      `| ${normalized.entities.length} ` +
      `| ${normalized.processes.length} ` +
      `| ${documentedApps} ` +
      `| ${other} |`,
    );
  });
  lines.push('');

  lines.push(heading(2, 'Ecosystem Inventory'));
  lines.push('');
  lines.push('| Solutions | Tables | Processes | Apps | Connectors | Environment Variables | Email Templates | Reports | Dashboards | Plugins |');
  lines.push('|-----------|--------|-----------|------|------------|-----------------------|-----------------|---------|------------|---------|');
  lines.push(
    `| ${items.length} ` +
    `| ${consolidated.entities.length} ` +
    `| ${consolidated.processes.length} ` +
    `| ${consolidated.apps.filter(isDocumentedApp).length} ` +
    `| ${allConnectors.length} ` +
    `| ${consolidated.environmentVariables.length} ` +
    `| ${consolidated.emailTemplates.length} ` +
    `| ${consolidated.reports.length} ` +
    `| ${consolidated.dashboards.length} ` +
    `| ${consolidated.pluginAssemblies.length} |`,
  );
  lines.push('');

  const entityUsage = new Map<string, Set<string>>();
  items.forEach((sol) => {
    sol.entities.forEach((entity) => {
      const key = entity.logicalName;
      if (!entityUsage.has(key)) entityUsage.set(key, new Set<string>());
      entityUsage.get(key)!.add(sol.metadata.displayName || sol.metadata.uniqueName);
    });
  });
  const sharedEntities = Array.from(entityUsage.entries()).filter(([, sols]) => sols.size > 1);

  if (sharedEntities.length > 0) {
    lines.push(heading(2, 'Shared Dataverse Tables'));
    lines.push('');
    lines.push('| Table | Shared Across | Present In Solutions |');
    lines.push('|-------|---------------|----------------------|');
    sharedEntities
      .sort(([a], [b]) => a.localeCompare(b))
      .forEach(([table, sols]) => {
        lines.push(`| ${entityDisplayLabel(table, consolidatedEntityMap)} | ${sols.size} solutions | ${sortByLabel(Array.from(sols), (s) => s).map((s) => mdEscape(s)).join(', ')} |`);
      });
    lines.push('');
  }

  lines.push(heading(2, 'Apps Across Solutions'));
  lines.push('');
  lines.push('| App | Type | Tables | Connectors |');
  lines.push('|-----|------|--------|------------|');
  sortByLabel(consolidated.apps.filter(isDocumentedApp), (app) => appTitle(app)).forEach((app) => {
    lines.push(
      `| ${mdEscape(appTitle(app))} ` +
      `| ${appTypeLabel(app.appType)} ` +
      `| ${tableBulletedList(sortByLabel(app.entities ?? [], (e) => e).map((table) => entityDisplayLabel(table, consolidatedEntityMap)))} ` +
      `| ${tableBulletedList(sortByLabel(app.connectors ?? [], (c) => c))} |`,
    );
  });
  lines.push('');

  const processUsage = new Map<string, Set<string>>();
  items.forEach((sol) => {
    sol.processes.forEach((process) => {
      const key = stripTrailingGuid(process.uniqueName || process.name || process.displayName || '').toLowerCase();
      if (!key) return;
      if (!processUsage.has(key)) processUsage.set(key, new Set<string>());
      processUsage.get(key)!.add(sol.metadata.displayName || sol.metadata.uniqueName);
    });
  });

  lines.push(heading(2, 'Processes Across Solutions'));
  lines.push('');
  lines.push('| Process | Category | Solutions | No. of Solutions | Primary Table | Referenced Tables |');
  lines.push('|---------|----------|-----------|------------------|---------------|-------------------|');
  sortByLabel(consolidated.processes, (process) => processTitle(process)).forEach((process) => {
    const relatedTables = uniqueStrings([...(process.relatedEntities ?? []), process.primaryEntity]);
    const key = stripTrailingGuid(process.uniqueName || process.name || process.displayName || '').toLowerCase();
    const solutionsForProcess = sortByLabel(Array.from(processUsage.get(key) ?? []), (s) => s);
    lines.push(
      `| ${mdEscape(processTitle(process))} ` +
      `| ${processCategoryLabel(process.category)} ` +
      `| ${solutionsForProcess.length > 0 ? solutionsForProcess.map((s) => mdEscape(s)).join(', ') : '–'} ` +
      `| ${solutionsForProcess.length} ` +
      `| ${entityDisplayLabel(process.primaryEntity, consolidatedEntityMap)} ` +
      `| ${relatedTables.length > 0 ? sortByLabel(relatedTables, (table) => table).map((table) => entityDisplayLabel(table, consolidatedEntityMap)).join(', ') : '–'} |`,
    );
  });
  lines.push('');

  return renderMarkdownDocument([
    createOutputSection('consolidated-summary', lines.join('\n')),
  ], { addBackToTopLinks: true });
}
