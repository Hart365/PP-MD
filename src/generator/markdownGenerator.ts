/**
 * @file markdownGenerator.ts
 * @description Converts a {@link ParsedSolution} into comprehensive Markdown
 * documentation.  Each section of the output is produced by a dedicated
 * function which assembles the text and, where applicable, embeds a Mermaid
 * diagram code block.
 *
 * Mermaid diagram types used:
 *  - erDiagram          → Entity Relationship Diagram
 */

import type {
  ParsedSolution,
  EntityDefinition,
  ProcessDefinition,
  ProcessStep,
  AppDefinition,
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
} from '../types/solution';

import { ProcessCategory, WebResourceType, AttributeType, AppType } from '../types/solution';

export interface MarkdownGenerationOptions {
  erdMode?: 'compact' | 'detailed-relationships';
  documentContext?: DocumentContext;
}

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
  return Array.from(new Set(values.filter((value): value is string => !!value && value.trim().length > 0)));
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

/**
 * Appends a Back to Top link at the end of each level-2 section.
 *
 * The link targets the generated Table of Contents heading anchor.
 */
function appendBackToTopLinks(markdown: string): string {
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
      connectors: uniqueStrings([...(current.connectors ?? []), ...(incoming.connectors ?? [])]),
      isEnabled: current.isEnabled !== false || incoming.isEnabled !== false,
      version: current.version || incoming.version,
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
    },
    entities,
    optionSets,
    forms,
    views,
    processes,
    apps,
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
    `> **Generated by PP-MD** — Power Platform Documentation Generator`,
    `> Generated on: ${new Date().toLocaleString()}`,
    '',
    heading(2, 'Solution Overview'),
    '',
    `- **Unique Name:** ${mdEscape(metadata.uniqueName)}`,
    `- **Display Name:** ${mdEscape(metadata.displayName)}`,
    `- **Version:** ${mdEscape(metadata.version)}`,
    `- **Publisher:** ${mdEscape(metadata.publisherName)}`,
    `- **Type:** ${metadata.isManaged ? 'Managed' : 'Unmanaged'}`,
  );

  if (metadata.description) {
    lines.push(`- **Description:** ${mdEscape(metadata.description)}`);
  }

  lines.push('');
  return lines.join('\n');
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
function generateTableOfContents(solution: ParsedSolution): string {
  const lines: string[] = [heading(2, 'Table of Contents'), ''];
  const documentedApps = solution.apps.filter(isDocumentedApp);

  const sections: Array<{ label: string; count: number }> = [
    { label: 'Entity Relationship Diagram',    count: solution.entities.length },
    { label: 'Tables & Columns',               count: solution.entities.length },
    { label: 'Global Option Sets',             count: solution.optionSets.length },
    { label: 'Forms & Views',                  count: solution.forms.length + solution.views.length },
    { label: 'Processes & Automation',         count: solution.processes.length },
    { label: 'Power Apps',                     count: documentedApps.length },
    { label: 'Web Resources',                  count: solution.webResources.length },
    { label: 'Security Roles',                 count: solution.securityRoles.length },
    { label: 'Column Level Security Profiles', count: solution.fieldSecurityProfiles.length },
    { label: 'Connection References',          count: solution.connectionReferences.length },
    { label: 'Environment Variables',          count: solution.environmentVariables.length },
    { label: 'Email Templates',                count: solution.emailTemplates.length },
    { label: 'Reports & Dashboards',           count: solution.reports.length + solution.dashboards.length },
    { label: 'Plugin Assemblies & Steps',      count: solution.pluginAssemblies.length },
  ];

  sections.forEach(({ label, count }) => {
    if (count > 0) {
      lines.push(`- [${label}](${headingAnchor(label)}) _(${count})_`);
    }
  });

  lines.push('');
  return lines.join('\n');
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
      });

      const attrsToShow = entity.attributes.filter((attr) =>
        relationshipColumns.has(attr.name.toLowerCase()) ||
        attr.type === AttributeType.Lookup ||
        attr.type === AttributeType.Owner ||
        attr.type === AttributeType.Customer ||
        attr.type === AttributeType.PartyList ||
        /id$/i.test(attr.name),
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
): string {
  if (entities.length === 0) return '';

  const lines: string[] = [heading(2, 'Tables & Columns'), ''];

  sortByLabel(entities, (entity) => labelWithSchema(entity.displayName || entity.logicalName, entity.logicalName)).forEach((entity) => {
    const entityDisplayName = entity.displayName || entity.logicalName;
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
    if (entity.attributes.length > 0) {
      lines.push(heading(4, 'Columns'));
      lines.push('');

      // Determine which optional columns to show
      const hasDescriptions = entity.attributes.some((a) => !!a.description);
      const hasAuditInfo = entity.attributes.some((a) => a.isAuditEnabled !== undefined);

      if (hasDescriptions && hasAuditInfo) {
        lines.push('| Display Name | Schema Name | Type | Required | Custom | Audited | Notes | Description |');
        lines.push('|--------------|-------------|------|----------|--------|---------|-------|-------------|');
      } else if (hasDescriptions) {
        lines.push('| Display Name | Schema Name | Type | Required | Custom | Notes | Description |');
        lines.push('|--------------|-------------|------|----------|--------|-------|-------------|');
      } else if (hasAuditInfo) {
        lines.push('| Display Name | Schema Name | Type | Required | Custom | Audited | Notes |');
        lines.push('|--------------|-------------|------|----------|--------|---------|-------|');
      } else {
        lines.push('| Display Name | Schema Name | Type | Required | Custom | Notes |');
        lines.push('|--------------|-------------|------|----------|--------|-------|');
      }

      sortByLabel(entity.attributes, (attr) => labelWithSchema(attr.displayName || attr.name, attr.name)).forEach((attr) => {
        const notes: string[] = [];
        if (attr.isPrimaryName) notes.push('🔑 Primary Name');
        if (attr.lookupTarget) notes.push(`→ ${entityDisplayLabel(attr.lookupTarget, entityMap)}`);
        if (attr.optionSetName) notes.push(`OptionSet: ${attr.optionSetName}`);
        if (attr.maxLength)     notes.push(`Max: ${attr.maxLength}`);
        if (attr.precision)     notes.push(`Precision: ${attr.precision}`);

        let row = `| ${mdEscape(attr.displayName || attr.name)} ` +
          `| \`${mdEscape(attr.name)}\` ` +
          `| ${attr.type} ` +
          `| ${attr.required ? '✅ Yes' : 'No'} ` +
          `| ${attr.isCustom ? '✳️ Yes' : 'No'} `;

        if (hasAuditInfo) {
          row += `| ${attr.isAuditEnabled ? '🔍 Yes' : 'No'} `;
        }

        row += `| ${mdEscape(notes.join(', '))} |`;

        if (hasDescriptions) {
          row += ` ${mdEscape(attr.description)} |`;
        }

        lines.push(row);
      });
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
    const localOSAttrs = entity.attributes.filter(
      (a) => (a.type === AttributeType.OptionSet || a.type === AttributeType.MultiSelectOptionSet) && a.options,
    );
    if (localOSAttrs.length > 0) {
      lines.push(heading(4, 'Choice Column Values'));
      sortByLabel(localOSAttrs, (attr) => labelWithSchema(attr.displayName || attr.name, attr.name)).forEach((attr) => {
        if (!attr.options) return;
        lines.push(`**${attr.displayName || attr.name} (\`${attr.name}\`)**:`);
        lines.push('');
        lines.push('| Label | Value | Color | Description |');
        lines.push('|-------|-------|-------|-------------|');
        sortByLabel(attr.options, (opt) => opt.label || String(opt.value)).forEach((opt) => {
          lines.push(`| ${mdEscape(opt.label)} | ${opt.value} | ${mdEscape(opt.color)} | ${mdEscape(opt.description)} |`);
        });
        lines.push('');
      });
    }

    // Cross-reference global option sets used
    const globalOSRefs = entity.attributes
      .filter((a) => a.optionSetName)
      .map((a) => a.optionSetName as string);
    if (globalOSRefs.length > 0) {
      const globalDefs = optionSets.filter((os) => globalOSRefs.includes(os.name));
      if (globalDefs.length > 0) {
        lines.push(heading(4, 'Global Choice References'));
        sortByLabel(globalDefs, (os) => labelWithSchema(os.displayName || os.name, os.name)).forEach((os) => {
          lines.push(`**${os.displayName || os.name} (\`${os.name}\`)**:`);
          lines.push('');
          lines.push('| Label | Value | Color | Description |');
          lines.push('|-------|-------|-------|-------------|');
          sortByLabel(os.options, (opt) => opt.label || String(opt.value)).forEach((opt) => {
            lines.push(`| ${mdEscape(opt.label)} | ${opt.value} | ${mdEscape(opt.color)} | ${mdEscape(opt.description)} |`);
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
    lines.push('| Label | Value | Color | Description |');
    lines.push('|-------|-------|-------|-------------|');
    sortByLabel(os.options, (opt) => opt.label || String(opt.value)).forEach((opt) => {
      lines.push(`| ${mdEscape(opt.label || String(opt.value))} | ${opt.value} | ${mdEscape(opt.color)} | ${mdEscape(opt.description)} |`);
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
        lines.push('| Field (attribute) | Label | Required |');
        lines.push('|-------------------|-------|----------|');
        form.fields.forEach((f) => {
          lines.push(
            `| \`${mdEscape(f.attributeName)}\` ` +
            `| ${mdEscape(f.label)} ` +
            `| ${f.required ? '✅' : ''} |`,
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
 * step tables (no per-process Mermaid flow diagrams).
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

  lines.push('| App Name | Unique Name | Type | Tables | Connectors | Version | Status |');
  lines.push('|----------|-------------|------|--------|------------|---------|--------|');
  sortByLabel(documentedApps, (app) => appTitle(app)).forEach((app) => {
    const tableLabels = sortByLabel(app.entities ?? [], (name) => name).map((name) => entityDisplayLabel(name, entityMap));
    const connectorLabels = sortByLabel(app.connectors ?? [], (name) => name);
    lines.push(
      `| ${mdEscape(appTitle(app))} ` +
      `| \`${mdEscape(app.uniqueName)}\` ` +
      `| ${appTypeLabel(app.appType)} ` +
      `| ${tableLabels.length > 0 ? tableLabels.join(', ') : '–'} ` +
      `| ${connectorLabels.length > 0 ? connectorLabels.map((c) => mdEscape(c)).join(', ') : '–'} ` +
      `| ${app.version || '–'} ` +
      `| ${app.isEnabled !== false ? '✅ Enabled' : '⛔ Disabled'} |`,
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
  attributeDisplayMap: Map<string, string>,
): string {
  if (roles.length === 0 && profiles.length === 0) return '';

  const lines: string[] = [];

  if (roles.length > 0) {
    lines.push(heading(2, 'Security Roles'));
    lines.push('');
    lines.push('| Role Name | Privileges |');
    lines.push('|-----------|-----------|');
    sortByLabel(roles, (role) => role.displayName || role.name).forEach((role) => {
      lines.push(
        `| ${mdEscape(role.displayName || role.name)} ` +
        `| ${role.privileges.length} |`,
      );
    });
    lines.push('');

    sortByLabel(roles, (role) => role.displayName || role.name).forEach((role) => {
      if (role.privileges.length > 0) {
        lines.push(heading(3, role.displayName || role.name));
        lines.push('');

        // Privilege matrix by operation per table
        const matrix = new Map<string, Record<PrivilegeOperation, number>>();
        role.privileges.forEach((priv) => {
          const parsed = parseRolePrivilege(priv.privilegeName);
          if (!parsed) return;
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
        });

        if (matrix.size > 0) {
          lines.push(heading(4, 'Privilege Matrix'));
          lines.push('');
          lines.push('| Table | Logical Name | Create | Read | Write | Delete | Append | Append To | Assign | Share | Unshare |');
          lines.push('|-------|--------------|--------|------|-------|--------|--------|-----------|--------|-------|---------|');

          sortByLabel(Array.from(matrix.entries()), ([table]) => table).forEach(([table, ops]) => {
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
        }
      }
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
      lines.push(
        `| ${mdEscape(ev.displayName || ev.schemaName)} ` +
        `| \`${mdEscape(ev.schemaName)}\` ` +
        `| ${ev.type} ` +
        `| ${ev.hasCurrentValue ? '✅' : '⚠️ Not set'} ` +
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
  const entityMap = buildEntityDisplayMap(solution.entities);
  const attributeDisplayMap = buildAttributeDisplayMap(solution.entities);

  const sections: string[] = [
    generateHeader(solution, options.documentContext),
    generateTableOfContents(solution),
    generateERD(solution.entities, options),
    generateEntitiesSection(solution.entities, solution.optionSets, entityMap),
    generateOptionSetsSection(solution.optionSets),
    generateFormsViewsSection(solution, entityMap),
    generateProcessesSection(solution.processes, entityMap, solution.connectionReferences, solution.environmentVariables),
    generateAppsSection(solution.apps, entityMap),
    generateWebResourcesSection(solution.webResources),
    generateSecuritySection(solution.securityRoles, solution.fieldSecurityProfiles, entityMap, attributeDisplayMap),
    generateIntegrationSection(solution.connectionReferences, solution.environmentVariables, solution.emailTemplates),
    generateReportsSection(solution.reports, solution.dashboards),
    generatePluginsSection(solution.pluginAssemblies),
    generateWarningsSection(solution.warnings),
  ];

  // Filter out empty sections and join
  return appendBackToTopLinks(sections.filter((s) => s.trim().length > 0).join('\n'));
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
      `| ${(app.entities ?? []).length > 0 ? sortByLabel(app.entities ?? [], (e) => e).map((table) => entityDisplayLabel(table, consolidatedEntityMap)).join(', ') : '–'} ` +
      `| ${(app.connectors ?? []).length > 0 ? sortByLabel(app.connectors ?? [], (c) => c).map((c) => mdEscape(c)).join(', ') : '–'} |`,
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

  return appendBackToTopLinks(lines.join('\n'));
}
