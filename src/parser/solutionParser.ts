/**
 * @file solutionParser.ts
 * @description Parses a Power Platform solution ZIP file (Blob or File) into a
 * structured {@link ParsedSolution} object.  This is a pure function module
 * with no React dependencies so it can be unit-tested in isolation.
 *
 * Strategy overview:
 *  1. Open the ZIP with JSZip.
 *  2. Read and parse solution.xml → SolutionMetadata.
 *  3. Read customizations.xml (or [Content_Types].xml variants) → parse entities,
 *     forms, views, processes, web resources, security roles, etc.
 *  4. For each sub-folder (e.g. Workflows/, CanvasApps/, PluginAssemblies/) read
 *     individual definition files as needed.
 *  5. Return the assembled ParsedSolution.
 */

import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';

import type {
  ParsedSolution,
  SolutionMetadata,
  EntityDefinition,
  EntityAttribute,
  EntityRelationship,
  OptionSetDefinition,
  OptionSetOption,
  FormDefinition,
  ViewDefinition,
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
  SecurityRoleDefinition,
  FieldSecurityProfileDefinition,
  ConnectionReferenceDefinition,
  EnvironmentVariableDefinition,
  EmailTemplateDefinition,
  ReportDefinition,
  DashboardDefinition,
  PluginAssemblyDefinition,
  PluginStepDefinition,
  RolePrivilege,
  FormField,
  SolutionDependency,
  AppSiteMapArea,
  AppSiteMapGroup,
  AppSiteMapSubArea,
  AppSiteMapSettings,
  CanvasAppInsights,
  SolutionComponentInventoryItem,
} from '../types/solution';

import {
  AttributeType,
  WebResourceType,
  ProcessCategory,
  AppType,
} from '../types/solution';

// ---------------------------------------------------------------------------
// XML parser configuration
// ---------------------------------------------------------------------------

/**
 * Shared fast-xml-parser instance configured for Dataverse XML conventions.
 * - attributeNamePrefix: preserves attribute names under '@_' key
 * - ignoreAttributes: false – we need XML attributes (e.g. unmodifiedValue=)
 * - isArray: ensures collections are always arrays even when single element
 */
const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
  parseAttributeValue: true,
  // Treat numeric strings as strings to avoid silent coercion bugs
  parseTagValue: false,
  trimValues: true,
  // Prevent entity expansion attacks (OWASP – XXE prevention)
  processEntities: false,
  // Use a strict non-zero limit to prevent Billion-Laughs / expansion attacks
  // (fast-xml-parser vulnerability GHSA-jp2q-39xq-3w4g fixed in 5.5.9)
  htmlEntities: false,
  isArray: (tagName) => {
    // Tags that should always be parsed as arrays
    const pluralTags = [
      'Entity', 'Attribute', 'relationship', 'EntityRelationship',
      'ManyToManyRelationship', 'EntitySetName',
      'optionset', 'option', 'Option', 'OptionSetOption',
      'form', 'Form', 'FormXml',
      'SavedQuery', 'SystemForm',
      'Workflow', 'workflow',
      'AppModule',
      'WebResource',
      'RolePrivilege', 'Role',
      'connectionreference', 'ConnectionReference',
      'environmentvariabledefinition', 'EnvironmentVariableDefinition',
      'PluginAssembly', 'PluginType', 'SdkMessageProcessingStep',
      'Report', 'Dashboard',
      'step', 'Step', 'processStep',
      'FieldPermission', 'FieldSecurityProfile',
    ];
    return pluralTags.includes(tagName);
  },
});

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Safely reads a ZIP entry as a UTF-8 string.  Returns an empty string if the
 * entry does not exist rather than throwing.
 *
 * @param zip    - The JSZip instance
 * @param path   - Path within the ZIP (case-insensitive search)
 * @returns The file contents or an empty string
 */
async function readZipEntry(zip: JSZip, path: string): Promise<string> {
  // Try exact path first, then case-insensitive scan
  let file = zip.file(path);
  if (!file) {
    const lower = path.toLowerCase();
    zip.forEach((relativePath, entry) => {
      if (!file && relativePath.toLowerCase() === lower) {
        file = entry;
      }
    });
  }
  if (!file) return '';
  return file.async('string');
}

/**
 * Collects all ZIP entries whose path starts with a given prefix.
 *
 * @param zip    - The JSZip instance
 * @param prefix - Path prefix to filter by (case-insensitive)
 * @returns Map of relative-path → JSZip entry
 */
function getEntriesWithPrefix(zip: JSZip, prefix: string): Map<string, JSZip['files'][string]> {
  const result = new Map<string, JSZip['files'][string]>();
  const lower = prefix.toLowerCase();
  zip.forEach((path, entry) => {
    if (path.toLowerCase().startsWith(lower)) {
      result.set(path, entry);
    }
  });
  return result;
}

/**
 * Safely reads a property from an XML-parsed object, returning a fallback when
 * the property is missing or undefined.
 *
 * @param obj      - Object to read from
 * @param key      - Property key
 * @param fallback - Value to return when key is absent
 */
function xmlStr(obj: Record<string, unknown>, key: string, fallback = ''): string {
  const val = obj[key];
  if (val === undefined || val === null) return fallback;
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
    return String(val);
  }
  if (Array.isArray(val)) {
    const first = val.find((item) => item !== undefined && item !== null);
    if (first === undefined) return fallback;
    if (typeof first === 'string' || typeof first === 'number' || typeof first === 'boolean') {
      return String(first);
    }
    if (typeof first === 'object') {
      const firstObj = first as Record<string, unknown>;
      return (
        xmlStr(firstObj, '#text') ||
        xmlStr(firstObj, '@_Name') ||
        xmlStr(firstObj, 'Name') ||
        xmlStr(firstObj, '@_name') ||
        xmlStr(firstObj, 'name') ||
        xmlStr(firstObj, 'LogicalName') ||
        xmlStr(firstObj, 'logicalname') ||
        xmlStr(firstObj, 'Value') ||
        xmlStr(firstObj, 'value') ||
        xmlStr(firstObj, 'Description') ||
        xmlStr(firstObj, 'description') ||
        fallback
      );
    }
    return fallback;
  }
  if (typeof val === 'object') {
    const valObj = val as Record<string, unknown>;
    return (
      xmlStr(valObj, '#text') ||
      xmlStr(valObj, '@_Name') ||
      xmlStr(valObj, 'Name') ||
      xmlStr(valObj, '@_name') ||
      xmlStr(valObj, 'name') ||
      xmlStr(valObj, 'LogicalName') ||
      xmlStr(valObj, 'logicalname') ||
      xmlStr(valObj, 'Value') ||
      xmlStr(valObj, 'value') ||
      xmlStr(valObj, 'Description') ||
      xmlStr(valObj, 'description') ||
      fallback
    );
  }
  return String(val);
}

function firstObject(value: unknown): Record<string, unknown> | undefined {
  if (!value) return undefined;
  if (Array.isArray(value)) {
    const first = value.find((item) => typeof item === 'object' && item !== null);
    return first as Record<string, unknown> | undefined;
  }
  if (typeof value === 'object') return value as Record<string, unknown>;
  return undefined;
}

function asObjectArray(value: unknown): Record<string, unknown>[] {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value.filter((item): item is Record<string, unknown> => typeof item === 'object' && item !== null);
  }
  if (typeof value === 'object') return [value as Record<string, unknown>];
  return [];
}

function xmlStrAny(obj: Record<string, unknown>, keys: string[], fallback = ''): string {
  for (const key of keys) {
    const value = xmlStr(obj, key);
    if (value) return value;
  }
  return fallback;
}

function environmentVariableDisplayName(node: Record<string, unknown>, schemaName?: string): string {
  return (
    xmlStrAny(node, ['displayname', 'DisplayName', '@_displayname', '@_DisplayName']) ||
    getLocalizedLabel(node, schemaName || xmlStrAny(node, ['Name', '@_Name'])) ||
    schemaName ||
    ''
  );
}

function parseSolutionDependencies(root: Record<string, unknown>, solNode: Record<string, unknown>): SolutionDependency[] {
  const dependencies: SolutionDependency[] = [];

  const addDependency = (dep: SolutionDependency) => {
    const name = dep.solutionName?.trim();
    if (!name) return;
    const key = `${name.toLowerCase()}|${(dep.version || '').toLowerCase()}`;
    if (dependencies.some((existing) => `${existing.solutionName.toLowerCase()}|${(existing.version || '').toLowerCase()}` === key)) {
      return;
    }
    dependencies.push({
      ...dep,
      solutionName: name,
      displayName: dep.displayName || name,
    });
  };

  const dependencyContainers = [
    solNode['DependencyNodes'],
    solNode['dependencynodes'],
    solNode['Dependencies'],
    solNode['dependencies'],
    root['DependencyNodes'],
    root['dependencynodes'],
    root['Dependencies'],
    root['dependencies'],
  ];

  dependencyContainers.forEach((container) => {
    asObjectArray(container).forEach((depRoot) => {
      const nodes = [
        ...asObjectArray(depRoot['DependencyNode']),
        ...asObjectArray(depRoot['dependencynode']),
        ...asObjectArray(depRoot['Dependency']),
        ...asObjectArray(depRoot['dependency']),
      ];
      nodes.forEach((node) => {
        const solutionName = xmlStrAny(node, [
          '@_solution', '@_Solution',
          '@_solutionname', '@_solutionName',
          'solution', 'Solution',
          'solutionname', 'solutionName',
          '@_schemaname', '@_schemaName',
          'schemaname', 'schemaName',
          '@_name', '@_Name', 'name', 'Name',
        ]);
        if (!solutionName) return;
        addDependency({
          solutionName,
          displayName: xmlStrAny(node, ['@_displayname', '@_displayName', 'displayname', 'displayName', '@_description', 'description']) || solutionName,
          version: xmlStrAny(node, ['@_version', '@_Version', 'version', 'Version']) || undefined,
          isInternal: parseBooleanLike(xmlStrAny(node, ['@_internal', '@_isinternal', 'internal', 'isinternal'])) || false,
          dependentComponentInfo: xmlStrAny(node, ['@_dependentcomponentinfo', 'dependentcomponentinfo', '@_dependentComponentInfo', 'dependentComponentInfo']) || undefined,
        });
      });
    });
  });

  const missingDependencyRoots = [
    root['MissingDependencies'],
    root['missingdependencies'],
    solNode['MissingDependencies'],
    solNode['missingdependencies'],
  ];

  missingDependencyRoots.forEach((missingRoot) => {
    asObjectArray(missingRoot).forEach((container) => {
      const missingItems = [
        ...asObjectArray(container['MissingDependency']),
        ...asObjectArray(container['missingdependency']),
      ];

      missingItems.forEach((missing) => {
        const requiredNode =
          firstObject(missing['Required']) ||
          firstObject(missing['required']) ||
          missing;
        const dependentNode =
          firstObject(missing['Dependent']) ||
          firstObject(missing['dependent']);

        const solutionName = xmlStrAny(requiredNode, [
          '@_solution', '@_Solution',
          '@_solutionname', '@_solutionName',
          'solution', 'Solution',
          'solutionname', 'solutionName',
          '@_requiredsolution', '@_requiredSolution',
          'requiredsolution', 'requiredSolution',
          '@_schemaname', '@_schemaName',
          'schemaname', 'schemaName',
          '@_displayname', '@_displayName',
          'displayname', 'displayName',
        ]);
        if (!solutionName) return;

        const dependentInfo = dependentNode
          ? [
            xmlStrAny(dependentNode, ['@_type', 'type']),
            xmlStrAny(dependentNode, ['@_displayname', '@_displayName', 'displayname', 'displayName']),
            xmlStrAny(dependentNode, ['@_schemaname', '@_schemaName', 'schemaname', 'schemaName']),
          ].filter((value) => !!value).join(' | ') || undefined
          : undefined;

        addDependency({
          solutionName,
          displayName: xmlStrAny(requiredNode, ['@_displayname', '@_displayName', 'displayname', 'displayName', '@_description', 'description']) || solutionName,
          version: xmlStrAny(requiredNode, ['@_version', '@_Version', 'version', 'Version']) || undefined,
          isInternal: parseBooleanLike(xmlStrAny(requiredNode, ['@_internal', '@_isinternal', 'internal', 'isinternal'])) || false,
          dependentComponentInfo: dependentInfo,
        });
      });
    });
  });

  return dependencies;
}

function getLocalizedLabel(node: Record<string, unknown>, fallback = ''): string {
  const directDisplayName =
    xmlStr(node, 'DisplayName') ||
    xmlStr(node, 'displayname') ||
    xmlStr(node, '@_DisplayName') ||
    xmlStr(node, '@_displayname') ||
    xmlStr(node, 'LocalizedName') ||
    xmlStr(node, 'localizedname');
  if (directDisplayName) return directDisplayName;

  const containers = [
    node['LocalizedNames'],
    node['localizednames'],
    node['DisplayNames'],
    node['displaynames'],
    node['Labels'],
    node['labels'],
    node['LocalizedLabels'],
    node['localizedlabels'],
  ];

  for (const container of containers) {
    const c = firstObject(container);
    if (!c) continue;

    const rawLabelNodes = c['LocalizedName'] ?? c['localizedname'] ?? c['displayname'] ?? c['Label'] ?? c['label'] ?? c['LocalizedLabel'] ?? c['localizedlabel'];
    const labelNodes = Array.isArray(rawLabelNodes) ? rawLabelNodes : rawLabelNodes ? [rawLabelNodes] : [];
    if (labelNodes.length === 0) continue;

    const englishNode = labelNodes.find((n) => {
      const ln = n as Record<string, unknown>;
      return xmlStr(ln, '@_languagecode') === '1033' || xmlStr(ln, 'languagecode') === '1033';
    }) as Record<string, unknown> | undefined;

    const selected = englishNode ?? (labelNodes[0] as Record<string, unknown>);
    const label =
      xmlStr(selected, '@_description') ||
      xmlStr(selected, 'description') ||
      xmlStr(selected, '@_label') ||
      xmlStr(selected, 'label') ||
      xmlStr(selected, '#text');

    if (label) return label;
  }

  return fallback;
}

function parsePermissionFlag(value: string): boolean {
  const normalized = value.trim().toLowerCase();
  return ['1', '4', 'true', 'yes', 'allow', 'allowed'].includes(normalized);
}

function parseBooleanLike(value: string | undefined): boolean | undefined {
  if (!value) return undefined;
  const normalized = value.trim().toLowerCase();
  if (!normalized) return undefined;
  if (['1', 'true', 'yes', 'y', 'enabled', 'allow', 'allowed'].includes(normalized)) return true;
  if (['0', 'false', 'no', 'n', 'disabled', 'disallow', 'denied'].includes(normalized)) return false;
  return undefined;
}

function normalizeMetadataKey(key: string): string {
  return key
    .replace(/^@_/, '')
    .replace(/[^a-z0-9]/gi, '')
    .toLowerCase();
}

interface ParsedBooleanMetadata {
  value: boolean;
  sourceKey: string;
}

function parseBooleanMetadataWithSource(value: unknown, candidateKeys: string[]): ParsedBooleanMetadata | undefined {
  const normalizedCandidates = new Set(candidateKeys.map((key) => normalizeMetadataKey(key)));
  const queue: Array<{ value: unknown; path: string }> = [{ value, path: '' }];

  const parseFromNode = (nodeValue: unknown, sourceKey: string): ParsedBooleanMetadata | undefined => {
    if (nodeValue === undefined || nodeValue === null) return undefined;
    if (typeof nodeValue === 'string' || typeof nodeValue === 'number' || typeof nodeValue === 'boolean') {
      const parsed = parseBooleanLike(String(nodeValue));
      return parsed === undefined ? undefined : { value: parsed, sourceKey };
    }

    const node = firstObject(nodeValue);
    if (!node) return undefined;

    const directFields = [
      '@_Value', '@_value',
      'Value', 'value',
      '#text',
      '@_CanBeChanged', 'CanBeChanged', 'canbechanged',
    ];

    for (const field of directFields) {
      const fieldValue = xmlStr(node, field);
      const parsed = parseBooleanLike(fieldValue);
      if (parsed !== undefined) {
        return { value: parsed, sourceKey: `${sourceKey}.${field}` };
      }
    }

    const nested = findFirstStringByKey(node, new Set(['value', '#text']));
    const parsedNested = parseBooleanLike(nested);
    return parsedNested === undefined ? undefined : { value: parsedNested, sourceKey: `${sourceKey}.nested` };
  };

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current) continue;

    if (!current.value || typeof current.value !== 'object') continue;

    if (Array.isArray(current.value)) {
      current.value.forEach((item, index) => queue.push({ value: item, path: `${current.path}[${index}]` }));
      continue;
    }

    const record = current.value as Record<string, unknown>;
    for (const [key, raw] of Object.entries(record)) {
      const path = current.path ? `${current.path}.${key}` : key;

      if (normalizedCandidates.has(normalizeMetadataKey(key))) {
        const parsed = parseFromNode(raw, path);
        if (parsed) return parsed;
      }

      if (raw && typeof raw === 'object') {
        queue.push({ value: raw, path });
      }
    }
  }

  return undefined;
}

function parseBooleanMetadata(value: unknown, candidateKeys: string[]): boolean | undefined {
  return parseBooleanMetadataWithSource(value, candidateKeys)?.value;
}

function parseProcessActivationStatus(node: Record<string, unknown>): boolean | undefined {
  // Different process types export status under different fields.
  // If none are present, return undefined so docs can render a blank status.
  const props = (node['properties'] ?? {}) as Record<string, unknown>;
  const candidates = [
    xmlStr(node, 'Activated'),
    xmlStr(node, '@_Activated'),
    xmlStr(node, 'IsActivated'),
    xmlStr(node, '@_IsActivated'),
    xmlStr(node, 'State'),
    xmlStr(node, '@_State'),
    xmlStr(node, 'statecode'),
    xmlStr(node, '@_statecode'),
    xmlStr(node, 'Status'),
    xmlStr(node, '@_Status'),
    xmlStr(node, 'statuscode'),
    xmlStr(node, '@_statuscode'),
    xmlStr(props, 'Activated'),
    xmlStr(props, 'IsActivated'),
    xmlStr(props, 'State'),
    xmlStr(props, 'status'),
    xmlStr(props, 'state'),
    xmlStr(props, 'statuscode'),
    xmlStr(props, 'statecode'),
  ];

  for (const candidate of candidates) {
    const parsed = parseBooleanLike(candidate);
    if (parsed !== undefined) return parsed;

    const normalized = candidate.trim().toLowerCase();
    if (!normalized) continue;
    if (['active', 'activated', 'started', 'running', 'published'].includes(normalized)) return true;
    if (['inactive', 'deactivated', 'stopped', 'draft', 'unpublished'].includes(normalized)) return false;
  }

  return undefined;
}

function parseRoleDepth(value: string): number {
  const normalized = value.trim().toLowerCase();
  if (!normalized) return 0;

  const asNumber = Number(normalized);
  if (!Number.isNaN(asNumber)) {
    return Math.max(0, Math.min(4, asNumber));
  }

  const namedMap: Record<string, number> = {
    none: 0,
    basic: 1,
    user: 1,
    local: 2,
    businessunit: 2,
    deep: 3,
    parentchildbusinessunit: 3,
    organization: 4,
    organisation: 4,
    global: 4,
  };

  const collapsed = normalized.replace(/[^a-z]/g, '');
  if (namedMap[collapsed] !== undefined) return namedMap[collapsed];

  const embeddedNumber = normalized.match(/\d+/);
  if (embeddedNumber) {
    const n = Number(embeddedNumber[0]);
    if (!Number.isNaN(n)) return Math.max(0, Math.min(4, n));
  }

  return 0;
}

function titleCaseWords(value: string): string {
  return value
    .split(' ')
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ');
}

function humanizeIdentifier(value: string): string {
  const cleaned = value
    .trim()
    .replace(/^shared_/i, '')
    .replace(/[_-]+/g, ' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/\s+/g, ' ');
  return titleCaseWords(cleaned);
}

function stripTrailingGuid(value: string): string {
  return value.replace(/[\s_-]?[0-9a-f]{8}(?:[\s_-]?[0-9a-f]{4}){3}[\s_-]?[0-9a-f]{12}$/i, '').trim();
}

function splitCondensedToken(token: string): string {
  const raw = token.trim();
  if (!raw) return raw;

  const leading = raw.match(/^[^A-Za-z0-9]+/)?.[0] ?? '';
  const trailing = raw.match(/[^A-Za-z0-9]+$/)?.[0] ?? '';
  const core = raw.slice(leading.length, raw.length - trailing.length);
  if (!core || core.length < 12) return raw;

  const hasLongLowerRun = /[a-z]{8,}/.test(core);
  if (!hasLongLowerRun) return raw;

  const lower = core.toLowerCase();
  const dictionary = new Set([
    'a', 'an', 'and', 'as', 'at', 'by', 'for', 'from', 'in', 'into', 'new', 'of', 'on', 'or', 'the', 'to', 'when', 'with',
    'add', 'added', 'approval', 'approve', 'case', 'change', 'create', 'created', 'customer', 'delete', 'email', 'notify',
    'procure', 'process', 'record', 'request', 'send', 'team', 'triage', 'update', 'updated',
  ]);

  const bestByIndex: Array<string[] | undefined> = new Array(lower.length + 1);
  bestByIndex[0] = [];

  for (let index = 0; index < lower.length; index += 1) {
    const base = bestByIndex[index];
    if (!base) continue;

    for (let end = index + 1; end <= lower.length; end += 1) {
      const candidate = lower.slice(index, end);
      if (!dictionary.has(candidate)) continue;

      const next = [...base, candidate];
      const existing = bestByIndex[end];
      if (!existing || next.length < existing.length) {
        bestByIndex[end] = next;
      }
    }
  }

  const segmented = bestByIndex[lower.length];
  if (!segmented || segmented.length < 3) return raw;

  const hasConnectorWord = segmented.some((word) => ['a', 'an', 'and', 'for', 'from', 'in', 'of', 'on', 'or', 'the', 'to', 'when', 'with'].includes(word));
  if (!hasConnectorWord) return raw;

  const rebuilt = segmented
    .map((word, i) => (i === 0 && /^[A-Z]/.test(core) ? word.charAt(0).toUpperCase() + word.slice(1) : word))
    .join(' ');

  return `${leading}${rebuilt}${trailing}`;
}

function normalizeCondensedSegments(value: string): string {
  return value
    .split(/(\s+)/)
    .map((part) => (/^\s+$/.test(part) ? part : splitCondensedToken(part)))
    .join('')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeDisplayNameForReadability(displayName: string | undefined, fallbackName: string): string {
  const trimmedDisplay = stripTrailingGuid((displayName ?? '').trim());
  const fallback = stripTrailingGuid((fallbackName || '').trim());

  if (!trimmedDisplay) {
    return fallback ? humanizeIdentifier(fallback) : '';
  }

  const condensedDisplay = trimmedDisplay.toLowerCase().replace(/[^a-z0-9]/g, '');
  const condensedFallback = fallback.toLowerCase().replace(/[^a-z0-9]/g, '');

  const displayLooksCompressed = /[a-z][A-Z]/.test(trimmedDisplay)
    || /[_-]/.test(trimmedDisplay)
    || (fallback.length > 0 && condensedDisplay === condensedFallback && /[_-]|[a-z][A-Z]/.test(fallback));

  if (displayLooksCompressed) {
    const basis = /[_-]|[a-z][A-Z]/.test(fallback) ? fallback : trimmedDisplay;
    return normalizeCondensedSegments(humanizeIdentifier(basis));
  }

  return normalizeCondensedSegments(trimmedDisplay);
}

function connectorDisplayFromId(connectorId: string): string {
  if (!connectorId) return '';
  const idNoQuery = connectorId.split('?')[0];
  const segment = idNoQuery.split('/').filter(Boolean).pop() ?? idNoQuery;
  const normalized = segment.toLowerCase();
  const normalizedNoShared = normalized.replace(/^shared_/, '');

  const known: Record<string, string> = {
    commondataserviceforapps: 'Microsoft Dataverse',
    commondataservice: 'Microsoft Dataverse',
    office365: 'Office 365',
    office365users: 'Office 365 Users',
    sharepointonline: 'SharePoint',
    teams: 'Microsoft Teams',
    microsoftteams: 'Microsoft Teams',
    sql: 'SQL Server',
    outlook: 'Outlook',
  };

  return known[normalizedNoShared] ?? known[normalized] ?? humanizeIdentifier(segment);
}

function uniqueStrings(values: Array<string | undefined | null>): string[] {
  return Array.from(new Set(values.filter((value): value is string => !!value && value.trim().length > 0)));
}

function isModuleTextFile(path: string): boolean {
  const lower = path.toLowerCase();
  return lower.endsWith('.json') || lower.endsWith('.xml') || lower.endsWith('.yml') || lower.endsWith('.yaml');
}

function readStructuredContent(text: string, path: string): Record<string, unknown> | undefined {
  const lowerPath = path.toLowerCase();
  if (!text.trim()) return undefined;

  if (lowerPath.endsWith('.json')) {
    try {
      return JSON.parse(text) as Record<string, unknown>;
    } catch {
      return undefined;
    }
  }

  if (lowerPath.endsWith('.xml')) {
    try {
      return xmlParser.parse(text) as Record<string, unknown>;
    } catch {
      return undefined;
    }
  }

  return undefined;
}

function findFirstStringByKey(value: unknown, candidateKeys: Set<string>): string | undefined {
  const queue: unknown[] = [value];

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current || typeof current !== 'object') continue;

    if (Array.isArray(current)) {
      current.forEach((item) => queue.push(item));
      continue;
    }

    const record = current as Record<string, unknown>;
    for (const [key, raw] of Object.entries(record)) {
      if (candidateKeys.has(key.toLowerCase()) && (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean')) {
        const valueStr = String(raw).trim();
        if (valueStr) return valueStr;
      }
      if (raw && typeof raw === 'object') queue.push(raw);
    }
  }

  return undefined;
}

function collectStringsByKey(value: unknown, candidateKeys: Set<string>, out: Set<string>): void {
  const queue: unknown[] = [value];

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current || typeof current !== 'object') continue;

    if (Array.isArray(current)) {
      current.forEach((item) => queue.push(item));
      continue;
    }

    const record = current as Record<string, unknown>;
    for (const [key, raw] of Object.entries(record)) {
      if (candidateKeys.has(key.toLowerCase())) {
        if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') {
          const valueStr = String(raw).trim();
          if (valueStr) out.add(valueStr);
        } else if (Array.isArray(raw)) {
          raw.forEach((item) => {
            if (typeof item === 'string' || typeof item === 'number' || typeof item === 'boolean') {
              const valueStr = String(item).trim();
              if (valueStr) out.add(valueStr);
            } else {
              queue.push(item);
            }
          });
        } else if (raw && typeof raw === 'object') {
          queue.push(raw);
        }
      } else if (raw && typeof raw === 'object') {
        queue.push(raw);
      }
    }
  }
}

function findArraySizeByKey(value: unknown, candidateKeys: Set<string>): number | undefined {
  const queue: unknown[] = [value];
  while (queue.length > 0) {
    const current = queue.shift();
    if (!current || typeof current !== 'object') continue;

    if (Array.isArray(current)) {
      current.forEach((item) => queue.push(item));
      continue;
    }

    const record = current as Record<string, unknown>;
    for (const [key, raw] of Object.entries(record)) {
      if (candidateKeys.has(key.toLowerCase()) && Array.isArray(raw)) {
        return raw.length;
      }
      if (raw && typeof raw === 'object') queue.push(raw);
    }
  }

  return undefined;
}

/**
 * Map from solution XML ComponentType number to our {@link ComponentType} enum.
 * Reference: https://learn.microsoft.com/power-apps/developer/data-platform/reference/entities/solutioncomponent
 * Exported for use by future tooling / unit tests.
 */
export const COMPONENT_TYPE_MAP: Record<number, string> = {
  1: 'Entity',
  2: 'Attribute',
  3: 'Relationship',
  4: 'Attribute',          // AttributeMap
  20: 'Role',
  7: 'Role',
  8: 'RolePrivilege',
  9: 'OptionSet',
  10: 'FieldPermission',
  14: 'Workflow',
  24: 'Form',              // SystemForm
  26: 'SavedQuery',        // View
  29: 'Workflow',          // WorkflowAccessRight
  31: 'Report',
  32: 'ReportEntity',
  33: 'ReportCategory',
  35: 'Report',
  36: 'EmailTemplate',
  44: 'Dashboard',
  60: 'Dashboard',
  61: 'WebResource',
  62: 'Process',
  70: 'FieldSecurityProfile',
  71: 'ConnectionReference',
  80: 'AppModule',
  90: 'PluginAssembly',
  91: 'PluginType',
  92: 'SdkMessageProcessingStep',
  95: 'ServiceEndpoint',
  300: 'CanvasApp',
  318: 'AIPlugin',
  380: 'EnvironmentVariableDefinition',
  400: 'WebResource',
  430: 'MobileOfflineProfile',
} as const;

// ---------------------------------------------------------------------------
// solution.xml parsing
// ---------------------------------------------------------------------------

/**
 * Parses the root solution.xml file and returns {@link SolutionMetadata}.
 *
 * @param xml - Raw XML string from solution.xml
 * @returns Parsed metadata
 */
function parseSolutionMetadata(xml: string): SolutionMetadata {
  const doc = xmlParser.parse(xml) as Record<string, unknown>;
  const root = (doc['ImportExportXml'] ?? doc) as Record<string, unknown>;
  const solNode = (root['SolutionManifest'] ?? root['solutionmanifest'] ?? root) as Record<string, unknown>;

  const uniqueName = xmlStrAny(solNode, ['UniqueName', 'uniquename', '@_UniqueName', '@_uniquename']);
  const version    = xmlStrAny(solNode, ['Version', 'version', '@_Version', '@_version']);
  const isManaged  = ['true', '1', 'yes'].includes(xmlStrAny(solNode, ['Managed', 'managed', '@_Managed', '@_managed']).toLowerCase());

  // Publisher node
  const publisher = (solNode['Publisher'] ?? solNode['publisher'] ?? {}) as Record<string, unknown>;
  const publisherName = getLocalizedLabel(
    publisher,
    xmlStrAny(publisher, ['FriendlyName', 'friendlyname', 'DisplayName', 'displayname', 'UniqueName', 'uniquename']) || 'Unknown Publisher',
  );

  // Publisher prefix
  const publisherPrefix = xmlStrAny(publisher, ['Prefix', 'prefix', 'CustomizationPrefix', 'customizationprefix', '@_Prefix', '@_prefix']);

  // Solution display name
  const displayName = getLocalizedLabel(solNode, uniqueName);

  // Description
  const descriptions = firstObject(solNode['Descriptions'] ?? solNode['descriptions']);
  const description =
    xmlStrAny(solNode, ['Description', 'description']) ||
    (descriptions
      ? xmlStrAny(firstObject(descriptions['Description'] ?? descriptions['description']) ?? descriptions, ['@_description', 'description', '#text'])
      : '');

  // Solution prefix / customization prefix
  const solutionPrefix = xmlStrAny(solNode, ['Prefix', 'prefix', 'CustomizationPrefix', 'customizationprefix', '@_Prefix', '@_prefix']);

  // Parse dependencies from all known solution.xml structures
  const dependencies = parseSolutionDependencies(root, solNode);

  const componentInventoryMap = new Map<string, number>();
  const rootComponentsContainer = firstObject(solNode['RootComponents'] ?? solNode['rootcomponents']);
  const rootComponents = rootComponentsContainer
    ? [
      ...asObjectArray(rootComponentsContainer['RootComponent']),
      ...asObjectArray(rootComponentsContainer['rootcomponent']),
    ]
    : [];

  rootComponents.forEach((component) => {
    const componentTypeRaw = xmlStrAny(component, ['@_type', '@_Type', 'type', 'Type', 'ComponentType', 'componenttype']);
    const componentTypeNumber = Number(componentTypeRaw);
    const resolvedType = !Number.isNaN(componentTypeNumber)
      ? (COMPONENT_TYPE_MAP[componentTypeNumber] ?? `Unknown (${componentTypeNumber})`)
      : (componentTypeRaw || 'Unknown');
    componentInventoryMap.set(resolvedType, (componentInventoryMap.get(resolvedType) ?? 0) + 1);
  });

  const componentInventory: SolutionComponentInventoryItem[] = Array.from(componentInventoryMap.entries())
    .map(([componentType, count]) => ({ componentType, count }))
    .sort((a, b) => a.componentType.localeCompare(b.componentType));

  return {
    uniqueName,
    displayName,
    version,
    publisherName,
    publisherUniqueName: xmlStrAny(publisher, ['UniqueName', 'uniquename', '@_UniqueName', '@_uniquename']) || undefined,
    publisherPrefix: publisherPrefix || undefined,
    solutionPrefix: solutionPrefix || undefined,
    description,
    isManaged,
    dependencies,
    componentInventory,
  };
}

function upsertInventoryMinimum(inventoryMap: Map<string, number>, componentType: string, minimumCount: number): void {
  if (!componentType || minimumCount <= 0) return;
  const existing = inventoryMap.get(componentType) ?? 0;
  if (minimumCount > existing) {
    inventoryMap.set(componentType, minimumCount);
  }
}

function enrichComponentInventory(
  metadata: SolutionMetadata,
  data: {
    optionSets: OptionSetDefinition[];
    forms: FormDefinition[];
    views: ViewDefinition[];
    processes: ProcessDefinition[];
    apps: AppDefinition[];
    dataflows: DataflowDefinition[];
    customApis: CustomAPIDefinition[];
    offlineProfiles: OfflineProfileDefinition[];
    webResources: WebResourceDefinition[];
    securityRoles: SecurityRoleDefinition[];
    fieldSecurityProfiles: FieldSecurityProfileDefinition[];
    connectionReferences: ConnectionReferenceDefinition[];
    environmentVariables: EnvironmentVariableDefinition[];
    emailTemplates: EmailTemplateDefinition[];
    reports: ReportDefinition[];
    dashboards: DashboardDefinition[];
    pluginAssemblies: PluginAssemblyDefinition[];
  },
): void {
  const map = new Map(metadata.componentInventory.map((item) => [item.componentType, item.count]));

  upsertInventoryMinimum(map, 'OptionSet', data.optionSets.length);
  upsertInventoryMinimum(map, 'Form', data.forms.length);
  upsertInventoryMinimum(map, 'SavedQuery', data.views.length);
  upsertInventoryMinimum(map, 'Workflow', data.processes.length);
  upsertInventoryMinimum(map, 'Dataflow', data.dataflows.length);
  upsertInventoryMinimum(map, 'CustomAPI', data.customApis.length);
  upsertInventoryMinimum(map, 'MobileOfflineProfile', data.offlineProfiles.length);
  upsertInventoryMinimum(map, 'WebResource', data.webResources.length);
  upsertInventoryMinimum(map, 'Role', data.securityRoles.length);
  upsertInventoryMinimum(map, 'FieldSecurityProfile', data.fieldSecurityProfiles.length);
  upsertInventoryMinimum(map, 'ConnectionReference', data.connectionReferences.length);
  upsertInventoryMinimum(map, 'EnvironmentVariableDefinition', data.environmentVariables.length);
  upsertInventoryMinimum(map, 'EmailTemplate', data.emailTemplates.length);
  upsertInventoryMinimum(map, 'Report', data.reports.length);
  upsertInventoryMinimum(map, 'Dashboard', data.dashboards.length);
  upsertInventoryMinimum(map, 'PluginAssembly', data.pluginAssemblies.length);

  const modelDrivenApps = data.apps.filter((app) => app.appType === AppType.ModelDriven).length;
  const canvasApps = data.apps.filter((app) => app.appType === AppType.Canvas || app.appType === AppType.CustomPage).length;
  upsertInventoryMinimum(map, 'AppModule', modelDrivenApps);
  upsertInventoryMinimum(map, 'CanvasApp', canvasApps);

  metadata.componentInventory = Array.from(map.entries())
    .map(([componentType, count]) => ({ componentType, count } as SolutionComponentInventoryItem))
    .sort((left, right) => left.componentType.localeCompare(right.componentType));
}

// ---------------------------------------------------------------------------
// customizations.xml → entities
// ---------------------------------------------------------------------------

/**
 * Maps the XML Dataverse attribute type string to our {@link AttributeType} enum.
 *
 * @param xmlType - XML type string (e.g. "String", "Lookup", "optionsetattribute")
 */
function mapAttributeType(xmlType: string): AttributeType {
  const raw = (xmlType || '').trim().toLowerCase();
  const lastToken = raw.split('.').pop() || raw;
  const t = lastToken
    .replace(/[_\s-]+/g, '')
    .replace(/metadata$/, '')
    .replace(/attribute$/, '')
    .replace(/type$/, '');

  if (/^\d+$/.test(t)) {
    const numericMap: Record<number, AttributeType> = {
      0: AttributeType.String,
      1: AttributeType.Integer,
      2: AttributeType.Decimal,
      3: AttributeType.Lookup,
      4: AttributeType.DateTime,
      5: AttributeType.Memo,
      6: AttributeType.OptionSet,
      7: AttributeType.Boolean,
      8: AttributeType.Money,
      9: AttributeType.Owner,
      10: AttributeType.UniqueIdentifier,
      11: AttributeType.PartyList,
      12: AttributeType.Lookup,
      13: AttributeType.BigInt,
      14: AttributeType.String,
      15: AttributeType.Memo,
      16: AttributeType.Integer,
      17: AttributeType.OptionSet,
      18: AttributeType.MultiSelectOptionSet,
      19: AttributeType.ManagedProperty,
      20: AttributeType.File,
      21: AttributeType.Image,
    };
    return numericMap[Number(t)] ?? AttributeType.Unknown;
  }

  const map: Record<string, AttributeType> = {
    string:             AttributeType.String,
    nvarchar:           AttributeType.String,
    nchar:              AttributeType.String,
    varchar:            AttributeType.String,
    char:               AttributeType.String,
    memo:               AttributeType.Memo,
    ntext:              AttributeType.Memo,
    text:               AttributeType.Memo,
    decimal:            AttributeType.Decimal,
    double:             AttributeType.Decimal,
    float:              AttributeType.Decimal,
    integer:            AttributeType.Integer,
    int:                AttributeType.Integer,
    whole:              AttributeType.Integer,
    bigint:             AttributeType.BigInt,
    money:              AttributeType.Money,
    boolean:            AttributeType.Boolean,
    datetime:           AttributeType.DateTime,
    date:               AttributeType.DateTime,
    lookup:             AttributeType.Lookup,
    owner:              AttributeType.Owner,
    optionset:          AttributeType.OptionSet,
    picklist:           AttributeType.OptionSet,
    state:              AttributeType.OptionSet,
    status:             AttributeType.OptionSet,
    statusreason:       AttributeType.OptionSet,
    multiselectoptionset: AttributeType.MultiSelectOptionSet,
    multiselectpicklist:  AttributeType.MultiSelectOptionSet,
    uniqueidentifier:   AttributeType.UniqueIdentifier,
    guid:               AttributeType.UniqueIdentifier,
    image:              AttributeType.Image,
    file:               AttributeType.File,
    customer:           AttributeType.Customer,
    partylist:          AttributeType.PartyList,
    virtual:            AttributeType.Virtual,
    managedproperty:    AttributeType.ManagedProperty,
  };

  if (t.includes('lookup')) return AttributeType.Lookup;
  if (t.includes('owner')) return AttributeType.Owner;
  if (t.includes('customer')) return AttributeType.Customer;
  if (t.includes('partylist')) return AttributeType.PartyList;

  return map[t] ?? AttributeType.Unknown;
}

/**
 * Maps the numeric web resource type code to {@link WebResourceType}.
 *
 * @param code - Numeric type code from XML attribute
 */
function mapWebResourceType(code: number | string): WebResourceType {
  const map: Record<number, WebResourceType> = {
    1:  WebResourceType.HTML,
    2:  WebResourceType.CSS,
    3:  WebResourceType.JavaScript,
    4:  WebResourceType.XML,
    5:  WebResourceType.PNG,
    6:  WebResourceType.JPG,
    7:  WebResourceType.GIF,
    8:  WebResourceType.XAP,
    9:  WebResourceType.XSL,
    10: WebResourceType.ICO,
    11: WebResourceType.SVG,
    12: WebResourceType.Resx,
  };
  return map[Number(code)] ?? WebResourceType.Unknown;
}

/**
 * Parses a single entity node from customizations.xml into an
 * {@link EntityDefinition}.
 *
 * @param entityNode - Parsed XML entity node object
 * @param warnings   - Array to accumulate non-fatal warnings
 */
function parseEntityNode(
  entityNode: Record<string, unknown>,
  warnings: string[],
  customPrefixes: string[] = [],
): EntityDefinition {
  const entityInfo = (entityNode['EntityInfo'] ?? entityNode) as Record<string, unknown>;
  const entity     = (entityInfo['entity'] ?? entityInfo) as Record<string, unknown>;

  const logicalName = (
    xmlStr(entityNode, 'Name') ||
    xmlStr(entityNode, '@_Name') ||
    xmlStr(entityNode, 'LogicalName') ||
    xmlStr(entity, 'Name') ||
    xmlStr(entity, '@_Name') ||
    xmlStr(entity, 'LogicalName') ||
    xmlStr(entity, 'logicalname')
  ).toLowerCase();
  const displayName = (() => {
    const nameNode = firstObject(entityNode['Name']);
    const explicitDisplayName =
      xmlStr(nameNode ?? {}, '@_LocalizedName') ||
      xmlStr(nameNode ?? {}, '@_DisplayName') ||
      xmlStr(nameNode ?? {}, 'LocalizedName') ||
      xmlStr(nameNode ?? {}, 'DisplayName');

    return (
      explicitDisplayName ||
      getLocalizedLabel(entity, '') ||
      getLocalizedLabel(entityNode, logicalName)
    );
  })();

  const isCustom         =
    parseBooleanMetadata(entity, ['IsCustomEntity']) ??
    parseBooleanLike(xmlStr(entity, 'IsCustomEntity')) ??
    parseBooleanLike(xmlStr(entity, '@_IsCustomEntity')) ??
    false;
  const isActivity       = xmlStr(entity, 'IsActivity') === 'true';
  const changeTracking   = xmlStr(entity, 'ChangeTrackingEnabled') === 'true';
  const objectTypeCode   = (() => {
    const otc = entity['ObjectTypeCode'] as string | number | undefined;
    return otc !== undefined ? Number(otc) : undefined;
  })();
  const ownershipType    = xmlStr(entity, 'OwnershipType') as 'User' | 'Organization' | 'None' | '';

  // Parse attributes
  const attributesRoot = (entity['attributes'] ?? entity['Attributes'] ?? {}) as Record<string, unknown>;
  const rawAttrs: unknown[] = (() => {
    const a = attributesRoot['attribute'] ?? attributesRoot['Attribute'];
    if (!a) return [];
    return Array.isArray(a) ? a : [a];
  })();

  const normalizedCustomPrefixes = customPrefixes
    .map((prefix) => prefix.trim().toLowerCase())
    .filter(Boolean);

  const attributes: EntityAttribute[] = rawAttrs.map((rAttr) => {
    const a = rAttr as Record<string, unknown>;
    const attrName    = xmlStr(a, 'Name').toLowerCase() || xmlStr(a, '@_Name').toLowerCase() || xmlStr(a, 'LogicalName');
    const typeStr     =
      xmlStr(a, 'Type') ||
      xmlStr(a, '@_Type') ||
      xmlStr(a, 'AttributeType') ||
      xmlStr(a, 'attributetype') ||
      xmlStr(a, 'AttributeTypeName') ||
      xmlStr(a, 'attributetypename') ||
      xmlStr(a, 'AttributeTypeDisplayName') ||
      xmlStr(a, 'attributetypedisplayname') ||
      xmlStr(a, 'TypeName') ||
      xmlStr(a, 'typename') ||
      findFirstStringByKey(a, new Set(['type', 'attributetype', 'attributetypename', 'attributetypedisplayname', 'typename'])) ||
      '';
    const mappedAttrType = mapAttributeType(typeStr);
    // RequiredLevel may be a plain string OR a child element with @_Value attribute (schema varies by export version)
    const requiredLevelValue = (() => {
      const rl = a['RequiredLevel'];
      if (!rl) return '';
      if (typeof rl === 'string') return rl;
      const rlObj = firstObject(rl);
      return (rlObj ? xmlStr(rlObj, '@_Value') || xmlStr(rlObj, 'Value') || xmlStr(rlObj, '#text') : '') || String(rl);
    })();
    const parseBoolMeta = (keys: string[]): boolean | undefined => parseBooleanMetadata(a, keys);
    const parseBoolMetaWithSource = (keys: string[]): ParsedBooleanMetadata | undefined => parseBooleanMetadataWithSource(a, keys);
    const requiredLevelNormalized = requiredLevelValue.toLowerCase();
    const required = ['required', 'systemrequired', 'applicationrequired'].includes(requiredLevelNormalized);
    const customMeta = parseBoolMetaWithSource(['IsCustomAttribute', 'IsCustomField']);
    const isCustomByPrefix =
      !!attrName &&
      normalizedCustomPrefixes.some((prefix) => attrName.toLowerCase().startsWith(`${prefix}_`));
    const isCustomAttr = customMeta?.value ?? isCustomByPrefix;
    const isPrimaryName = xmlStr(a, 'IsPrimaryName') === 'true' || xmlStr(a, '@_IsPrimaryName') === 'true';
    const isAuditEnabled = parseBoolMeta(['IsAuditEnabled']);
    const isSecured = parseBoolMeta(['IsSecured']);
    const displayMaskValue = xmlStr(a, 'DisplayMask') || xmlStr(a, 'displaymask');
    const advancedFindFromDisplayMask: ParsedBooleanMetadata | undefined =
      displayMaskValue && displayMaskValue.toLowerCase().split('|').includes('validforadvancedfind')
        ? { value: true, sourceKey: 'DisplayMask' }
        : undefined;
    const advancedFindMeta =
      parseBoolMetaWithSource(['IsValidForAdvancedFind', 'ValidForAdvancedFind']) ?? advancedFindFromDisplayMask;
    const isValidForAdvancedFind = advancedFindMeta?.value;
    const isManaged = parseBoolMeta(['IsManaged']);
    const numberMeta = (key: string): number | undefined => {
      const direct = xmlStr(a, key) || xmlStr(a, `@_${key}`) || xmlStr(a, key.toLowerCase()) || xmlStr(a, `@_${key.toLowerCase()}`);
      if (direct) {
        const parsed = Number(direct);
        if (!Number.isNaN(parsed)) return parsed;
      }
      const node = firstObject(a[key] ?? a[key.toLowerCase()]);
      if (!node) return undefined;
      const nodeValue = xmlStr(node, '@_Value') || xmlStr(node, 'Value') || xmlStr(node, '#text');
      if (!nodeValue) return undefined;
      const parsed = Number(nodeValue);
      return Number.isNaN(parsed) ? undefined : parsed;
    };
    const textMeta = (keys: string[]): string | undefined => {
      for (const key of keys) {
        const direct = xmlStr(a, key) || xmlStr(a, `@_${key}`) || xmlStr(a, key.toLowerCase()) || xmlStr(a, `@_${key.toLowerCase()}`);
        if (direct) return direct;

        const node = firstObject(a[key] ?? a[key.toLowerCase()]);
        if (!node) continue;
        const nodeValue = xmlStr(node, '@_Value') || xmlStr(node, 'Value') || xmlStr(node, '#text');
        if (nodeValue) return nodeValue;
      }
      return undefined;
    };

    const maxLength = numberMeta('MaxLength');
    const precision = numberMeta('Precision');
    const minValue =
      textMeta(['MinValue', 'Min']) ||
      findFirstStringByKey(a, new Set(['minvalue', 'min'])) ||
      numberMeta('MinValue')?.toString();
    const maxValue =
      textMeta(['MaxValue', 'Max']) ||
      findFirstStringByKey(a, new Set(['maxvalue', 'max'])) ||
      numberMeta('MaxValue')?.toString();
    const format = textMeta(['Format', 'DateTimeBehavior', 'ImeMode']) || findFirstStringByKey(a, new Set(['format', 'datetimebehavior', 'imemode']));
    const defaultValue = textMeta(['DefaultValue', 'Default']) || findFirstStringByKey(a, new Set(['defaultvalue', 'default'])) || undefined;

    const lookupTargets: string[] | undefined = (() => {
      const targetsNode = a['Targets'] ?? a['targets'];
      if (!targetsNode) return undefined;

      const parsedTargets: string[] = [];

      if (typeof targetsNode === 'string') {
        targetsNode.split(',').map((part) => part.trim()).filter(Boolean).forEach((item) => parsedTargets.push(item));
      } else {
        const targetsObj = firstObject(targetsNode);
        if (targetsObj) {
          const rawTarget = targetsObj['Target'] ?? targetsObj['target'];
          const targetNodes = Array.isArray(rawTarget) ? rawTarget : rawTarget ? [rawTarget] : [];
          targetNodes.forEach((target) => {
            if (typeof target === 'string') {
              const value = target.trim();
              if (value) parsedTargets.push(value);
              return;
            }
            const t = target as Record<string, unknown>;
            const value = xmlStr(t, '#text') || xmlStr(t, '@_Name') || xmlStr(t, 'Name');
            if (value) parsedTargets.push(value.trim());
          });
        }
      }

      return uniqueStrings(parsedTargets);
    })();

    const attrType =
      mappedAttrType === AttributeType.Unknown && lookupTargets && lookupTargets.length > 0
        ? AttributeType.Lookup
        : mappedAttrType;

    const lookupTarget = lookupTargets?.[0];

    const optionSetName: string | undefined = (() => {
      if (attrType !== AttributeType.OptionSet && attrType !== AttributeType.MultiSelectOptionSet) {
        return undefined;
      }

      const osRef = firstObject(a['OptionSet'] ?? a['optionset']);
      if (!osRef) {
        const directName = xmlStr(a, 'OptionSetName') || xmlStr(a, 'optionsetname');
        return directName || undefined;
      }

      return xmlStr(osRef, '@_Name') || xmlStr(osRef, 'Name') || undefined;
    })();

    const attrDisplayName = getLocalizedLabel(a, attrName);

    // Attribute description from Descriptions/Description[@languagecode='1033']/@description
    const attrDescription = (() => {
      const descsNode = firstObject(a['Descriptions'] ?? a['descriptions']);
      if (!descsNode) return undefined;
      const descItems = descsNode['Description'] ?? descsNode['description'];
      const descArr = Array.isArray(descItems) ? descItems : descItems ? [descItems] : [];
      const engNode = descArr.find((d: unknown) => {
        const dn = d as Record<string, unknown>;
        return xmlStr(dn, '@_languagecode') === '1033' || xmlStr(dn, 'languagecode') === '1033';
      }) as Record<string, unknown> | undefined ?? (descArr[0] as Record<string, unknown> | undefined);
      return engNode ? (xmlStr(engNode, '@_description') || xmlStr(engNode, 'description') || undefined) : undefined;
    })();

    const options = (() => {
      if (attrType !== AttributeType.OptionSet && attrType !== AttributeType.MultiSelectOptionSet) {
        return undefined;
      }

      const osContainer = firstObject(a['OptionSet'] ?? a['optionset']);
      if (!osContainer) return undefined;

      const optionSetDefaultRaw =
        xmlStr(osContainer, '@_DefaultValue') ||
        xmlStr(osContainer, '@_defaultvalue') ||
        xmlStr(osContainer, 'DefaultValue') ||
        xmlStr(osContainer, 'defaultvalue') ||
        xmlStr(firstObject(osContainer['DefaultValue'] ?? osContainer['defaultvalue']) ?? {}, '@_Value') ||
        xmlStr(firstObject(osContainer['DefaultValue'] ?? osContainer['defaultvalue']) ?? {}, 'Value') ||
        undefined;
      const optionSetDefaultNumber = optionSetDefaultRaw !== undefined && optionSetDefaultRaw !== ''
        ? Number(optionSetDefaultRaw)
        : undefined;

      const valuesContainer = firstObject(osContainer['Options'] ?? osContainer['options']);
      const rawOptions = valuesContainer?.['Option'] ?? valuesContainer?.['option'];
      const optionNodes = Array.isArray(rawOptions) ? rawOptions : rawOptions ? [rawOptions] : [];
      if (optionNodes.length === 0) return undefined;

      const parsed = optionNodes
        .map((opt) => {
          const o = opt as Record<string, unknown>;
          const rawValue = xmlStr(o, '@_value') || xmlStr(o, 'value') || xmlStr(o, '@_Value') || xmlStr(o, 'Value');
          const value = Number(rawValue);
          if (Number.isNaN(value)) return undefined;

          const label = getLocalizedLabel(o, xmlStr(o, '@_description') || rawValue);
          const description =
            xmlStr(o, 'Description') ||
            xmlStr(o, 'description') ||
            xmlStr(firstObject(o['Descriptions'] ?? o['descriptions']) ?? {}, 'Description') ||
            undefined;

          return {
            value,
            label,
            description,
            color: xmlStr(o, '@_color') || undefined,
            isDefault:
              (xmlStr(o, '@_default') || xmlStr(o, '@_isdefault') || xmlStr(o, 'default') || xmlStr(o, 'isdefault') || '').toLowerCase() === 'true' ||
              (!!defaultValue && defaultValue === String(value)) ||
              (optionSetDefaultNumber !== undefined && !Number.isNaN(optionSetDefaultNumber) && optionSetDefaultNumber === value),
          } as OptionSetOption;
        })
        .filter((item): item is OptionSetOption => !!item);

      return parsed.length > 0 ? parsed : undefined;
    })();

    return {
      name:            attrName,
      displayName:     attrDisplayName,
      description:     attrDescription,
      type:            attrType,
      required,
      requiredLevel:   requiredLevelValue || undefined,
      isCustom:        isCustomAttr,
      isPrimaryName,
      isAuditEnabled,
      isSecured,
      isValidForAdvancedFind,
      isManaged,
      maxLength,
      precision,
      minValue,
      maxValue,
      format,
      defaultValue,
      lookupTarget,
      lookupTargets,
      options,
      optionSetName,
      metadataSources:
        customMeta || advancedFindMeta
          ? {
            isCustom: customMeta?.sourceKey,
            isValidForAdvancedFind: advancedFindMeta?.sourceKey,
          }
          : undefined,
    } as EntityAttribute;
  });

  // Parse relationships
  const relationships: EntityRelationship[] = [];
  const relRoot = (entity['EntityRelationships'] ?? entity['relationships'] ?? {}) as Record<string, unknown>;

  const parseRelCascade = (ri: Record<string, unknown>) => {
    const cascNode = firstObject(ri['CascadeConfiguration'] ?? ri['cascadeconfiguration']);
    return {
      cascadeDelete:   cascNode ? (xmlStr(cascNode, 'Delete')   || xmlStr(cascNode, 'delete')   || undefined) : (xmlStr(ri, 'CascadeDelete') || undefined),
      cascadeAssign:   cascNode ? (xmlStr(cascNode, 'Assign')   || xmlStr(cascNode, 'assign')   || undefined) : undefined,
      cascadeReparent: cascNode ? (xmlStr(cascNode, 'Reparent') || xmlStr(cascNode, 'reparent') || undefined) : undefined,
    };
  };

  const parseRelDescription = (ri: Record<string, unknown>): string | undefined => {
    const descsNode = firstObject(ri['LocalizedDescriptions'] ?? ri['Descriptions'] ?? ri['descriptions']);
    if (!descsNode) return undefined;
    const descItems = descsNode['Description'] ?? descsNode['description'];
    const descArr = Array.isArray(descItems) ? descItems : descItems ? [descItems] : [];
    const engNode = descArr.find((d: unknown) => {
      const dn = d as Record<string, unknown>;
      return xmlStr(dn, '@_languagecode') === '1033' || xmlStr(dn, 'languagecode') === '1033';
    }) as Record<string, unknown> | undefined ?? (descArr[0] as Record<string, unknown> | undefined);
    return engNode ? (xmlStr(engNode, '@_description') || xmlStr(engNode, 'description') || undefined) : undefined;
  };

  const parseRel = (relNode: unknown, relType: 'OneToMany' | 'ManyToMany' | 'ManyToOne') => {
    const r = relNode as Record<string, unknown>;
    const relItems = r['EntityRelationship'] ?? r['relationship'] ?? r;
    const items = Array.isArray(relItems) ? relItems : [relItems];
    items.forEach((item) => {
      const ri = item as Record<string, unknown>;
      const cascade = parseRelCascade(ri);
      relationships.push({
        name:                    xmlStr(ri, '@_Name') || xmlStr(ri, 'Name') || xmlStr(ri, 'SchemaName'),
        type:                    relType,
        referencedEntity:        xmlStr(ri, 'ReferencedEntity').toLowerCase() || logicalName,
        referencingEntity:       xmlStr(ri, 'ReferencingEntity').toLowerCase() || logicalName,
        referencingAttribute:    xmlStr(ri, 'ReferencingAttribute') || undefined,
        referencedAttribute:     xmlStr(ri, 'ReferencedAttribute') || undefined,
        cascadeDelete:           cascade.cascadeDelete,
        cascadeAssign:           cascade.cascadeAssign,
        cascadeReparent:         cascade.cascadeReparent,
        relationshipDescription: parseRelDescription(ri),
      });
    });
  };

  const o2m = relRoot['OneToManyRelationships'] as Record<string, unknown> | undefined;
  const m2m = relRoot['ManyToManyRelationships'] as Record<string, unknown> | undefined;
  const m2o = relRoot['ManyToOneRelationships'] as Record<string, unknown> | undefined;
  if (o2m) parseRel(o2m, 'OneToMany');
  if (m2m) parseRel(m2m, 'ManyToMany');
  if (m2o) parseRel(m2o, 'ManyToOne');

  if (!logicalName) {
    warnings.push(`Encountered entity node without a logical name – skipping`);
  }

  // Entity set name (OData collection name) and primary attribute
  const entitySetName     = xmlStr(entity, 'EntitySetName') || xmlStr(entity, 'CollectionSchemaName') || undefined;
  const primaryAttributeName = xmlStr(entity, 'PrimaryAttribute') || xmlStr(entity, 'PrimaryAttributeName') || undefined;

  // Entity description
  const entityDescription = (() => {
    const descsNode = firstObject(entity['Descriptions'] ?? entity['descriptions']);
    if (!descsNode) return undefined;
    const descItems = descsNode['Description'] ?? descsNode['description'];
    const descArr = Array.isArray(descItems) ? descItems : descItems ? [descItems] : [];
    const engNode = descArr.find((d: unknown) => {
      const dn = d as Record<string, unknown>;
      return xmlStr(dn, '@_languagecode') === '1033' || xmlStr(dn, 'languagecode') === '1033';
    }) as Record<string, unknown> | undefined ?? (descArr[0] as Record<string, unknown> | undefined);
    return engNode ? (xmlStr(engNode, '@_description') || xmlStr(engNode, 'description') || undefined) : undefined;
  })();

  return {
    name:                 logicalName,
    logicalName,
    displayName,
    description:          entityDescription,
    isCustom,
    isActivity,
    changeTracking,
    objectTypeCode,
    ownershipType:        (ownershipType || 'User') as 'User' | 'Organization' | 'None',
    attributes,
    relationships,
    keys:                 [],
    entitySetName,
    primaryAttributeName,
  };
}

function normalizeRelationshipType(rawType: string): 'OneToMany' | 'ManyToMany' | 'ManyToOne' {
  const normalized = rawType.trim().toLowerCase();
  if (normalized === 'manytomany' || normalized === 'many-to-many') return 'ManyToMany';
  if (normalized === 'manytoone' || normalized === 'many-to-one') return 'ManyToOne';
  return 'OneToMany';
}

/**
 * Parses root-level <EntityRelationship> tags (as emitted by many Dataverse
 * exports) and merges them into the relevant entity relationship arrays.
 */
function mergeRootEntityRelationships(
  root: Record<string, unknown>,
  entities: EntityDefinition[],
): void {
  const relContainer = (root['EntityRelationships'] ?? root['entityrelationships'] ?? {}) as Record<string, unknown>;
  const rawRootRelationships = relContainer['EntityRelationship'] ?? relContainer['entityrelationship'];
  if (!rawRootRelationships) return;

  const relItems = Array.isArray(rawRootRelationships) ? rawRootRelationships : [rawRootRelationships];
  const entityByName = new Map(entities.map((entity) => [entity.logicalName.toLowerCase(), entity]));
  const seen = new Set<string>();

  // Preserve any relationships that were already parsed from per-entity sections.
  entities.forEach((entity) => {
    entity.relationships.forEach((rel) => {
      seen.add(`${rel.name}|${rel.type}|${rel.referencedEntity}|${rel.referencingEntity}|${rel.referencingAttribute || ''}`.toLowerCase());
    });
  });

  relItems.forEach((item) => {
    const relNode = item as Record<string, unknown>;
    const cascNode = firstObject(relNode['CascadeConfiguration'] ?? relNode['cascadeconfiguration']);
    const relationship: EntityRelationship = {
      name:                    xmlStr(relNode, '@_Name') || xmlStr(relNode, 'Name') || xmlStr(relNode, 'SchemaName'),
      type:                    normalizeRelationshipType(xmlStr(relNode, 'EntityRelationshipType') || xmlStr(relNode, 'RelationshipType') || 'OneToMany'),
      referencedEntity:        (xmlStr(relNode, 'ReferencedEntityName') || xmlStr(relNode, 'ReferencedEntity')).toLowerCase(),
      referencingEntity:       (xmlStr(relNode, 'ReferencingEntityName') || xmlStr(relNode, 'ReferencingEntity')).toLowerCase(),
      referencingAttribute:    xmlStr(relNode, 'ReferencingAttributeName') || xmlStr(relNode, 'ReferencingAttribute') || undefined,
      referencedAttribute:     xmlStr(relNode, 'ReferencedAttributeName') || xmlStr(relNode, 'ReferencedAttribute') || undefined,
      cascadeDelete:           cascNode ? (xmlStr(cascNode, 'Delete') || xmlStr(cascNode, 'delete') || undefined) : (xmlStr(relNode, 'CascadeDelete') || undefined),
      cascadeAssign:           cascNode ? (xmlStr(cascNode, 'Assign') || xmlStr(cascNode, 'assign') || undefined) : undefined,
      cascadeReparent:         cascNode ? (xmlStr(cascNode, 'Reparent') || xmlStr(cascNode, 'reparent') || undefined) : undefined,
    };

    if (!relationship.referencedEntity || !relationship.referencingEntity) return;

    const dedupeKey = `${relationship.name}|${relationship.type}|${relationship.referencedEntity}|${relationship.referencingEntity}|${relationship.referencingAttribute || ''}`.toLowerCase();
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);

    const referencingEntity = entityByName.get(relationship.referencingEntity);
    if (referencingEntity) {
      referencingEntity.relationships.push(relationship);
    }

    // Also attach an inverted perspective to the referenced entity for better per-entity docs.
    const referencedEntity = entityByName.get(relationship.referencedEntity);
    if (referencedEntity && relationship.type !== 'ManyToMany') {
      const inverseType: 'ManyToOne' | 'OneToMany' = relationship.type === 'OneToMany' ? 'ManyToOne' : 'OneToMany';
      const inverse: EntityRelationship = {
        ...relationship,
        type: inverseType,
      };
      const inverseKey = `${inverse.name}|${inverse.type}|${inverse.referencedEntity}|${inverse.referencingEntity}|${inverse.referencingAttribute || ''}`.toLowerCase();
      if (!seen.has(inverseKey)) {
        seen.add(inverseKey);
        referencedEntity.relationships.push(inverse);
      }
    }
  });
}

// ---------------------------------------------------------------------------
// customizations.xml → processes
// ---------------------------------------------------------------------------

/**
 * Maps the numeric Process Category to our {@link ProcessCategory} enum.
 *
 * @param cat - Category number (e.g. 0 = Workflow, 4 = Action, 4 = BusinessProcessFlow)
 */
function mapProcessCategory(cat: number): ProcessCategory {
  const map: Record<number, ProcessCategory> = {
    0:  ProcessCategory.Workflow,
    1:  ProcessCategory.Dialog,
    2:  ProcessCategory.BusinessRule,
    3:  ProcessCategory.Action,
    4:  ProcessCategory.BusinessProcessFlow,
    5:  ProcessCategory.CustomAction,
    6:  ProcessCategory.PowerAutomateFlow,
  };
  return map[cat] ?? ProcessCategory.Workflow;
}

function addEntityMatchesFromText(text: string, knownEntities: Set<string>, matches: Set<string>): void {
  const normalized = text.toLowerCase();
  knownEntities.forEach((entity) => {
    if (
      normalized === entity ||
      normalized.includes(`/${entity}`) ||
      normalized.includes(`.${entity}`) ||
      normalized.includes(`'${entity}'`) ||
      normalized.includes(`"${entity}"`) ||
      normalized.includes(` ${entity} `) ||
      normalized.includes(`(${entity})`) ||
      normalized.includes(`=${entity}`)
    ) {
      matches.add(entity);
    }
  });
}

function extractReferencedEntities(value: unknown, knownEntities: Set<string>, matches: Set<string>): void {
  if (!value) return;

  if (typeof value === 'string') {
    addEntityMatchesFromText(value, knownEntities, matches);
    return;
  }

  if (Array.isArray(value)) {
    value.forEach((item) => extractReferencedEntities(item, knownEntities, matches));
    return;
  }

  if (typeof value === 'object') {
    Object.entries(value as Record<string, unknown>).forEach(([key, inner]) => {
      addEntityMatchesFromText(key, knownEntities, matches);
      extractReferencedEntities(inner, knownEntities, matches);
    });
  }
}

function classifyCanvasAppType(metadataText: string): AppType | undefined {
  const normalized = metadataText.toLowerCase();
  if (normalized.includes('custompage') || normalized.includes('custom page')) {
    return AppType.CustomPage;
  }
  if (normalized.includes('codeapp') || normalized.includes('code app') || normalized.includes('aiplugin') || normalized.includes('ai plugin')) {
    return AppType.CodeApp;
  }
  if (
    normalized.includes('canvasmanifest') ||
    normalized.includes('connectionreferences') ||
    normalized.includes('screenorder') ||
    normalized.includes('publishinfo') ||
    normalized.includes('appsettings')
  ) {
    return AppType.Canvas;
  }
  return undefined;
}

function extractConnectorNames(value: unknown, out: Set<string>): void {
  if (!value) return;
  if (typeof value === 'string') {
    const normalized = value.toLowerCase();
    const sharedMatch = normalized.match(/shared_[a-z0-9_]+/g) ?? [];
    sharedMatch.forEach((token) => out.add(connectorDisplayFromId(token)));
    if (normalized.includes('/apis/')) {
      out.add(connectorDisplayFromId(value));
    }
    return;
  }
  if (Array.isArray(value)) {
    value.forEach((item) => extractConnectorNames(item, out));
    return;
  }
  if (typeof value === 'object') {
    Object.entries(value as Record<string, unknown>).forEach(([key, inner]) => {
      extractConnectorNames(key, out);
      extractConnectorNames(inner, out);
    });
  }
}

function extractSitemapAreas(rawText: string): string[] {
  if (!rawText) return [];
  const areas = new Set<string>();
  const areaRegex = /<area[^>]*\b(?:title|name|id)="([^"]+)"/gi;
  let m: RegExpExecArray | null = areaRegex.exec(rawText);
  while (m) {
    const value = m[1]?.trim();
    if (value) areas.add(value);
    m = areaRegex.exec(rawText);
  }
  return Array.from(areas);
}

function parseSiteMapStructure(rawText: string): {
  areas: string[];
  entities: string[];
  structure: AppSiteMapArea[];
  settings?: AppSiteMapSettings;
} {
  if (!rawText) return { areas: [], entities: [], structure: [] };

  const areas = new Set<string>();
  const entities = new Set<string>();
  const structure: AppSiteMapArea[] = [];
  let settings: AppSiteMapSettings | undefined;

  const normalizeLabel = (node: Record<string, unknown>, fallback = ''): string =>
    xmlStrAny(node, [
      '@_Title', '@_title',
      '@_Name', '@_name',
      '@_Id', '@_id',
      'Title', 'title',
      'Name', 'name',
      'Id', 'id',
    ]) || fallback;

  const normalizeEntity = (node: Record<string, unknown>): string =>
    xmlStrAny(node, ['@_Entity', '@_entity', 'Entity', 'entity']);

  const findEmbeddedSiteMap = (value: unknown): Record<string, unknown> | undefined => {
    if (!value) return undefined;
    if (Array.isArray(value)) {
      for (const item of value) {
        const found = findEmbeddedSiteMap(item);
        if (found) return found;
      }
      return undefined;
    }
    if (typeof value !== 'object') return undefined;

    const obj = value as Record<string, unknown>;
    const direct = firstObject(obj['SiteMap']) || firstObject(obj['sitemap']);
    if (direct) return direct;

    for (const child of Object.values(obj)) {
      const found = findEmbeddedSiteMap(child);
      if (found) return found;
    }
    return undefined;
  };

  try {
    const doc = xmlParser.parse(rawText) as Record<string, unknown>;
    const siteMapRoot = findEmbeddedSiteMap(doc) || doc;

    settings = {
      showHome: parseBooleanLike(xmlStrAny(siteMapRoot, ['@_ShowHome', '@_showhome', 'ShowHome', 'showhome'])),
      showPinned: parseBooleanLike(xmlStrAny(siteMapRoot, ['@_ShowPinned', '@_showpinned', 'ShowPinned', 'showpinned'])),
      showRecents: parseBooleanLike(xmlStrAny(siteMapRoot, ['@_ShowRecents', '@_showrecents', 'ShowRecents', 'showrecents'])),
      enableCollapsibleGroups: parseBooleanLike(xmlStrAny(siteMapRoot, ['@_EnableCollapsibleGroups', '@_enablecollapsiblegroups', 'EnableCollapsibleGroups', 'enablecollapsiblegroups'])),
    };

    const buildSubArea = (node: Record<string, unknown>): AppSiteMapSubArea => {
      const entity = normalizeEntity(node);
      if (entity) entities.add(entity);
      return {
        id: xmlStrAny(node, ['@_Id', '@_id', 'Id', 'id']) || undefined,
        title: normalizeLabel(node) || undefined,
        entity: entity || undefined,
        url: xmlStrAny(node, ['@_Url', '@_url', 'Url', 'url']) || undefined,
      };
    };

    const buildGroup = (node: Record<string, unknown>): AppSiteMapGroup => {
      const subAreas = [
        ...asObjectArray(node['SubArea']),
        ...asObjectArray(node['subarea']),
      ].map(buildSubArea);

      return {
        id: xmlStrAny(node, ['@_Id', '@_id', 'Id', 'id']) || undefined,
        title: normalizeLabel(node) || undefined,
        subAreas,
      };
    };

    const buildArea = (node: Record<string, unknown>): AppSiteMapArea => {
      const title = normalizeLabel(node);
      if (title) areas.add(title);

      const groups = [
        ...asObjectArray(node['Group']),
        ...asObjectArray(node['group']),
      ].map(buildGroup);

      const directSubAreas = [
        ...asObjectArray(node['SubArea']),
        ...asObjectArray(node['subarea']),
      ];
      if (directSubAreas.length > 0) {
        groups.push({
          id: undefined,
          title: undefined,
          subAreas: directSubAreas.map(buildSubArea),
        });
      }

      return {
        id: xmlStrAny(node, ['@_Id', '@_id', 'Id', 'id']) || undefined,
        title: title || undefined,
        groups,
      };
    };

    const areaNodes = [
      ...asObjectArray(siteMapRoot['Area']),
      ...asObjectArray(siteMapRoot['area']),
    ];

    areaNodes.forEach((areaNode) => {
      structure.push(buildArea(areaNode));
    });
  } catch {
    // Fall through to regex-based extraction only
  }

  extractSitemapAreas(rawText).forEach((value) => areas.add(value));

  const entityRegex = /<(?:SubArea|subarea)[^>]*\b(?:Entity|entity)="([^"]+)"/g;
  let match = entityRegex.exec(rawText);
  while (match) {
    const value = match[1]?.trim();
    if (value) entities.add(value);
    match = entityRegex.exec(rawText);
  }

  return {
    areas: Array.from(areas),
    entities: Array.from(entities),
    structure,
    settings,
  };
}

function findFlowMatches(flow: Record<string, unknown>, candidates: string[]): string[] {
  if (candidates.length === 0) return [];
  const serialized = JSON.stringify(flow).toLowerCase();
  return Array.from(new Set(candidates.filter((candidate) => candidate && serialized.includes(candidate.toLowerCase()))));
}

function normalizeIdentifierKey(value: string | undefined): string {
  return (value ?? '').trim().toLowerCase();
}

function mergeAppDefinition(existing: AppDefinition, incoming: AppDefinition): AppDefinition {
  const mergedSiteMap = incoming.siteMap && incoming.siteMap.length > 0
    ? incoming.siteMap
    : existing.siteMap;

  return {
    ...existing,
    ...incoming,
    name: existing.name || incoming.name,
    displayName: existing.displayName || incoming.displayName,
    uniqueName: existing.uniqueName || incoming.uniqueName,
    version: existing.version || incoming.version,
    entities: uniqueStrings([...(existing.entities ?? []), ...(incoming.entities ?? [])]),
    sitemapAreas: uniqueStrings([...(existing.sitemapAreas ?? []), ...(incoming.sitemapAreas ?? [])]),
    connectors: uniqueStrings([...(existing.connectors ?? []), ...(incoming.connectors ?? [])]),
    siteMap: mergedSiteMap,
    siteMapSettings: existing.siteMapSettings || incoming.siteMapSettings,
    canvasInsights: mergeCanvasInsights(existing.canvasInsights, incoming.canvasInsights),
  };
}

function extractCanvasInsightsFromText(rawText: string): CanvasAppInsights | undefined {
  if (!rawText) return undefined;

  const screenNames = new Set<string>();
  const navigation = new Set<string>();

  const screenRegex = /^\s*([A-Za-z0-9_\- ]{1,120})\s+As\s+screen\b/gim;
  let screenMatch = screenRegex.exec(rawText);
  while (screenMatch) {
    const name = (screenMatch[1] ?? '').trim();
    if (name) screenNames.add(name);
    screenMatch = screenRegex.exec(rawText);
  }

  const navRegex = /Navigate\s*\(\s*([A-Za-z0-9_\- ]{1,120})/gim;
  let navMatch = navRegex.exec(rawText);
  while (navMatch) {
    const target = (navMatch[1] ?? '').replace(/["'`]/g, '').trim();
    if (target) navigation.add(target);
    navMatch = navRegex.exec(rawText);
  }

  if (screenNames.size === 0 && navigation.size === 0) return undefined;

  const insight: CanvasAppInsights = {};
  if (screenNames.size > 0) {
    insight.screenNames = Array.from(screenNames).sort((a, b) => a.localeCompare(b));
    insight.screenCount = screenNames.size;
  }
  if (navigation.size > 0) {
    const defaultSource = insight.screenNames?.[0] ?? 'Unknown';
    insight.navigation = Array.from(navigation)
      .map((to) => ({ from: defaultSource, to }))
      .sort((a, b) => `${a.from}|${a.to}`.localeCompare(`${b.from}|${b.to}`));
  }
  return insight;
}

function countNodes(
  value: unknown,
  predicate: (node: Record<string, unknown>) => boolean,
): number {
  let count = 0;
  const visit = (current: unknown) => {
    if (!current) return;
    if (Array.isArray(current)) {
      current.forEach(visit);
      return;
    }
    if (typeof current !== 'object') return;

    const node = current as Record<string, unknown>;
    if (predicate(node)) count += 1;
    Object.values(node).forEach(visit);
  };
  visit(value);
  return count;
}

function extractCanvasAppInsights(metadata: Record<string, unknown>): CanvasAppInsights {
  const screenNames = new Set<string>();
  const screenControls = new Map<string, Set<string>>();
  const navigationLinks = new Set<string>();
  const screens = countNodes(metadata, (node) => {
    const type = xmlStrAny(node, ['Type', 'type', 'ControlType', 'controlType', '@_Type', '@_type']).toLowerCase();
    const kind = xmlStrAny(node, ['Kind', 'kind']).toLowerCase();
    const isScreen = type === 'screen' || kind === 'screen';
    if (isScreen) {
      const screenName = xmlStrAny(node, ['Name', 'name', 'DisplayName', 'displayName']);
      if (screenName) screenNames.add(screenName);
    }
    return isScreen;
  });

  const controls = countNodes(metadata, (node) => {
    const type = xmlStrAny(node, ['ControlType', 'controlType', 'Type', 'type']).toLowerCase();
    if (!type) return false;
    return !['screen', 'appinfo', 'app'].includes(type);
  });

  const dataSources = new Set<string>();
  const variables = new Set<string>();
  const resources = new Set<string>();

  const isControlNode = (node: Record<string, unknown>): boolean => {
    const type = xmlStrAny(node, ['ControlType', 'controlType', 'Type', 'type', '@_type', '@_Type']).toLowerCase();
    if (!type) return false;
    return !['screen', 'appinfo', 'app', 'datasource', 'table'].includes(type);
  };

  const visit = (value: unknown, currentScreen = '') => {
    if (!value) return;
    if (Array.isArray(value)) {
      value.forEach((item) => visit(item, currentScreen));
      return;
    }
    if (typeof value !== 'object') return;

    const node = value as Record<string, unknown>;
    const nodeName = xmlStrAny(node, ['Name', 'name', 'DisplayName', 'displayName']);
    const nodeType = xmlStrAny(node, ['Type', 'type', 'Kind', 'kind', 'ControlType', 'controlType']).toLowerCase();

    let nextScreen = currentScreen;
    if (nodeType === 'screen' && nodeName) {
      nextScreen = nodeName;
      screenNames.add(nodeName);
      if (!screenControls.has(nodeName)) screenControls.set(nodeName, new Set<string>());
    } else if (nextScreen && nodeName && isControlNode(node)) {
      if (!screenControls.has(nextScreen)) screenControls.set(nextScreen, new Set<string>());
      screenControls.get(nextScreen)!.add(nodeName);
    }

    const serializedNode = JSON.stringify(node);
    const navRegex = /Navigate\s*\(\s*([^,)]+)/gi;
    let navMatch = navRegex.exec(serializedNode);
    while (navMatch) {
      const targetRaw = navMatch[1]?.replace(/["'`]/g, '').trim();
      if (targetRaw) {
        const target = targetRaw.replace(/\b(Screen|scr)_?/gi, '').trim() || targetRaw;
        if (nextScreen && target) {
          navigationLinks.add(`${nextScreen}|${target}`);
        }
      }
      navMatch = navRegex.exec(serializedNode);
    }

    const dsName = xmlStrAny(node, ['DataSourceName', 'dataSourceName', 'Name', 'name']);
    if (dsName && (nodeType.includes('datasource') || nodeType.includes('table'))) {
      dataSources.add(dsName);
    }

    const varName = xmlStrAny(node, ['VariableName', 'variableName', 'CollectionName', 'collectionName']);
    if (varName) variables.add(varName);

    const media = xmlStrAny(node, ['MediaName', 'mediaName', 'ResourceName', 'resourceName', 'FileName', 'fileName']);
    if (media && /\.(png|jpg|jpeg|gif|svg|bmp|ico|mp4|mp3|wav|json)$/i.test(media)) {
      resources.add(media);
    }

    Object.values(node).forEach((inner) => visit(inner, nextScreen));
  };

  visit(metadata);

  const insight: CanvasAppInsights = {};
  const resolvedScreenCount = Math.max(screens, screenNames.size);
  if (resolvedScreenCount > 0) insight.screenCount = resolvedScreenCount;
  if (controls > 0) insight.controlCount = controls;
  if (dataSources.size > 0) insight.dataSourceCount = dataSources.size;
  if (variables.size > 0) insight.variableCount = variables.size;
  if (resources.size > 0) insight.resourceCount = resources.size;
  if (screenNames.size > 0) insight.screenNames = Array.from(screenNames).sort((a, b) => a.localeCompare(b));
  if (dataSources.size > 0) insight.dataSources = Array.from(dataSources).sort((a, b) => a.localeCompare(b));
  if (variables.size > 0) insight.variables = Array.from(variables).sort((a, b) => a.localeCompare(b));
  if (resources.size > 0) insight.resources = Array.from(resources).sort((a, b) => a.localeCompare(b));
  if (screenControls.size > 0) {
    insight.screens = Array.from(screenControls.entries())
      .map(([name, controls]) => ({
        name,
        controls: Array.from(controls).sort((a, b) => a.localeCompare(b)),
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }
  if (navigationLinks.size > 0) {
    insight.navigation = Array.from(navigationLinks)
      .map((value) => {
        const [from, to] = value.split('|');
        return { from, to };
      })
      .filter((link) => !!link.from && !!link.to)
      .sort((a, b) => `${a.from}|${a.to}`.localeCompare(`${b.from}|${b.to}`));
  }
  return insight;
}

function mergeCanvasInsights(base: CanvasAppInsights | undefined, incoming: CanvasAppInsights | undefined): CanvasAppInsights | undefined {
  if (!base && !incoming) return undefined;

  const mergeCount = (left: number | undefined, right: number | undefined): number | undefined => {
    const value = Math.max(left ?? 0, right ?? 0);
    return value > 0 ? value : undefined;
  };

  const mergedScreenNames = uniqueStrings([...(base?.screenNames ?? []), ...(incoming?.screenNames ?? [])]);
  const mergedDataSources = uniqueStrings([...(base?.dataSources ?? []), ...(incoming?.dataSources ?? [])]);
  const mergedVariables = uniqueStrings([...(base?.variables ?? []), ...(incoming?.variables ?? [])]);
  const mergedResources = uniqueStrings([...(base?.resources ?? []), ...(incoming?.resources ?? [])]);

  const merged: CanvasAppInsights = {
    screenCount: mergeCount(base?.screenCount, incoming?.screenCount),
    controlCount: mergeCount(base?.controlCount, incoming?.controlCount),
    dataSourceCount: mergeCount(base?.dataSourceCount, incoming?.dataSourceCount),
    variableCount: mergeCount(base?.variableCount, incoming?.variableCount),
    resourceCount: mergeCount(base?.resourceCount, incoming?.resourceCount),
    screenNames: mergedScreenNames,
    dataSources: mergedDataSources,
    variables: mergedVariables,
    resources: mergedResources,
  };

  if ((merged.screenCount ?? 0) < mergedScreenNames.length) {
    merged.screenCount = mergedScreenNames.length;
  }
  if ((merged.dataSourceCount ?? 0) < mergedDataSources.length) {
    merged.dataSourceCount = mergedDataSources.length;
  }
  if ((merged.variableCount ?? 0) < mergedVariables.length) {
    merged.variableCount = mergedVariables.length;
  }
  if ((merged.resourceCount ?? 0) < mergedResources.length) {
    merged.resourceCount = mergedResources.length;
  }

  const screenMap = new Map<string, Set<string>>();
  [...(base?.screens ?? []), ...(incoming?.screens ?? [])].forEach((screen) => {
    if (!screenMap.has(screen.name)) screenMap.set(screen.name, new Set<string>());
    screen.controls.forEach((control) => screenMap.get(screen.name)!.add(control));
  });
  if (screenMap.size > 0) {
    merged.screens = Array.from(screenMap.entries())
      .map(([name, controls]) => ({ name, controls: Array.from(controls).sort((a, b) => a.localeCompare(b)) }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  const navMap = new Set<string>();
  [...(base?.navigation ?? []), ...(incoming?.navigation ?? [])].forEach((link) => {
    navMap.add(`${link.from}|${link.to}`);
  });
  if (navMap.size > 0) {
    merged.navigation = Array.from(navMap)
      .map((value) => {
        const [from, to] = value.split('|');
        return { from, to };
      })
      .sort((a, b) => `${a.from}|${a.to}`.localeCompare(`${b.from}|${b.to}`));
  }

  return merged;
}

function parseDashboardComponents(formXml: string): string[] {
  if (!formXml) return [];
  const components = new Set<string>();

  try {
    const doc = xmlParser.parse(formXml) as Record<string, unknown>;
    const visit = (value: unknown): void => {
      if (!value) return;
      if (Array.isArray(value)) {
        value.forEach(visit);
        return;
      }
      if (typeof value !== 'object') return;

      const node = value as Record<string, unknown>;
      const descriptor =
        xmlStr(node, '@_name') ||
        xmlStr(node, '@_id') ||
        xmlStr(node, '@_chartid') ||
        xmlStr(node, '@_gridid') ||
        xmlStr(node, '@_controlid') ||
        xmlStr(node, 'name');
      if (descriptor) components.add(descriptor);
      Object.values(node).forEach(visit);
    };

    visit(doc);
  } catch {
    return [];
  }

  return Array.from(components);
}

/**
 * Extracts process steps from a workflow activity node.
 * Handles both classic workflow <step> elements and BPF <stage> elements.
 *
 * @param activityNode - Parsed XML activity / workflow XML node
 */
function extractWorkflowSteps(activityNode: Record<string, unknown>): ProcessStep[] {
  const steps: ProcessStep[] = [];

  // Business Process Flows store stages, not steps
  const stagesRoot = (activityNode['stages'] ?? activityNode['Stages'] ?? {}) as Record<string, unknown>;
  const rawStages = stagesRoot['stage'] ?? stagesRoot['Stage'];
  if (rawStages) {
    const stageItems = Array.isArray(rawStages) ? rawStages : [rawStages];
    stageItems.forEach((s, idx) => {
      const stageNode = s as Record<string, unknown>;
      steps.push({
        id:          xmlStr(stageNode, '@_stageid') || xmlStr(stageNode, 'stageid') || String(idx),
        name:        xmlStr(stageNode, '@_name') || xmlStr(stageNode, 'name') || `Stage ${idx + 1}`,
        stepType:    'Stage',
        description: xmlStr(stageNode, '@_stagecategory') ? `Category: ${xmlStr(stageNode, '@_stagecategory')}` : undefined,
      });
    });
    if (steps.length > 0) return steps;
  }

  // Classic workflows store steps as <step> elements in ActivityXml
  const rawSteps = activityNode['step'] ?? activityNode['Step'] ?? activityNode['steps'];
  if (!rawSteps) return steps;
  const items = Array.isArray(rawSteps) ? rawSteps : [rawSteps];
  items.forEach((s, idx) => {
    const stepNode = s as Record<string, unknown>;
    steps.push({
      id:       xmlStr(stepNode, '@_stepId') || xmlStr(stepNode, 'stepId') || String(idx),
      name:     xmlStr(stepNode, 'stepName') || xmlStr(stepNode, '@_name') || `Step ${idx + 1}`,
      stepType: xmlStr(stepNode, '@_type') || xmlStr(stepNode, 'type') || 'Unknown',
      description: xmlStr(stepNode, 'description') || undefined,
    });
  });
  return steps;
}

// ---------------------------------------------------------------------------
// Power Automate flow JSON parsing
// ---------------------------------------------------------------------------

/**
 * Extracts meaningful information from a Power Automate flow definition JSON.
 * Handles both the real PP solution export format ({ properties: { definition } })
 * and the simplified Logic App format ({ definition } or direct).
 *
 * @param flowJson - Parsed JSON object (from the flow's definition file)
 */
function parseFlowDefinition(
  flowJson: Record<string, unknown>,
): { steps: ProcessStep[]; connectors: string[]; triggerDescription: string; displayName: string | undefined; triggerEntity: string | undefined } {
  const steps: ProcessStep[] = [];
  const connectors = new Set<string>();
  let triggerDescription = '';
  let displayName: string | undefined;
  let triggerEntity: string | undefined;

  try {
    // Real Power Platform solution exports wrap everything under 'properties'
    const props = (flowJson['properties'] ?? {}) as Record<string, unknown>;

    // Prefer display name from properties
    const propDisplayName =
      xmlStr(props, 'displayName') ||
      xmlStr(props, 'DisplayName') ||
      xmlStr(flowJson, 'displayName') ||
      xmlStr(flowJson, 'DisplayName');
    if (propDisplayName) displayName = propDisplayName;

    // Handle three formats:
    // 1. { properties: { definition: { triggers, actions } } }  — real PP export
    // 2. { definition: { triggers, actions } }
    // 3. { triggers, actions }  — direct Logic App definition
    const def = (props['definition'] ?? flowJson['definition'] ?? flowJson) as Record<string, unknown>;
    const triggers = (def['triggers'] ?? {}) as Record<string, unknown>;
    const actions  = (def['actions']  ?? {}) as Record<string, unknown>;

    // Use connectionReferences from properties for the most accurate connector names
    const connRefs = (props['connectionReferences'] ?? {}) as Record<string, unknown>;
    Object.entries(connRefs).forEach(([key, ref]) => {
      const refObj = ref as Record<string, unknown>;
      const refDisplayName = xmlStr(refObj, 'displayName') || xmlStr(refObj, 'DisplayName');
      if (refDisplayName) {
        connectors.add(refDisplayName);
      } else {
        const id = xmlStr(refObj, 'id') || xmlStr(refObj, 'apiId');
        if (id) connectors.add(connectorDisplayFromId(id));
        else if (key) connectors.add(humanizeIdentifier(key));
      }
    });

    // Parse trigger — humanize key names like "When_a_row_is_added" → "When a row is added"
    // Also extract the primary entity (table) the trigger fires on when available.
    Object.entries(triggers).forEach(([triggerName, triggerDef]) => {
      const td = triggerDef as Record<string, unknown>;
      const triggerType = xmlStr(td, 'type');
      const inputs = (td['inputs'] ?? {}) as Record<string, unknown>;
      const apiId = ((inputs['host'] ?? {}) as Record<string, unknown>)['apiId'] as string | undefined;
      if (apiId) connectors.add(connectorDisplayFromId(apiId));
      triggerDescription = `${humanizeIdentifier(triggerName)} (${triggerType || '–'})`;

      // Extract the Dataverse table the trigger fires on (Dataverse connector triggers).
      if (!triggerEntity) {
        const params = (inputs['parameters'] ?? inputs['body'] ?? {}) as Record<string, unknown>;
        const rawEntity =
          (params['entityName'] as string | undefined) ||
          (params['tableName'] as string | undefined) ||
          (params['entity_name'] as string | undefined) ||
          (params['table_name'] as string | undefined) ||
          (inputs['entityName'] as string | undefined) ||
          (inputs['tableName'] as string | undefined);
        if (rawEntity) triggerEntity = rawEntity.toLowerCase().trim();
      }
    });

    // Parse actions recursively
    const parseActions = (actionMap: Record<string, unknown>, parentId?: string): ProcessStep[] => {
      return Object.entries(actionMap).map(([actionName, actionDef]) => {
        const ad = actionDef as Record<string, unknown>;
        const actionType = xmlStr(ad, 'type');

        // Extract connector from host.apiId or host.connectionName
        const inputs  = (ad['inputs'] ?? {}) as Record<string, unknown>;
        const host    = (inputs['host'] ?? {}) as Record<string, unknown>;
        const apiId   = host['apiId'] as string | undefined ?? host['connectionName'] as string | undefined;
        if (apiId) connectors.add(connectorDisplayFromId(apiId));

        // Extract Dataverse table references from action inputs (List rows, Get row, Create record, etc.)
        const params = (inputs['parameters'] ?? inputs['body'] ?? {}) as Record<string, unknown>;
        const referencedTableRaw =
          (params['entityName'] as string | undefined) ||
          (params['tableName'] as string | undefined) ||
          (params['entity_name'] as string | undefined) ||
          (params['table_name'] as string | undefined) ||
          (inputs['entityName'] as string | undefined) ||
          (inputs['tableName'] as string | undefined);
        const stepReferencedEntities: string[] = referencedTableRaw
          ? [referencedTableRaw.toLowerCase().trim()]
          : [];

        // Recurse into nested actions (Scope/Foreach/Until/Condition/Switch branches).
        let children: ProcessStep[] | undefined;
        const nested: ProcessStep[] = [];

        const nestedActions = ad['actions'] as Record<string, unknown> | undefined;
        if (nestedActions) {
          nested.push(...parseActions(nestedActions, actionName));
        }

        const elseActions = ((ad['else'] as Record<string, unknown> | undefined)?.['actions']) as Record<string, unknown> | undefined;
        if (elseActions) {
          nested.push(...parseActions(elseActions, `${actionName}_else`));
        }

        const cases = ad['cases'] as Record<string, unknown> | undefined;
        if (cases) {
          Object.entries(cases).forEach(([caseName, caseNode]) => {
            const caseActions = ((caseNode as Record<string, unknown>)['actions']) as Record<string, unknown> | undefined;
            if (caseActions) {
              nested.push(...parseActions(caseActions, `${actionName}_${caseName}`));
            }
          });
        }

        const defaultActions = ((ad['default'] as Record<string, unknown> | undefined)?.['actions']) as Record<string, unknown> | undefined;
        if (defaultActions) {
          nested.push(...parseActions(defaultActions, `${actionName}_default`));
        }

        if (nested.length > 0) children = nested;

        // Collect referenced entities from children to bubble up
        const childEntities = nested.flatMap((child) => child.referencedEntities ?? []);
        const allReferencedEntities = [...stepReferencedEntities, ...childEntities].filter((v, i, arr) => arr.indexOf(v) === i);

        const actionDescription = [`Type: ${actionType}`];
        const operation = xmlStr(ad, 'operationOptions');
        if (operation) actionDescription.push(`Operation: ${operation}`);

        return {
          id:       parentId ? `${parentId}_${actionName}` : actionName,
          name:     humanizeIdentifier(actionName),
          stepType: actionType,
          description: actionDescription.join(' | '),
          children,
          referencedEntities: allReferencedEntities.length > 0 ? allReferencedEntities : undefined,
        } as ProcessStep;
      }).filter((s) => !!s.name);
    };

    steps.push(...parseActions(actions));
  } catch {
    // best-effort; return what we have
  }

  return {
    steps,
    connectors: Array.from(connectors),
    triggerDescription,
    displayName,
    triggerEntity,
  };
}

// ---------------------------------------------------------------------------
// customizations.xml → web resources
// ---------------------------------------------------------------------------

/**
 * Parses a WebResource node from customizations.xml.
 *
 * @param node - Parsed XML web resource node
 */
function parseWebResourceNode(node: Record<string, unknown>): WebResourceDefinition {
  const name       = xmlStr(node, 'Name') || xmlStr(node, '@_Name');
  const schemaName = xmlStr(node, 'SchemaName') || name;
  const typeCode   = node['WebResourceType'] ?? node['@_WebResourceType'] ?? 0;
  const rType      = mapWebResourceType(Number(typeCode));
  const content64  = xmlStr(node, 'Content');
  let content: string | undefined;
  let contentLength: number | undefined;

  if (content64) {
    try {
      // Decode from base64 to string for text-based resources
      if ([WebResourceType.JavaScript, WebResourceType.TypeScript, WebResourceType.HTML,
           WebResourceType.CSS, WebResourceType.XML, WebResourceType.XSL,
           WebResourceType.SVG, WebResourceType.Resx].includes(rType)) {
        content = atob(content64);
        contentLength = content.length;
      } else {
        contentLength = Math.ceil(content64.length * 0.75); // approximate byte length
      }
    } catch {
      contentLength = 0;
    }
  }

  const displayName = getLocalizedLabel(node, name);

  const enabledForMobile = parseBooleanLike(
    xmlStr(node, 'IsEnabledForMobileClient') ||
    xmlStr(node, '@_IsEnabledForMobileClient') ||
    xmlStr(node, 'EnabledForMobileClient') ||
    xmlStr(node, '@_EnabledForMobileClient') ||
    xmlStr(node, 'EnabledForMobile') ||
    xmlStr(node, '@_EnabledForMobile'),
  );

  const availableOffline = parseBooleanLike(
    xmlStr(node, 'IsAvailableForMobileOffline') ||
    xmlStr(node, '@_IsAvailableForMobileOffline') ||
    xmlStr(node, 'AvailableForMobileOffline') ||
    xmlStr(node, '@_AvailableForMobileOffline') ||
    xmlStr(node, 'IsOfflineAvailable') ||
    xmlStr(node, '@_IsOfflineAvailable') ||
    xmlStr(node, 'Offline') ||
    xmlStr(node, '@_Offline'),
  );

  return {
    name,
    schemaName,
    displayName,
    resourceType: rType,
    enabledForMobile,
    availableOffline,
    content,
    contentLength,
  };
}

// ---------------------------------------------------------------------------
// Main parse function
// ---------------------------------------------------------------------------

/**
 * Parses a Power Platform solution ZIP file into a {@link ParsedSolution}.
 *
 * @param file    - The ZIP file as a Browser File or Blob object
 * @param onProgress - Optional progress callback receiving a 0–100 number
 * @returns A fully-populated ParsedSolution
 * @throws If the ZIP cannot be read or does not contain a solution.xml
 */
export async function parseSolutionZip(
  file: File | Blob,
  onProgress?: (percent: number) => void,
): Promise<ParsedSolution> {
  const warnings: string[] = [];
  const report = (pct: number) => onProgress?.(pct);

  // ── Step 1: Open the ZIP ────────────────────────────────────────────────
  report(5);
  const zip = await JSZip.loadAsync(file);

  // ── Step 2: Parse solution.xml ──────────────────────────────────────────
  report(10);
  const solutionXml = await readZipEntry(zip, 'solution.xml');
  if (!solutionXml) {
    throw new Error('Invalid Power Platform solution: solution.xml not found in the ZIP archive.');
  }
  const metadata = parseSolutionMetadata(solutionXml);

  // ── Step 3: Parse customizations.xml ───────────────────────────────────
  report(20);
  const customizationsXml = await readZipEntry(zip, 'customizations.xml');

  const entities:                  EntityDefinition[]                  = [];
  const optionSets:                OptionSetDefinition[]               = [];
  const forms:                     FormDefinition[]                    = [];
  const views:                     ViewDefinition[]                    = [];
  const processes:                 ProcessDefinition[]                 = [];
  const apps:                      AppDefinition[]                     = [];
  const agents:                    AgentDefinition[]                   = [];
  const aiModels:                  AIModelDefinition[]                 = [];
  const desktopFlows:              DesktopFlowDefinition[]             = [];
  const dataflows:                 DataflowDefinition[]                = [];
  const customApis:                CustomAPIDefinition[]               = [];
  const offlineProfiles:           OfflineProfileDefinition[]          = [];
  const webResources:              WebResourceDefinition[]             = [];
  const securityRoles:             SecurityRoleDefinition[]            = [];
  const fieldSecurityProfiles:     FieldSecurityProfileDefinition[]    = [];
  const connectionReferences:      ConnectionReferenceDefinition[]     = [];
  const environmentVariables:      EnvironmentVariableDefinition[]     = [];
  const emailTemplates:            EmailTemplateDefinition[]           = [];
  const reports:                   ReportDefinition[]                  = [];
  const dashboards:                DashboardDefinition[]               = [];
  const pluginAssemblies:          PluginAssemblyDefinition[]          = [];

  if (customizationsXml) {
    const doc = xmlParser.parse(customizationsXml) as Record<string, unknown>;
    const root = (doc['ImportExportXml'] ?? doc) as Record<string, unknown>;

    // ── Entities ─────────────────────────────────────────────────────────
    report(25);
    const entitiesRoot = (root['Entities'] ?? root['entities'] ?? {}) as Record<string, unknown>;
    const rawEntities: unknown[] = (() => {
      const e = entitiesRoot['Entity'] ?? entitiesRoot['entity'];
      if (!e) return [];
      return Array.isArray(e) ? e : [e];
    })();
    rawEntities.forEach((e) => {
      try {
        const customPrefixes = [metadata.publisherPrefix, metadata.solutionPrefix].filter((prefix): prefix is string => !!prefix);
        entities.push(parseEntityNode(e as Record<string, unknown>, warnings, customPrefixes));
      } catch (err) {
        warnings.push(`Failed to parse entity: ${(err as Error).message}`);
      }
    });

    // Some solution exports represent relationships only at the root level
    // via <EntityRelationships><EntityRelationship .../></EntityRelationships>.
    mergeRootEntityRelationships(root, entities);

    // ── Global OptionSets ─────────────────────────────────────────────────
    report(30);
    const optSetsRoot = (root['optionsets'] ?? root['OptionSets'] ?? {}) as Record<string, unknown>;
    const rawOptionSets: unknown[] = (() => {
      const o = optSetsRoot['optionset'] ?? optSetsRoot['OptionSet'];
      if (!o) return [];
      return Array.isArray(o) ? o : [o];
    })();
    rawOptionSets.forEach((os) => {
      const osNode     = os as Record<string, unknown>;
      const osName     = xmlStr(osNode, '@_Name') || xmlStr(osNode, 'Name');
      const osDisplayName = getLocalizedLabel(osNode, osName);
      const osDescription = xmlStr(osNode, 'Description') || xmlStr(osNode, 'description') || undefined;
      const optionsArr = (() => {
        const optsRoot = (osNode['Options'] ?? osNode['options'] ?? {}) as Record<string, unknown>;
        const opts = optsRoot['Option'] ?? optsRoot['option'];
        if (!opts) return [];
        return Array.isArray(opts) ? opts : [opts];
      })();
      const options: OptionSetOption[] = optionsArr.map((opt) => {
        const o2 = opt as Record<string, unknown>;
        const rawValue = xmlStr(o2, '@_value') || xmlStr(o2, 'value') || xmlStr(o2, '@_Value') || xmlStr(o2, 'Value');
        const value = Number(rawValue);
        const fallbackLabel = xmlStr(o2, '@_description') || rawValue;
        return {
          value:       Number.isNaN(value) ? 0 : value,
          label:       getLocalizedLabel(o2, fallbackLabel),
          description: xmlStr(o2, 'Description') || xmlStr(o2, 'description') || undefined,
          color:       xmlStr(o2, '@_color') || undefined,
          isDefault:   (xmlStr(o2, '@_default') || xmlStr(o2, 'default') || '').toLowerCase() === 'true',
        } as OptionSetOption;
      });
      optionSets.push({ name: osName, displayName: osDisplayName, description: osDescription, isGlobal: true, options } as OptionSetDefinition);
    });

    // ── Forms ─────────────────────────────────────────────────────────────
    report(35);
    const formsRoot = (root['Forms'] ?? root['forms'] ?? root['SystemForms'] ?? {}) as Record<string, unknown>;
    const rawForms: unknown[] = (() => {
      const f = formsRoot['SystemForm'] ?? formsRoot['Form'] ?? formsRoot['form'];
      if (!f) return [];
      return Array.isArray(f) ? f : [f];
    })();
    rawForms.forEach((f) => {
      const fn = f as Record<string, unknown>;
      const entityName = xmlStr(fn, 'ObjectTypeCode') || xmlStr(fn, '@_entityLogicalName') || '';
      const formType   = xmlStr(fn, 'Type') || xmlStr(fn, '@_type') || 'Main';
      const displayN   = xmlStr(fn, 'Name') || '';
      // Extract fields from FormXml/form/tabs/tab/columns/column/sections/section/rows/row/cell/control
      const fields: FormField[] = [];
      try {
        const formXmlStr = xmlStr(fn, 'FormXml') || xmlStr(fn, 'formXML') || '';
        if (formXmlStr) {
          const formDoc = xmlParser.parse(formXmlStr) as Record<string, unknown>;
          const formNode = (formDoc['form'] ?? formDoc) as Record<string, unknown>;
          type FieldContext = {
            location: 'body' | 'header' | 'footer';
            tabName?: string;
            sectionName?: string;
          };

          const contextFromNode = (
            parentContext: FieldContext,
            key: string,
            node: Record<string, unknown>,
          ): FieldContext => {
            const nextContext: FieldContext = { ...parentContext };
            const keyLower = key.toLowerCase();

            if (keyLower === 'header') {
              nextContext.location = 'header';
              nextContext.tabName = undefined;
              nextContext.sectionName = undefined;
            } else if (keyLower === 'footer') {
              nextContext.location = 'footer';
              nextContext.tabName = undefined;
              nextContext.sectionName = undefined;
            } else if (keyLower === 'tab') {
              nextContext.location = 'body';
              nextContext.tabName = xmlStr(node, '@_name') || xmlStr(node, 'name') || nextContext.tabName;
            } else if (keyLower === 'section') {
              nextContext.location = 'body';
              nextContext.sectionName = xmlStr(node, '@_name') || xmlStr(node, 'name') || nextContext.sectionName;
            }

            return nextContext;
          };

          const upsertField = (attributeName: string, context: FieldContext) => {
            const normalizedAttribute = attributeName.trim();
            if (!normalizedAttribute) return;

            const existing = fields.find((x) => x.attributeName.toLowerCase() === normalizedAttribute.toLowerCase());
            if (existing) {
              if (!existing.location) existing.location = context.location;
              if (!existing.tabName && context.tabName) existing.tabName = context.tabName;
              if (!existing.sectionName && context.sectionName) existing.sectionName = context.sectionName;
              return;
            }

            fields.push({
              attributeName: normalizedAttribute,
              location: context.location,
              tabName: context.tabName,
              sectionName: context.sectionName,
            });
          };

          const walkNode = (node: Record<string, unknown>, context: FieldContext, keyHint = 'form') => {
            const scopedContext = contextFromNode(context, keyHint, node);

            if (keyHint.toLowerCase() === 'cell') {
              const region = xmlStr(node, '@_id') || xmlStr(node, '@_name') || xmlStr(node, 'id') || '';
              const regionLower = region.toLowerCase();
              if (regionLower.includes('header')) {
                scopedContext.location = 'header';
                scopedContext.tabName = undefined;
                scopedContext.sectionName = undefined;
              } else if (regionLower.includes('footer')) {
                scopedContext.location = 'footer';
                scopedContext.tabName = undefined;
                scopedContext.sectionName = undefined;
              }
            }

            const ctrl = node['control'];
            if (ctrl) {
              const ctrls = Array.isArray(ctrl) ? ctrl : [ctrl];
              ctrls.forEach((c: unknown) => {
                const cn = c as Record<string, unknown>;
                const id = xmlStr(cn, '@_id') || xmlStr(cn, 'id');
                if (id) {
                  upsertField(id, scopedContext);
                }
              });
            }

            Object.entries(node).forEach(([childKey, value]) => {
              if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
                walkNode(value as Record<string, unknown>, scopedContext, childKey);
              } else if (Array.isArray(value)) {
                value.forEach((item) => {
                  if (typeof item === 'object' && item !== null) {
                    walkNode(item as Record<string, unknown>, scopedContext, childKey);
                  }
                });
              }
            });
          };

          walkNode(formNode, { location: 'body' });
        }
      } catch {
        // best-effort
      }
      forms.push({ name: displayN, displayName: displayN, entityLogicalName: entityName, formType, fields });
    });

    // ── Views (SavedQuery) ────────────────────────────────────────────────
    report(40);
    const savedQueriesRoot = (root['SavedQueries'] ?? root['savedqueries'] ?? {}) as Record<string, unknown>;
    const rawViews: unknown[] = (() => {
      const v = savedQueriesRoot['SavedQuery'] ?? savedQueriesRoot['savedquery'];
      if (!v) return [];
      return Array.isArray(v) ? v : [v];
    })();
    rawViews.forEach((v) => {
      const vn = v as Record<string, unknown>;
      const name       = xmlStr(vn, 'Name') || xmlStr(vn, '@_Name');
      const entityName = xmlStr(vn, 'ReturnedTypeCode') || xmlStr(vn, '@_returnedTypeCode') || '';
      const viewType   = xmlStr(vn, 'QueryType') || '0';
      const fetchXml   = xmlStr(vn, 'fetchxml') || xmlStr(vn, 'FetchXml') || undefined;

      // Extract column names from layoutxml
      const columns: string[] = [];
      try {
        const layoutXml = xmlStr(vn, 'layoutxml') || xmlStr(vn, 'LayoutXml') || '';
        if (layoutXml) {
          const layoutDoc = xmlParser.parse(layoutXml) as Record<string, unknown>;
          const grid      = (layoutDoc['grid'] ?? layoutDoc) as Record<string, unknown>;
          const row       = (grid['row'] ?? {}) as Record<string, unknown>;
          const cells     = row['cell'];
          if (cells) {
            const cellArr = Array.isArray(cells) ? cells : [cells];
            cellArr.forEach((cell: unknown) => {
              const c = cell as Record<string, unknown>;
              const colName = xmlStr(c, '@_name') || xmlStr(c, 'name');
              if (colName) columns.push(colName);
            });
          }
        }
      } catch {
        // best-effort
      }

      views.push({ name, displayName: name, entityLogicalName: entityName, viewType, columns, fetchXml });
    });

    // ── Workflows / Processes ─────────────────────────────────────────────
    report(45);
    const wfsRoot = (root['Workflows'] ?? root['workflows'] ?? {}) as Record<string, unknown>;
    const rawWfs: unknown[] = (() => {
      const w = wfsRoot['Workflow'] ?? wfsRoot['workflow'];
      if (!w) return [];
      return Array.isArray(w) ? w : [w];
    })();
    rawWfs.forEach((w) => {
      const wn = w as Record<string, unknown>;
      const uniqueName    = xmlStr(wn, 'Name') || xmlStr(wn, '@_Name') || xmlStr(wn, 'UniqueName');
      const catNum        = Number(xmlStr(wn, 'Category') || xmlStr(wn, '@_Category') || '0');
      const category      = mapProcessCategory(catNum);
      const primaryEntity = xmlStr(wn, 'PrimaryEntity') || xmlStr(wn, '@_PrimaryEntity') || undefined;
      const isActivated   = parseProcessActivationStatus(wn);
      const triggerType   = xmlStr(wn, 'TriggerOnCreate') === 'true' ? 'OnCreate'
                          : xmlStr(wn, 'TriggerOnDelete') === 'true' ? 'OnDelete'
                          : xmlStr(wn, 'TriggerOnUpdateAttribute') ? 'OnChange'
                          : 'OnDemand';
      const displayN      = (() => {
        const dn = (wn['LocalizedNames'] ?? {}) as Record<string, unknown>;
        const nn = (dn['LocalizedName'] ?? dn['localizedname']) as Record<string, unknown> | undefined;
        return nn ? (xmlStr(nn, '@_description') || xmlStr(nn, 'description')) : uniqueName;
      })();

      // Extract steps from embedded xmlnodes.xmlnode
      const xmlnodes  = (wn['XmlNodes'] ?? wn['xmlnodes'] ?? {}) as Record<string, unknown>;
      const xmlnode   = (xmlnodes['xmlnode'] ?? xmlnodes['XmlNode'] ?? {}) as Record<string, unknown>;
      const steps     = extractWorkflowSteps(xmlnode);

      processes.push({
        name:          uniqueName,
        displayName:   normalizeDisplayNameForReadability(displayN, uniqueName),
        uniqueName,
        category,
        primaryEntity,
        relatedEntities: primaryEntity ? [primaryEntity] : [],
        isActivated,
        triggerType,
        steps,
      } as ProcessDefinition);
    });

    // ── Web Resources ─────────────────────────────────────────────────────
    report(50);
    const wrRoot = (root['WebResources'] ?? root['webresources'] ?? {}) as Record<string, unknown>;
    const rawWRs: unknown[] = (() => {
      const wr = wrRoot['WebResource'] ?? wrRoot['webresource'];
      if (!wr) return [];
      return Array.isArray(wr) ? wr : [wr];
    })();
    rawWRs.forEach((wr) => {
      try {
        webResources.push(parseWebResourceNode(wr as Record<string, unknown>));
      } catch (err) {
        warnings.push(`Failed to parse web resource: ${(err as Error).message}`);
      }
    });

    // ── Security Roles ────────────────────────────────────────────────────
    report(55);
    const rolesRoot = (root['Roles'] ?? root['roles'] ?? {}) as Record<string, unknown>;
    const rawRoles: unknown[] = (() => {
      const r = rolesRoot['Role'] ?? rolesRoot['role'];
      if (!r) return [];
      return Array.isArray(r) ? r : [r];
    })();
    rawRoles.forEach((r) => {
      const rn   = r as Record<string, unknown>;
      const name = xmlStr(rn, 'Name') || xmlStr(rn, '@_name');
      const privRoot = (rn['RolePrivileges'] ?? rn['roleprivileges'] ?? {}) as Record<string, unknown>;
      const rawPrivs: unknown[] = (() => {
        const p = privRoot['RolePrivilege'] ?? privRoot['roleprivilege'];
        if (!p) return [];
        return Array.isArray(p) ? p : [p];
      })();
      const privileges: RolePrivilege[] = rawPrivs.map((p) => {
        const pn = p as Record<string, unknown>;
        const depthRaw =
          xmlStr(pn, '@_level') ||
          xmlStr(pn, 'level') ||
          xmlStr(pn, '@_depth') ||
          xmlStr(pn, 'depth') ||
          xmlStr(pn, 'AccessLevel') ||
          xmlStr(pn, '@_AccessLevel') ||
          xmlStr(pn, 'PrivilegeDepth');
        return {
          privilegeName: xmlStr(pn, '@_name') || xmlStr(pn, 'name') || xmlStr(pn, 'PrivilegeName') || xmlStr(pn, '@_PrivilegeName'),
          depth:         parseRoleDepth(depthRaw),
        };
      });
      const displayName = getLocalizedLabel(rn, name);
      securityRoles.push({ name, displayName, privileges } as SecurityRoleDefinition);
    });

    // ── Field Security Profiles ───────────────────────────────────────────
    report(57);
    const fspRoot = (root['FieldSecurityProfiles'] ?? root['fieldsecurityprofiles'] ?? {}) as Record<string, unknown>;
    const rawFSPs: unknown[] = (() => {
      const p = fspRoot['FieldSecurityProfile'] ?? fspRoot['fieldsecurityprofile'];
      if (!p) return [];
      return Array.isArray(p) ? p : [p];
    })();
    rawFSPs.forEach((fp) => {
      const fpn  = fp as Record<string, unknown>;
      const name =
        xmlStr(fpn, '@_name') ||
        xmlStr(fpn, 'Name') ||
        xmlStr(fpn, '@_Name') ||
        xmlStr(fpn, 'UniqueName') ||
        xmlStr(fpn, '@_UniqueName') ||
        xmlStr(fpn, 'LogicalName') ||
        xmlStr(fpn, '@_LogicalName') ||
        xmlStr(fpn, 'FieldSecurityProfileId');
      const displayName = getLocalizedLabel(fpn, name);
      const permsRoot = (fpn['FieldPermissions'] ?? fpn['fieldpermissions'] ?? {}) as Record<string, unknown>;
      const rawPerms: unknown[] = (() => {
        const pp = permsRoot['FieldPermission'] ?? permsRoot['fieldpermission'];
        if (!pp) return [];
        return Array.isArray(pp) ? pp : [pp];
      })();
      const permissions = rawPerms.map((pp) => {
        const ppn = pp as Record<string, unknown>;
        const attributeName =
          xmlStr(ppn, '@_attributelogicalname') ||
          xmlStr(ppn, 'attributelogicalname') ||
          xmlStr(ppn, '@_AttributeLogicalName') ||
          xmlStr(ppn, 'AttributeLogicalName') ||
          xmlStr(ppn, '@_attributename') ||
          xmlStr(ppn, 'attributename') ||
          xmlStr(ppn, '@_AttributeName') ||
          xmlStr(ppn, 'AttributeName') ||
          xmlStr(ppn, '@_field') ||
          xmlStr(ppn, 'field') ||
          xmlStr(ppn, 'FieldName') ||
          xmlStr(ppn, '@_FieldName') ||
          xmlStr(ppn, 'ColumnName');
        return {
          attributeName,
          canRead:       parsePermissionFlag(xmlStr(ppn, '@_canread') || xmlStr(ppn, 'canread') || xmlStr(ppn, 'CanRead') || xmlStr(ppn, '@_CanRead')),
          canUpdate:     parsePermissionFlag(xmlStr(ppn, '@_canupdate') || xmlStr(ppn, 'canupdate') || xmlStr(ppn, 'CanUpdate') || xmlStr(ppn, '@_CanUpdate')),
          canCreate:     parsePermissionFlag(xmlStr(ppn, '@_cancreate') || xmlStr(ppn, 'cancreate') || xmlStr(ppn, 'CanCreate') || xmlStr(ppn, '@_CanCreate')),
        };
      }).filter((perm) => !!perm.attributeName);
      fieldSecurityProfiles.push({ name, displayName, permissions } as FieldSecurityProfileDefinition);
    });

    // ── Connection References ─────────────────────────────────────────────
    report(60);
    const crRoot = (root['connectionreferences'] ?? root['ConnectionReferences'] ?? {}) as Record<string, unknown>;
    const rawCRs: unknown[] = (() => {
      const cr = crRoot['connectionreference'] ?? crRoot['ConnectionReference'];
      if (!cr) return [];
      return Array.isArray(cr) ? cr : [cr];
    })();
    rawCRs.forEach((cr) => {
      const crn = cr as Record<string, unknown>;
      const connectorId = xmlStr(crn, 'connectorid') || xmlStr(crn, 'ConnectorId');
      const connectorDisplayName =
        xmlStr(crn, 'connectorDisplayName') ||
        xmlStr(crn, 'ConnectorDisplayName') ||
        connectorDisplayFromId(connectorId);
      const displayName =
        xmlStr(crn, 'connectionreferencedisplayname') ||
        xmlStr(crn, 'ConnectionReferenceDisplayName') ||
        xmlStr(crn, 'DisplayName') ||
        getLocalizedLabel(crn, xmlStr(crn, 'Name'));
      connectionReferences.push({
        name:                   xmlStr(crn, '@_connectionreferencelogicalname') || xmlStr(crn, 'logicalname') || xmlStr(crn, 'Name'),
        displayName,
        connectorId,
        connectionId:           xmlStr(crn, 'connectionid') || undefined,
        connectorDisplayName,
      } as ConnectionReferenceDefinition);
    });

    // ── Environment Variables ─────────────────────────────────────────────
    report(62);
    const evRoot = (root['environmentvariabledefinitions'] ?? root['EnvironmentVariableDefinitions'] ?? {}) as Record<string, unknown>;
    const rawEVDs: unknown[] = (() => {
      const ev = evRoot['environmentvariabledefinition'] ?? evRoot['EnvironmentVariableDefinition'];
      if (!ev) return [];
      return Array.isArray(ev) ? ev : [ev];
    })();

    const envSchemaByDefinitionId = new Map<string, string>();
    const mapDefinitionIdToSchema = (node: Record<string, unknown>, schemaName?: string) => {
      const schema = (schemaName ?? '').trim();
      if (!schema) return;
      const definitionId = normalizeIdentifierKey(
        xmlStr(node, '@_environmentvariabledefinitionid') ||
        xmlStr(node, 'environmentvariabledefinitionid') ||
        xmlStr(node, 'EnvironmentVariableDefinitionId') ||
        xmlStr(node, '@_EnvironmentVariableDefinitionId') ||
        xmlStr(node, '@_id') ||
        xmlStr(node, 'id') ||
        xmlStr(node, 'Id'),
      );
      if (definitionId) {
        envSchemaByDefinitionId.set(definitionId, schema);
      }
    };

    const envValuesBySchema = new Map<string, string>();
    const evValuesRoot = (root['environmentvariablevalues'] ?? root['EnvironmentVariableValues'] ?? {}) as Record<string, unknown>;
    const rawEnvValues: unknown[] = (() => {
      const ev = evValuesRoot['environmentvariablevalue'] ?? evValuesRoot['EnvironmentVariableValue'];
      if (!ev) return [];
      return Array.isArray(ev) ? ev : [ev];
    })();
    rawEnvValues.forEach((ev) => {
      const evn = ev as Record<string, unknown>;
      const definitionId = normalizeIdentifierKey(
        xmlStr(evn, 'environmentvariabledefinitionid') ||
        xmlStr(evn, 'EnvironmentVariableDefinitionId') ||
        xmlStr(evn, '@_environmentvariabledefinitionid') ||
        xmlStr(evn, '@_EnvironmentVariableDefinitionId'),
      );
      const schemaName =
        xmlStr(evn, 'schemaname') ||
        xmlStr(evn, '@_schemaname') ||
        xmlStr(evn, 'EnvironmentVariableDefinitionSchemaName') ||
        xmlStr(evn, 'environmentvariabledefinitionschemaname') ||
        xmlStr(evn, 'Name') ||
        envSchemaByDefinitionId.get(definitionId) ||
        '';
      const value =
        xmlStr(evn, 'value') ||
        xmlStr(evn, 'Value') ||
        xmlStr(evn, '@_value') ||
        xmlStr(evn, 'CurrentValue') ||
        xmlStr(evn, 'currentvalue');
      if (schemaName && value) envValuesBySchema.set(schemaName, value);
    });

    rawEVDs.forEach((ev) => {
      const evn = ev as Record<string, unknown>;
      const schemaName =
        xmlStr(evn, 'schemaname') ||
        xmlStr(evn, '@_schemaname') ||
        xmlStr(evn, '@_SchemaName');
      mapDefinitionIdToSchema(evn, schemaName);
      const currentValue = envValuesBySchema.get(schemaName);
      environmentVariables.push({
        name:          schemaName || xmlStr(evn, 'Name'),
        schemaName,
        displayName:   environmentVariableDisplayName(evn, schemaName),
        description:   xmlStr(evn, 'description') || xmlStr(evn, 'Description') || undefined,
        type:          xmlStr(evn, 'type') || xmlStr(evn, 'Type') || 'String',
        defaultValue:  xmlStr(evn, 'defaultvalue') || undefined,
        hasCurrentValue: !!currentValue,
        currentValue,
      } as EnvironmentVariableDefinition);
    });

    // ── Email Templates ─────────────────────────────────────────────────────
    report(63);
    const emailRoot = (root['EmailTemplates'] ?? root['emailtemplates'] ?? root['Templates'] ?? root['templates'] ?? {}) as Record<string, unknown>;
    const rawEmailTemplates: unknown[] = (() => {
      const tpl = emailRoot['EmailTemplate'] ?? emailRoot['emailtemplate'] ?? emailRoot['Template'] ?? emailRoot['template'];
      if (!tpl) return [];
      return Array.isArray(tpl) ? tpl : [tpl];
    })();
    rawEmailTemplates.forEach((tpl) => {
      const tn = tpl as Record<string, unknown>;
      const name =
        xmlStr(tn, 'Name') ||
        xmlStr(tn, '@_Name') ||
        xmlStr(tn, 'Title') ||
        xmlStr(tn, '@_Title') ||
        xmlStr(tn, 'TemplateId');
      emailTemplates.push({
        name,
        displayName: getLocalizedLabel(tn, xmlStr(tn, 'DisplayName') || name),
        description: xmlStr(tn, 'Description') || undefined,
        subject: xmlStr(tn, 'Subject') || xmlStr(tn, '@_Subject') || undefined,
        entityLogicalName: xmlStr(tn, 'ObjectTypeCode') || xmlStr(tn, 'EntityName') || xmlStr(tn, '@_ObjectTypeCode') || undefined,
        templateType: xmlStr(tn, 'TemplateType') || xmlStr(tn, 'templatetype') || undefined,
        languageCode: xmlStr(tn, 'LanguageCode') || xmlStr(tn, '@_languagecode') || undefined,
      } as EmailTemplateDefinition);
    });

    // ── Dashboards ────────────────────────────────────────────────────────
    report(64);
    const dbRoot = (root['Dashboards'] ?? root['dashboards'] ?? {}) as Record<string, unknown>;
    const rawDBs: unknown[] = (() => {
      const d = dbRoot['Dashboard'] ?? dbRoot['SystemForm'];
      if (!d) return [];
      return Array.isArray(d) ? d : [d];
    })();
    rawDBs.forEach((d) => {
      const dn = d as Record<string, unknown>;
      const formXml = xmlStr(dn, 'FormXml') || xmlStr(dn, 'formxml') || '';
      dashboards.push({
        name:          xmlStr(dn, 'Name') || xmlStr(dn, '@_Name') || xmlStr(dn, 'FormId') || 'Dashboard',
        displayName:   getLocalizedLabel(dn, xmlStr(dn, 'DisplayName') || xmlStr(dn, 'Name') || xmlStr(dn, '@_Name') || 'Dashboard'),
        entityLogicalName: xmlStr(dn, 'ObjectTypeCode') || xmlStr(dn, '@_ObjectTypeCode') || undefined,
        dashboardType: xmlStr(dn, 'Type') || xmlStr(dn, 'DashboardType') || 'Standard',
        components:    parseDashboardComponents(formXml),
      } as DashboardDefinition);
    });

    // ── Reports ───────────────────────────────────────────────────────────
    report(66);
    const rptRoot = (root['Reports'] ?? root['reports'] ?? {}) as Record<string, unknown>;
    const rawRpts: unknown[] = (() => {
      const rp = rptRoot['Report'] ?? rptRoot['report'];
      if (!rp) return [];
      return Array.isArray(rp) ? rp : [rp];
    })();
    rawRpts.forEach((rp) => {
      const rpn = rp as Record<string, unknown>;
      reports.push({
        name:             xmlStr(rpn, 'Name') || xmlStr(rpn, '@_Name') || xmlStr(rpn, 'ReportId') || 'Report',
        displayName:      getLocalizedLabel(rpn, xmlStr(rpn, 'DisplayName') || xmlStr(rpn, 'Name') || xmlStr(rpn, '@_Name') || 'Report'),
        fileName:         xmlStr(rpn, 'FileName') || undefined,
        relatedEntities:  [xmlStr(rpn, 'ObjectTypeCode') || xmlStr(rpn, '@_ObjectTypeCode')].filter(Boolean),
        category:         xmlStr(rpn, 'ReportTypeCode') || xmlStr(rpn, 'Category') || undefined,
      } as ReportDefinition);
    });

    // ── App Modules ───────────────────────────────────────────────────────
    report(68);
    const amRoot = (root['AppModules'] ?? root['appmodules'] ?? {}) as Record<string, unknown>;
    const rawAMs: unknown[] = (() => {
      const am = amRoot['AppModule'] ?? amRoot['appmodule'];
      if (!am) return [];
      return Array.isArray(am) ? am : [am];
    })();
    const knownEntitiesForApps = new Set(entities.map((entity) => entity.logicalName.toLowerCase()));
    rawAMs.forEach((am) => {
      const amn = am as Record<string, unknown>;
      const uniqueName = xmlStr(amn, 'UniqueName') || xmlStr(amn, '@_UniqueName') || xmlStr(amn, 'Name') || xmlStr(amn, '@_Name');
      const displayName = getLocalizedLabel(amn, xmlStr(amn, 'Name') || xmlStr(amn, '@_Name') || uniqueName);
      const appEntities = new Set<string>();
      extractReferencedEntities(amn, knownEntitiesForApps, appEntities);

      const appConnectors = new Set<string>();
      extractConnectorNames(amn, appConnectors);

      const sitemapAreasSet = new Set<string>();
      const siteMapEntities = new Set<string>();
      const siteMapStructure: AppSiteMapArea[] = [];
      let siteMapSettings: AppSiteMapSettings | undefined;
      const siteMapSources = [
        xmlStr(amn, 'SiteMapXml'),
        xmlStr(amn, 'sitemapxml'),
        xmlStr(amn, 'AppModuleXml'),
        xmlStr(amn, 'appmodulexml'),
      ].filter((value) => !!value);

      siteMapSources.forEach((source) => {
        const parsed = parseSiteMapStructure(source);
        parsed.areas.forEach((area) => sitemapAreasSet.add(area));
        parsed.entities.forEach((entity) => siteMapEntities.add(entity));
        parsed.structure.forEach((area) => siteMapStructure.push(area));
        if (parsed.settings) {
          siteMapSettings = {
            showHome: siteMapSettings?.showHome ?? parsed.settings.showHome,
            showPinned: siteMapSettings?.showPinned ?? parsed.settings.showPinned,
            showRecents: siteMapSettings?.showRecents ?? parsed.settings.showRecents,
            enableCollapsibleGroups: siteMapSettings?.enableCollapsibleGroups ?? parsed.settings.enableCollapsibleGroups,
          };
        }
      });

      const siteMapUniqueName = xmlStr(amn, 'SiteMapUniqueName') || xmlStr(amn, 'sitemapuniquename');
      if (siteMapUniqueName) sitemapAreasSet.add(siteMapUniqueName);
      siteMapEntities.forEach((entity) => appEntities.add(entity));

      apps.push({
        name:         uniqueName,
        displayName,
        appType:      AppType.ModelDriven,
        uniqueName,
        isEnabled:    xmlStr(amn, 'IsEnabled') !== 'false',
        version:      xmlStr(amn, 'ClientVersion') || xmlStr(amn, 'Version') || undefined,
        entities:     Array.from(appEntities),
        sitemapAreas: Array.from(sitemapAreasSet),
        siteMap:      siteMapStructure,
        siteMapSettings,
        connectors:   Array.from(appConnectors),
      } as AppDefinition);
    });
  }

  // ── Step 3b: Merge env vars from folder exports ───────────────────────
  report(69);
  const envVarKey = (schemaName?: string, name?: string): string =>
    normalizeIdentifierKey(schemaName || name);

  const envSchemaByDefinitionId = new Map<string, string>();

  const mapDefinitionIdToSchema = (node: Record<string, unknown>, schemaName?: string) => {
    const schema = (schemaName ?? '').trim();
    if (!schema) return;
    const definitionId = normalizeIdentifierKey(
      xmlStr(node, '@_environmentvariabledefinitionid') ||
      xmlStr(node, 'environmentvariabledefinitionid') ||
      xmlStr(node, 'EnvironmentVariableDefinitionId') ||
      xmlStr(node, '@_EnvironmentVariableDefinitionId') ||
      xmlStr(node, '@_id') ||
      xmlStr(node, 'id') ||
      xmlStr(node, 'Id'),
    );
    if (definitionId) {
      envSchemaByDefinitionId.set(definitionId, schema);
    }
  };

  const mergeEnvironmentVariable = (incoming: EnvironmentVariableDefinition) => {
    const key = envVarKey(incoming.schemaName, incoming.name);
    if (!key) return;

    const existingIdx = environmentVariables.findIndex((item) => envVarKey(item.schemaName, item.name) === key);
    if (existingIdx < 0) {
      environmentVariables.push(incoming);
      return;
    }

    const existing = environmentVariables[existingIdx];
    environmentVariables[existingIdx] = {
      ...existing,
      ...incoming,
      name: existing.name || incoming.name,
      schemaName: existing.schemaName || incoming.schemaName,
      displayName: existing.displayName || incoming.displayName,
      description: existing.description || incoming.description,
      type: existing.type || incoming.type,
      defaultValue: existing.defaultValue || incoming.defaultValue,
      hasCurrentValue: existing.hasCurrentValue || incoming.hasCurrentValue,
      currentValue: existing.currentValue || incoming.currentValue,
    };
  };

  const envDefinitionEntries = getEntriesWithPrefix(zip, 'environmentvariabledefinitions/');
  for (const [path, entry] of envDefinitionEntries) {
    if (
      entry.dir ||
      !path.toLowerCase().endsWith('/environmentvariabledefinition.xml')
    ) {
      continue;
    }

    try {
      const xml = await entry.async('string');
      if (!xml.trim()) continue;
      const doc = xmlParser.parse(xml) as Record<string, unknown>;
      const node =
        firstObject(doc['environmentvariabledefinition']) ||
        firstObject(doc['EnvironmentVariableDefinition']) ||
        doc;

      const schemaName =
        xmlStr(node, 'schemaname') ||
        xmlStr(node, '@_schemaname') ||
        xmlStr(node, '@_SchemaName');
      mapDefinitionIdToSchema(node, schemaName);
      const defaultValue =
        xmlStr(node, 'defaultvalue') ||
        xmlStr(node, 'defaultValue') ||
        xmlStr(node, 'DefaultValue') ||
        undefined;

      mergeEnvironmentVariable({
        name: schemaName || xmlStr(node, 'Name') || path.split('/').slice(-2)[0],
        schemaName,
        displayName: environmentVariableDisplayName(node, schemaName),
        description: xmlStr(node, 'description') || xmlStr(node, 'Description') || undefined,
        type: xmlStr(node, 'type') || xmlStr(node, 'Type') || 'String',
        defaultValue,
        hasCurrentValue: false,
      } as EnvironmentVariableDefinition);
    } catch {
      warnings.push(`Could not parse environment variable definition: ${path}`);
    }
  }

  const envValueEntries = getEntriesWithPrefix(zip, 'environmentvariablevalues/');
  for (const [path, entry] of envValueEntries) {
    if (entry.dir || !path.toLowerCase().endsWith('.xml')) continue;

    try {
      const xml = await entry.async('string');
      if (!xml.trim()) continue;
      const doc = xmlParser.parse(xml) as Record<string, unknown>;
      const node =
        firstObject(doc['environmentvariablevalue']) ||
        firstObject(doc['EnvironmentVariableValue']) ||
        doc;

      const schemaName =
        xmlStr(node, 'schemaname') ||
        xmlStr(node, '@_schemaname') ||
        xmlStr(node, '@_SchemaName') ||
        envSchemaByDefinitionId.get(normalizeIdentifierKey(
          xmlStr(node, 'environmentvariabledefinitionid') ||
          xmlStr(node, 'EnvironmentVariableDefinitionId') ||
          xmlStr(node, '@_environmentvariabledefinitionid') ||
          xmlStr(node, '@_EnvironmentVariableDefinitionId'),
        )) ||
        path.split('/').pop()?.replace(/\.xml$/i, '');
      const value =
        xmlStr(node, 'value') ||
        xmlStr(node, 'Value') ||
        xmlStr(node, '@_value') ||
        xmlStr(node, 'CurrentValue') ||
        xmlStr(node, 'currentvalue');
      if (!schemaName || !value) continue;

      mergeEnvironmentVariable({
        name: schemaName,
        schemaName,
        displayName: schemaName,
        type: 'String',
        hasCurrentValue: true,
        currentValue: value,
      } as EnvironmentVariableDefinition);
    } catch {
      warnings.push(`Could not parse environment variable value: ${path}`);
    }
  }

  // ── Step 3b: Additional web resources from WebResources/ folder ─────────
  // Power Platform can also store web resources as individual files in a
  // WebResources/ folder within the solution ZIP.  Supplement what was found
  // in customizations.xml with anything discovered here.
  {
    const knownWrNames = new Set(webResources.map((wr) => (wr.schemaName || wr.name).toLowerCase()));

    /** Derive a WebResourceType from a file extension. */
    const typeFromExtension = (ext: string): WebResourceType => {
      const map: Record<string, WebResourceType> = {
        html: WebResourceType.HTML,
        htm:  WebResourceType.HTML,
        css:  WebResourceType.CSS,
        js:   WebResourceType.JavaScript,
        ts:   WebResourceType.TypeScript,
        xml:  WebResourceType.XML,
        png:  WebResourceType.PNG,
        jpg:  WebResourceType.JPG,
        jpeg: WebResourceType.JPG,
        gif:  WebResourceType.GIF,
        xap:  WebResourceType.XAP,
        xsl:  WebResourceType.XSL,
        xslt: WebResourceType.XSL,
        ico:  WebResourceType.ICO,
        svg:  WebResourceType.SVG,
        resx: WebResourceType.Resx,
      };
      return map[ext.toLowerCase()] ?? WebResourceType.Unknown;
    };

    for (const [path, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue;
      const lowerPath = path.toLowerCase();
      if (!lowerPath.startsWith('webresources/')) continue;

      const fileName = path.split('/').pop() ?? path;
      const extMatch = fileName.match(/\.([^.]+)$/);
      const ext = extMatch?.[1] ?? '';
      const resourceType = typeFromExtension(ext);

      // Use full path as schema name so we can deduplicate against customizations.xml entries
      const schemaName = path.replace(/^WebResources\//i, '');
      const schemaKey  = schemaName.toLowerCase();
      const nameKey    = fileName.toLowerCase();

      if (knownWrNames.has(schemaKey) || knownWrNames.has(nameKey)) continue;

      let content: string | undefined;
      let contentLength: number | undefined;
      try {
        if ([WebResourceType.JavaScript, WebResourceType.TypeScript, WebResourceType.HTML,
             WebResourceType.CSS, WebResourceType.XML, WebResourceType.XSL,
             WebResourceType.SVG, WebResourceType.Resx].includes(resourceType)) {
          content = await entry.async('string');
          contentLength = content.length;
        } else {
          const buf = await entry.async('uint8array');
          contentLength = buf.byteLength;
        }
      } catch {
        // best-effort — skip content if unreadable
      }

      knownWrNames.add(schemaKey);
      webResources.push({
        name:             fileName,
        schemaName,
        displayName:      stripTrailingGuid(fileName.replace(/\.[^.]+$/, '')) || fileName,
        resourceType,
        enabledForMobile: undefined,
        availableOffline: undefined,
        content,
        contentLength,
      } as WebResourceDefinition);
    }
  }

  // ── Step 4: Canvas apps from CanvasApps/ folder ─────────────────────────
  report(70);
  const canvasEntries = getEntriesWithPrefix(zip, 'CanvasApps/');
  const discoveredCanvasApps = new Set<string>();

  // Actual canvas/custom-page packages are typically exported as *.msapp
  for (const [path, entry] of canvasEntries) {
    if (path.endsWith('.msapp')) {
      const appName = path.split('/').pop()?.replace(/\.(msapp|json)$/, '') ?? path;
      const cleanName = stripTrailingGuid(appName);
      const displayName = humanizeIdentifier(cleanName || appName);
      let canvasInsights: CanvasAppInsights | undefined;

      try {
        const msappBytes = await entry.async('uint8array');
        const msappZip = await JSZip.loadAsync(msappBytes);
        let mergedInsights: CanvasAppInsights | undefined;

        for (const [innerPath, innerEntry] of Object.entries(msappZip.files)) {
          if (innerEntry.dir) continue;

          const lowerPath = innerPath.toLowerCase();

          if (lowerPath.endsWith('.json')) {
            try {
              const json = await innerEntry.async('string');
              if (!json.trim()) continue;
              const parsed = JSON.parse(json) as Record<string, unknown>;
              mergedInsights = mergeCanvasInsights(mergedInsights, extractCanvasAppInsights(parsed));
            } catch {
              // best-effort parsing for .msapp internals
            }
            continue;
          }

          if (lowerPath.endsWith('.fx.yaml') || lowerPath.endsWith('.yaml') || lowerPath.endsWith('.yml')) {
            try {
              const yamlText = await innerEntry.async('string');
              if (!yamlText.trim()) continue;
              mergedInsights = mergeCanvasInsights(mergedInsights, extractCanvasInsightsFromText(yamlText));
            } catch {
              // best-effort parsing for yaml internals
            }
          }
        }
        canvasInsights = mergedInsights;
      } catch {
        // best-effort only
      }

      // Avoid duplicates with model-driven apps already found
      const existingIdx = apps.findIndex((a) => a.uniqueName.toLowerCase() === appName.toLowerCase());
      if (existingIdx >= 0) {
        apps[existingIdx] = mergeAppDefinition(apps[existingIdx], {
          name: appName,
          displayName,
          appType: AppType.Canvas,
          uniqueName: appName,
          isEnabled: true,
          entities: [],
          connectors: [],
          canvasInsights,
        } as AppDefinition);
      } else {
        apps.push({
          name:        appName,
          displayName,
          appType:     AppType.Canvas,
          uniqueName:  appName,
          isEnabled:   true,
          entities:    [],
          connectors:  [],
          canvasInsights,
        } as AppDefinition);
      }
      discoveredCanvasApps.add(appName.toLowerCase());
    }
  }

  for (const [path, entry] of canvasEntries) {
    if (!path.endsWith('.json')) continue;
    try {
      const jsonStr = await entry.async('string');
      const appType = classifyCanvasAppType(jsonStr);
      if (!appType) continue;

      const metadata = JSON.parse(jsonStr) as Record<string, unknown>;
      const knownCanvasEntities = new Set(entities.map((entity) => entity.logicalName.toLowerCase()));
      const rawName =
        xmlStr(metadata, 'name') ||
        xmlStr(metadata, 'displayName') ||
        xmlStr((metadata['properties'] ?? {}) as Record<string, unknown>, 'displayName') ||
        xmlStr((metadata['properties'] ?? {}) as Record<string, unknown>, 'name') ||
        path.split('/').slice(-2)[0] ||
        path.split('/').pop()?.replace(/\.json$/, '') ||
        path;
      const uniqueName = stripTrailingGuid(rawName) || rawName;

      const appEntities = new Set<string>();
      extractReferencedEntities(metadata, knownCanvasEntities, appEntities);

      const appConnectors = new Set<string>();
      extractConnectorNames(metadata, appConnectors);
      const canvasInsights = extractCanvasAppInsights(metadata);

      const version =
        xmlStr(metadata, 'version') ||
        xmlStr((metadata['properties'] ?? {}) as Record<string, unknown>, 'version') ||
        xmlStr((metadata['publishInfo'] ?? {}) as Record<string, unknown>, 'version') ||
        undefined;

      const existingIdx = apps.findIndex((a) => a.uniqueName.toLowerCase() === uniqueName.toLowerCase());
      if (existingIdx >= 0) {
        apps[existingIdx] = mergeAppDefinition(apps[existingIdx], {
          name: uniqueName,
          displayName: humanizeIdentifier(uniqueName),
          appType,
          uniqueName,
          isEnabled: true,
          version,
          entities: Array.from(appEntities),
          connectors: Array.from(appConnectors),
          canvasInsights,
        } as AppDefinition);
      } else {
        apps.push({
          name: uniqueName,
          displayName: humanizeIdentifier(uniqueName),
          appType,
          uniqueName,
          isEnabled: true,
          version,
          entities: Array.from(appEntities),
          connectors: Array.from(appConnectors),
          canvasInsights,
        } as AppDefinition);
      }
      discoveredCanvasApps.add(uniqueName.toLowerCase());
    } catch {
      // best-effort only
    }
  }

  const appModuleEntries = getEntriesWithPrefix(zip, 'AppModules/');
  for (const [path, entry] of appModuleEntries) {
    if (entry.dir || (!path.toLowerCase().endsWith('.xml') && !path.toLowerCase().endsWith('.json'))) continue;

    try {
      const raw = await entry.async('string');
      if (!raw.trim()) continue;

      const parsedStructured = readStructuredContent(raw, path);
      const appNameFromData = parsedStructured
        ? findFirstStringByKey(parsedStructured, new Set(['uniquename', 'name', 'appuniquename']))
        : undefined;
      const displayNameFromData = parsedStructured
        ? findFirstStringByKey(parsedStructured, new Set(['displayname', 'title', 'appname']))
        : undefined;
      const versionFromData = parsedStructured
        ? findFirstStringByKey(parsedStructured, new Set(['clientversion', 'version']))
        : undefined;

      const appName = stripTrailingGuid(appNameFromData || path.split('/').slice(-2)[0] || path.split('/').pop()?.replace(/\.[^.]+$/i, '') || path);
      const displayName = displayNameFromData || humanizeIdentifier(appName);

      const { areas: sitemapAreas, entities: sitemapEntities, structure: siteMapStructure, settings: siteMapSettings } = parseSiteMapStructure(raw);
      const appConnectors = new Set<string>();
      if (parsedStructured) extractConnectorNames(parsedStructured, appConnectors);

      const incoming: AppDefinition = {
        name: appName,
        displayName,
        appType: AppType.ModelDriven,
        uniqueName: appName,
        isEnabled: true,
        version: versionFromData || undefined,
        entities: sitemapEntities,
        sitemapAreas: sitemapAreas,
        siteMap: siteMapStructure,
        siteMapSettings,
        connectors: Array.from(appConnectors),
      };

      const existingIdx = apps.findIndex((app) => app.uniqueName.toLowerCase() === appName.toLowerCase());
      if (existingIdx >= 0) {
        apps[existingIdx] = mergeAppDefinition(apps[existingIdx], incoming);
      } else {
        apps.push(incoming);
      }

      if (versionFromData) {
        const idx = apps.findIndex((app) => app.uniqueName.toLowerCase() === appName.toLowerCase());
        if (idx >= 0 && !apps[idx].version) {
          apps[idx].version = versionFromData;
        }
      }
    } catch {
      warnings.push(`Could not parse app module artifact: ${path}`);
    }
  }

  for (const folderPrefix of ['AIPlugins/', 'CodeApps/']) {
    const entries = getEntriesWithPrefix(zip, folderPrefix);
    for (const [path] of entries) {
      if (path.endsWith('/')) continue;
      const baseName = stripTrailingGuid(path.split('/').pop()?.replace(/\.[^.]+$/, '') ?? path);
      if (!baseName) continue;
      if (apps.some((a) => a.uniqueName.toLowerCase() === baseName.toLowerCase())) continue;
      apps.push({
        name: baseName,
        displayName: humanizeIdentifier(baseName),
        appType: AppType.CodeApp,
        uniqueName: baseName,
        isEnabled: true,
        entities: [],
        connectors: [],
      } as AppDefinition);
    }
  }

  // ── Step 4b: Agents / AI models / Desktop flows from solution folders ──
  report(72);
  const moduleEntries = Object.entries(zip.files)
    .filter(([path, entry]) => !entry.dir && isModuleTextFile(path));

  const agentNameKeys = new Set(['name', 'displayname', 'agentname', 'botname', 'title']);
  const aiModelNameKeys = new Set(['name', 'displayname', 'modelname', 'title']);
  const typeKeys = new Set(['type', 'agenttype', 'modeltype', 'category']);
  const languageKeys = new Set(['language', 'locale', 'languagecode']);
  const triggerKeys = new Set(['trigger', 'triggertype', 'channel', 'entrypoint']);
  const providerKeys = new Set(['provider', 'vendor', 'publisher']);
  const endpointKeys = new Set(['endpoint', 'deployment', 'deploymentname', 'url']);
  const versionKeys = new Set(['version', 'modelversion']);
  const connectorKeys = new Set(['connector', 'connectors', 'connectorid', 'connectionreference', 'connectionreferences']);
  const enabledKeys = new Set(['enabled', 'isenabled', 'active', 'isactive']);
  const stepKeys = new Set(['steps', 'actions', 'blocks']);

  for (const [path, entry] of moduleEntries) {
    const lowerPath = path.toLowerCase();
    const looksLikeAgent =
      lowerPath.includes('/agents/') ||
      lowerPath.startsWith('agents/') ||
      (lowerPath.includes('copilot') && lowerPath.includes('agent')) ||
      lowerPath.includes('/botcomponents/');
    const looksLikeAIModel =
      lowerPath.startsWith('aimodels/') ||
      lowerPath.includes('/aimodel') ||
      lowerPath.includes('/ai-model') ||
      lowerPath.includes('/ai_models/') ||
      lowerPath.includes('/models/ai/');
    const looksLikeDesktopFlow =
      lowerPath.includes('/desktopflows/') ||
      lowerPath.includes('/desktopflow/') ||
      lowerPath.includes('desktop-flow') ||
      lowerPath.includes('desktopflow') ||
      lowerPath.includes('/uiflows/');
    const looksLikeDataflow =
      lowerPath.startsWith('dataflows/') ||
      lowerPath.includes('/dataflows/') ||
      lowerPath.includes('/dataflow/');
    const looksLikeCustomApi =
      lowerPath.startsWith('customapis/') ||
      lowerPath.includes('/customapis/') ||
      lowerPath.includes('/customapi/') ||
      lowerPath.includes('/custom-api/');
    const looksLikeOfflineProfile =
      lowerPath.startsWith('mobileofflineprofiles/') ||
      lowerPath.includes('/mobileofflineprofiles/') ||
      lowerPath.includes('/offlineprofiles/') ||
      lowerPath.includes('/offlineprofile/');

    if (!looksLikeAgent && !looksLikeAIModel && !looksLikeDesktopFlow && !looksLikeDataflow && !looksLikeCustomApi && !looksLikeOfflineProfile) continue;

    let content: string;
    try {
      content = await entry.async('string');
    } catch {
      warnings.push(`Could not read module artifact: ${path}`);
      continue;
    }

    const parsed = readStructuredContent(content, path);
    const baseName = stripTrailingGuid(path.split('/').pop()?.replace(/\.[^.]+$/i, '') || path);

    if (looksLikeAgent) {
      const connectorSet = new Set<string>();
      if (parsed) collectStringsByKey(parsed, connectorKeys, connectorSet);
      const agentName =
        (parsed ? findFirstStringByKey(parsed, agentNameKeys) : undefined) ||
        humanizeIdentifier(baseName);
      const agentType = parsed ? findFirstStringByKey(parsed, typeKeys) : undefined;
      const language = parsed ? findFirstStringByKey(parsed, languageKeys) : undefined;
      const trigger = parsed ? findFirstStringByKey(parsed, triggerKeys) : undefined;
      if (!agents.some((agent) => agent.sourcePath.toLowerCase() === path.toLowerCase())) {
        agents.push({
          name: baseName,
          displayName: agentName,
          sourcePath: path,
          agentType: agentType || undefined,
          language: language || undefined,
          trigger: trigger || undefined,
          connectors: uniqueStrings(Array.from(connectorSet)),
        } as AgentDefinition);
      }
    }

    if (looksLikeAIModel) {
      const modelName =
        (parsed ? findFirstStringByKey(parsed, aiModelNameKeys) : undefined) ||
        humanizeIdentifier(baseName);
      const modelType = parsed ? findFirstStringByKey(parsed, typeKeys) : undefined;
      const provider = parsed ? findFirstStringByKey(parsed, providerKeys) : undefined;
      const endpoint = parsed ? findFirstStringByKey(parsed, endpointKeys) : undefined;
      const version = parsed ? findFirstStringByKey(parsed, versionKeys) : undefined;
      if (!aiModels.some((model) => model.sourcePath.toLowerCase() === path.toLowerCase())) {
        aiModels.push({
          name: baseName,
          displayName: modelName,
          sourcePath: path,
          modelType: modelType || undefined,
          provider: provider || undefined,
          endpoint: endpoint || undefined,
          version: version || undefined,
        } as AIModelDefinition);
      }
    }

    if (looksLikeDesktopFlow) {
      const connectorSet = new Set<string>();
      if (parsed) collectStringsByKey(parsed, connectorKeys, connectorSet);

      const flowName =
        (parsed ? findFirstStringByKey(parsed, new Set(['name', 'displayname', 'flowname', 'title'])) : undefined) ||
        humanizeIdentifier(baseName);
      const folder = parsed ? findFirstStringByKey(parsed, new Set(['folder', 'group', 'category'])) : undefined;
      const enabledRaw = parsed ? findFirstStringByKey(parsed, enabledKeys) : undefined;
      const isEnabled = enabledRaw
        ? (['true', '1', 'yes', 'enabled', 'active'].includes(enabledRaw.toLowerCase())
          ? true
          : ['false', '0', 'no', 'disabled', 'inactive'].includes(enabledRaw.toLowerCase())
            ? false
            : undefined)
        : undefined;
      const stepCount = parsed
        ? findArraySizeByKey(parsed, stepKeys)
        : (content.match(/<(step|action)\b/gi)?.length || undefined);

      if (!desktopFlows.some((flow) => flow.sourcePath.toLowerCase() === path.toLowerCase())) {
        desktopFlows.push({
          name: baseName,
          displayName: flowName,
          sourcePath: path,
          folder: folder || undefined,
          isEnabled,
          stepCount,
          connectors: uniqueStrings(Array.from(connectorSet)),
        } as DesktopFlowDefinition);
      }
    }

    if (looksLikeDataflow) {
      const connectorSet = new Set<string>();
      if (parsed) collectStringsByKey(parsed, connectorKeys, connectorSet);

      const displayName =
        (parsed ? findFirstStringByKey(parsed, new Set(['name', 'displayname', 'title', 'dataflowname'])) : undefined) ||
        humanizeIdentifier(baseName);
      const refreshMode = parsed
        ? findFirstStringByKey(parsed, new Set(['refreshmode', 'refresh', 'refreshtype', 'schedule']))
        : undefined;

      if (!dataflows.some((flow) => flow.sourcePath.toLowerCase() === path.toLowerCase())) {
        dataflows.push({
          name: baseName,
          displayName,
          sourcePath: path,
          connectors: uniqueStrings(Array.from(connectorSet)),
          refreshMode: refreshMode || undefined,
        } as DataflowDefinition);
      }
    }

    if (looksLikeCustomApi) {
      const displayName =
        (parsed ? findFirstStringByKey(parsed, new Set(['name', 'displayname', 'uniquename', 'title'])) : undefined) ||
        humanizeIdentifier(baseName);
      const boundEntityLogicalName = parsed
        ? findFirstStringByKey(parsed, new Set(['boundentitylogicalname', 'boundentity', 'entitylogicalname']))
        : undefined;
      const isFunctionRaw = parsed
        ? findFirstStringByKey(parsed, new Set(['isfunction', 'function', 'isfunctionapi']))
        : undefined;
      const isFunction = isFunctionRaw
        ? parseBooleanLike(isFunctionRaw)
        : undefined;

      if (!customApis.some((api) => api.sourcePath.toLowerCase() === path.toLowerCase())) {
        customApis.push({
          name: baseName,
          displayName,
          sourcePath: path,
          boundEntityLogicalName: boundEntityLogicalName || undefined,
          isFunction,
        } as CustomAPIDefinition);
      }
    }

    if (looksLikeOfflineProfile) {
      const profileName =
        (parsed ? findFirstStringByKey(parsed, new Set(['name', 'displayname', 'profilename', 'title'])) : undefined) ||
        humanizeIdentifier(baseName);
      const profileType = parsed
        ? findFirstStringByKey(parsed, new Set(['profiletype', 'type', 'category']))
        : undefined;
      const entitiesSet = new Set<string>();
      if (parsed) collectStringsByKey(parsed, new Set(['entity', 'entityname', 'logicalname', 'table']), entitiesSet);

      if (!offlineProfiles.some((profile) => profile.sourcePath.toLowerCase() === path.toLowerCase())) {
        offlineProfiles.push({
          name: baseName,
          displayName: profileName,
          sourcePath: path,
          profileType: profileType || undefined,
          entities: uniqueStrings(Array.from(entitiesSet)),
        } as OfflineProfileDefinition);
      }
    }
  }

  // ── Step 5: Power Automate flows from Workflows/ folder ─────────────────
  report(75);
  const workflowEntries = getEntriesWithPrefix(zip, 'Workflows/');
  const flowJsonMap = new Map<string, Record<string, unknown>>();

  for (const [path, entry] of workflowEntries) {
    if (path.endsWith('.json')) {
      try {
        const jsonStr  = await entry.async('string');
        const flowJson = JSON.parse(jsonStr) as Record<string, unknown>;
        // Normalise file name: strip curly braces that PP sometimes wraps around GUIDs
        const rawName = path.split('/').pop()?.replace(/\.json$/, '') ?? path;
        const flowName = rawName.replace(/^\{|\}$/g, '');
        flowJsonMap.set(flowName, flowJson);
      } catch {
        warnings.push(`Could not parse flow JSON: ${path}`);
      }
    }
  }

  // Match flow JSON to existing process definitions (by name similarity)
  // or create new Power Automate process entries
  const connectionRefCandidates = Array.from(new Set(connectionReferences.flatMap((cr) => [cr.name, cr.displayName].filter((value): value is string => !!value))));
  const envVarCandidates = Array.from(new Set(environmentVariables.flatMap((ev) => [ev.schemaName, ev.displayName].filter((value): value is string => !!value))));

  /** Normalise a process name for matching: lowercase + strip curly braces + strip trailing GUID. */
  const normForMatch = (s: string) => stripTrailingGuid(s.replace(/^\{|\}$/g, '')).toLowerCase();

  flowJsonMap.forEach((flowJson, flowName) => {
    const { steps, connectors, triggerDescription, displayName: jsonDisplayName, triggerEntity } = parseFlowDefinition(flowJson);
    const usedConnectionRefs = findFlowMatches(flowJson, connectionRefCandidates);
    const usedEnvVars = findFlowMatches(flowJson, envVarCandidates);

    // Prefer the display name embedded in the JSON properties (most human-readable)
    const flowNameWithoutGuid = stripTrailingGuid(flowName);
    const resolvedDisplayName = normalizeDisplayNameForReadability(
      jsonDisplayName,
      flowNameWithoutGuid || flowName,
    ) || humanizeIdentifier(flowNameWithoutGuid || flowName);

    const existingIdx = processes.findIndex((p) => {
      const pUnique        = p.uniqueName.toLowerCase();
      const pUniqueTrimmed = stripTrailingGuid(p.uniqueName).toLowerCase();
      const pUniqueNorm    = normForMatch(p.uniqueName);
      const pDisplay       = (p.displayName || '').toLowerCase();
      const jFile          = flowName.toLowerCase();
      const jFileTrimmed   = stripTrailingGuid(flowName).toLowerCase();
      const jFileNorm      = normForMatch(flowName);
      const jDisplay       = resolvedDisplayName.toLowerCase();
      return (
        pUnique === jFile ||
        pUniqueTrimmed === jFileTrimmed ||
        pUniqueNorm === jFileNorm ||
        pDisplay === jDisplay ||
        pUnique === jDisplay ||
        pUniqueTrimmed === jDisplay ||
        pDisplay === jFile ||
        pDisplay === jFileTrimmed ||
        // Handle "Name-GUID" or "Name_GUID" suffixes common in PP exports
        jFile.startsWith(pUnique + '-') ||
        jFile.startsWith(pUnique + '_') ||
        jFile.startsWith(pUniqueTrimmed + '-') ||
        jFile.startsWith(pUniqueTrimmed + '_') ||
        pUnique.startsWith(jFile + '-') ||
        pUnique.startsWith(jFile + '_') ||
        pUniqueTrimmed.startsWith(jFileTrimmed + '-') ||
        pUniqueTrimmed.startsWith(jFileTrimmed + '_')
      );
    });

    if (existingIdx >= 0) {
      const existing = processes[existingIdx];
      if (steps.length > 0) existing.steps = steps;
      if (connectors.length > 0) existing.flowConnectors = connectors;
      if (triggerDescription) existing.flowTrigger = triggerDescription;
      if (usedConnectionRefs.length > 0) existing.flowConnectionReferences = usedConnectionRefs;
      if (usedEnvVars.length > 0) existing.flowEnvironmentVariables = usedEnvVars;
      // Always prefer JSON display names for modern flows because they preserve
      // user-facing spacing/casing better than workflow metadata exports.
      if (jsonDisplayName) {
        existing.displayName = normalizeDisplayNameForReadability(jsonDisplayName, existing.uniqueName || flowName);
      }
      if (existing.displayName) {
        existing.displayName = normalizeDisplayNameForReadability(existing.displayName, existing.uniqueName || flowName);
      }
      existing.flowDefinition = flowJson;
      // A flow JSON in the Workflows/ folder always means this is a Power Automate Flow.
      existing.category = ProcessCategory.PowerAutomateFlow;
      // Set primary entity from trigger if not already known
      if (triggerEntity && !existing.primaryEntity) {
        existing.primaryEntity = triggerEntity;
      }
      // Bubble up entity references from flow steps to relatedEntities
      const collectExistingStepEntities = (stps: ProcessStep[]): string[] =>
        stps.flatMap((s) => [
          ...(s.referencedEntities ?? []),
          ...collectExistingStepEntities(s.children ?? []),
        ]);
      const stepEntities = collectExistingStepEntities(steps);
      const allRelated = [...(existing.relatedEntities ?? []), ...stepEntities, existing.primaryEntity, triggerEntity]
        .filter((v): v is string => !!v);
      existing.relatedEntities = Array.from(new Set(allRelated));
    } else {
      const collectNewEntities = (stps: ProcessStep[]): string[] =>
        stps.flatMap((s) => [
          ...(s.referencedEntities ?? []),
          ...collectNewEntities(s.children ?? []),
        ]);
      const relatedEntities = Array.from(new Set([
        ...(triggerEntity ? [triggerEntity] : []),
        ...collectNewEntities(steps),
      ]));
      processes.push({
        name:           flowName,
        displayName:    resolvedDisplayName,
        uniqueName:     flowName,
        category:       ProcessCategory.PowerAutomateFlow,
        primaryEntity:  triggerEntity,
        isActivated:    parseProcessActivationStatus(flowJson),
        steps,
        flowDefinition: flowJson,
        flowTrigger:    triggerDescription,
        flowConnectors: connectors,
        flowConnectionReferences: usedConnectionRefs,
        flowEnvironmentVariables: usedEnvVars,
        relatedEntities,
      } as ProcessDefinition);
    }
  });

  // ── Step 6: Plugin assemblies from PluginAssemblies/ folder ─────────────
  report(80);
  const pluginFolderEntries = getEntriesWithPrefix(zip, 'PluginAssemblies/');
  for (const [path] of pluginFolderEntries) {
    if (path.endsWith('.dll')) {
      const assemblyName = path.split('/').pop()?.replace(/\.dll$/, '') ?? path;
      if (!pluginAssemblies.some((pa) => pa.assemblyName === assemblyName)) {
        pluginAssemblies.push({
          name:         assemblyName,
          displayName:  assemblyName,
          assemblyName,
          isolationMode: 2, // default Sandbox
          steps:        [],
        } as PluginAssemblyDefinition);
      }
    }
  }

  // Parse plugin step registrations from customizations.xml
  // (SdkMessageProcessingSteps)
  report(85);
  if (customizationsXml) {
    try {
      const doc  = xmlParser.parse(customizationsXml) as Record<string, unknown>;
      const root = (doc['ImportExportXml'] ?? doc) as Record<string, unknown>;
      const sdkRoot = (root['SdkMessageProcessingSteps'] ?? root['sdkmessageprocessingsteps'] ?? {}) as Record<string, unknown>;
      const rawSteps: unknown[] = (() => {
        const s = sdkRoot['SdkMessageProcessingStep'] ?? sdkRoot['sdkmessageprocessingstep'];
        if (!s) return [];
        return Array.isArray(s) ? s : [s];
      })();

      rawSteps.forEach((s) => {
        const sn          = s as Record<string, unknown>;
        const messageNode = (sn['SdkMessage'] ?? {}) as Record<string, unknown>;
        const filterNode  = (sn['SdkMessageFilter'] ?? {}) as Record<string, unknown>;
        const pluginType  = xmlStr(sn, 'PluginTypeName') || xmlStr(sn, 'plugintypename') || '';
        const message =
          xmlStr(messageNode, '@_Name') ||
          xmlStr(messageNode, 'Name') ||
          xmlStr(sn, 'MessageName') ||
          xmlStr(sn, 'messagename') ||
          findFirstStringByKey(sn, new Set(['messagename', 'message', 'name'])) ||
          '';
        const primaryEntity =
          xmlStr(filterNode, 'PrimaryObjectTypeCode') ||
          xmlStr(filterNode, '@_PrimaryObjectTypeCode') ||
          xmlStr(filterNode, 'PrimaryEntity') ||
          xmlStr(filterNode, '@_PrimaryEntity') ||
          xmlStr(sn, 'PrimaryObjectTypeCode') ||
          xmlStr(sn, 'primaryobjecttypecode') ||
          xmlStr(sn, 'PrimaryEntity') ||
          findFirstStringByKey(sn, new Set(['primaryobjecttypecode', 'primaryentity', 'entitylogicalname'])) ||
          undefined;

        const step: PluginStepDefinition = {
          name:                 xmlStr(sn, 'Name') || xmlStr(sn, '@_Name'),
          message,
          primaryEntity,
          stage:                Number(xmlStr(sn, 'Stage') || '20'),
          mode:                 Number(xmlStr(sn, 'Mode') || '0'),
          pluginTypeName:       pluginType,
          filteringAttributes:  xmlStr(sn, 'FilteringAttributes') || undefined,
          description:          xmlStr(sn, 'Description') || undefined,
        };

        // Attach to the relevant plugin assembly, or create a placeholder
        const assemblyName = pluginType.split('.')[0] ?? pluginType;
        const existingAsm  = pluginAssemblies.find((pa) => pa.assemblyName === assemblyName || pa.name === assemblyName);
        if (existingAsm) {
          existingAsm.steps.push(step);
        } else {
          pluginAssemblies.push({
            name:          assemblyName,
            displayName:   assemblyName,
            assemblyName,
            isolationMode: 2,
            steps:         [step],
          } as PluginAssemblyDefinition);
        }
      });

      // Also look for PluginAssemblies node in customizations.xml
      const paRoot = (root['PluginAssemblies'] ?? root['pluginassemblies'] ?? {}) as Record<string, unknown>;
      const rawPAs: unknown[] = (() => {
        const pa = paRoot['PluginAssembly'] ?? paRoot['pluginassembly'];
        if (!pa) return [];
        return Array.isArray(pa) ? pa : [pa];
      })();
      rawPAs.forEach((pa) => {
        const pan = pa as Record<string, unknown>;
        const assemblyName = xmlStr(pan, 'Name') || xmlStr(pan, '@_Name');
        if (!pluginAssemblies.some((x) => x.assemblyName === assemblyName)) {
          pluginAssemblies.push({
            name:          assemblyName,
            displayName:   assemblyName,
            assemblyName,
            version:       xmlStr(pan, 'Version') || undefined,
            isolationMode: Number(xmlStr(pan, 'IsolationMode') || '2'),
            sourceType:    xmlStr(pan, 'SourceType') || undefined,
            steps:         [],
          } as PluginAssemblyDefinition);
        }
      });
    } catch (err) {
      warnings.push(`Failed to parse plugin/step data: ${(err as Error).message}`);
    }
  }

  report(95);

  const knownEntities = new Set(entities.map((entity) => entity.logicalName.toLowerCase()));
  processes.forEach((process) => {
    const related = new Set<string>((process.relatedEntities ?? []).map((entity) => entity.toLowerCase()));
    if (process.primaryEntity) related.add(process.primaryEntity.toLowerCase());
    process.steps.forEach((step) => {
      addEntityMatchesFromText(step.name, knownEntities, related);
      if (step.description) addEntityMatchesFromText(step.description, knownEntities, related);
    });
    if (process.flowDefinition) {
      extractReferencedEntities(process.flowDefinition, knownEntities, related);
    }
    process.relatedEntities = Array.from(related);
    if (!process.primaryEntity && process.relatedEntities.length > 0) {
      process.primaryEntity = process.relatedEntities[0];
    }
    if (process.displayName) {
      process.displayName = normalizeDisplayNameForReadability(process.displayName, process.uniqueName || process.name);
    }
  });

  enrichComponentInventory(metadata, {
    optionSets,
    forms,
    views,
    processes,
    apps,
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
  });

  // ── Step 7: Assemble and return ──────────────────────────────────────────
  const parsed: ParsedSolution = {
    metadata,
    entities:               entities.filter((e) => !!e.logicalName),
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
    warnings,
  };

  report(100);
  return parsed;
}
