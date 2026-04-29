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
  7: 'Role',
  8: 'RolePrivilege',
  9: 'FieldSecurityProfile',
  10: 'FieldPermission',
  14: 'Workflow',
  24: 'Form',              // SystemForm
  26: 'SavedQuery',        // View
  29: 'Workflow',          // WorkflowAccessRight
  31: 'Report',
  32: 'ReportEntity',
  33: 'ReportCategory',
  35: 'Report',
  44: 'Dashboard',
  60: 'AppModule',
  61: 'SiteMap',
  62: 'Process',
  70: 'FieldSecurityProfile',
  71: 'ConnectionReference',
  80: 'ConnectorDefinition',
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
  const solNode = (root['SolutionManifest'] ?? root) as Record<string, unknown>;

  const uniqueName = xmlStr(solNode, 'UniqueName') || xmlStr(solNode as Record<string, unknown>, '@_UniqueName');
  const version    = xmlStr(solNode, 'Version')    || xmlStr(solNode, '@_Version');
  const isManaged  = xmlStr(solNode, 'Managed') === 'true' || xmlStr(solNode, 'Managed') === '1';

  // Publisher node
  const publisher = (solNode['Publisher'] ?? {}) as Record<string, unknown>;
  const publisherNames = (publisher['LocalizedNames'] ?? {}) as Record<string, unknown>;
  const pNameNode  = publisherNames['LocalizedName'] as Record<string, unknown> | undefined;
  const publisherName = pNameNode
    ? xmlStr(pNameNode, '@_description') || xmlStr(pNameNode, 'description')
    : xmlStr(publisher, 'UniqueName');

  // Solution display name
  const localizedNames = (solNode['LocalizedNames'] ?? {}) as Record<string, unknown>;
  const lnNode = localizedNames['LocalizedName'] as Record<string, unknown> | undefined;
  const displayName = lnNode
    ? xmlStr(lnNode, '@_description') || xmlStr(lnNode, 'description')
    : uniqueName;

  // Description
  const descriptions = (solNode['Descriptions'] ?? {}) as Record<string, unknown>;
  const descNode = descriptions['Description'] as Record<string, unknown> | undefined;
  const description = descNode
    ? xmlStr(descNode, '@_description') || xmlStr(descNode, 'description')
    : '';

  return {
    uniqueName,
    displayName,
    version,
    publisherName,
    publisherUniqueName: xmlStr(publisher, 'UniqueName'),
    description,
    isManaged,
  };
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
  const t = raw
    .replace(/attribute$/, '')
    .replace(/[_\s-]+/g, '')
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

  const isCustom         = xmlStr(entity, 'IsCustomEntity') === 'true' || xmlStr(entity, '@_IsCustomEntity') === 'true';
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

  const attributes: EntityAttribute[] = rawAttrs.map((rAttr) => {
    const a = rAttr as Record<string, unknown>;
    const attrName    = xmlStr(a, 'Name').toLowerCase() || xmlStr(a, '@_Name').toLowerCase() || xmlStr(a, 'LogicalName');
    const typeStr     = xmlStr(a, 'Type') || xmlStr(a, '@_Type') || xmlStr(a, 'AttributeType');
    const attrType    = mapAttributeType(typeStr);
    // RequiredLevel may be a plain string OR a child element with @_Value attribute (schema varies by export version)
    const requiredLevelValue = (() => {
      const rl = a['RequiredLevel'];
      if (!rl) return '';
      if (typeof rl === 'string') return rl;
      const rlObj = firstObject(rl);
      return (rlObj ? xmlStr(rlObj, '@_Value') || xmlStr(rlObj, 'Value') || xmlStr(rlObj, '#text') : '') || String(rl);
    })();
    const required    = requiredLevelValue === 'Required' || requiredLevelValue === 'SystemRequired' || requiredLevelValue === 'ApplicationRequired';
    const isCustomAttr = xmlStr(a, 'IsCustomAttribute') === 'true' || xmlStr(a, '@_IsCustomAttribute') === 'true';
    const isPrimaryName = xmlStr(a, 'IsPrimaryName') === 'true' || xmlStr(a, '@_IsPrimaryName') === 'true';
    const isAuditEnabled = xmlStr(a, 'IsAuditEnabled') === 'true' || xmlStr(a, '@_IsAuditEnabled') === 'true'
      || (() => { const ia = firstObject(a['IsAuditEnabled']); return ia ? xmlStr(ia, '@_Value') !== 'false' && xmlStr(ia, 'Value') !== 'false' : false; })();
    const maxLength   = a['MaxLength'] !== undefined ? Number(a['MaxLength']) : undefined;
    const precision   = a['Precision'] !== undefined ? Number(a['Precision']) : undefined;

    const lookupTarget: string | undefined = (() => {
      if (attrType !== AttributeType.Lookup && attrType !== AttributeType.Owner && attrType !== AttributeType.Customer) {
        return undefined;
      }
      const targets = (a['Targets'] ?? a['targets'] ?? '') as string;
      return targets || undefined;
    })();

    const osName: string | undefined = (() => {
      const osRef = a['OptionSet'] as Record<string, unknown> | undefined;
      if (!osRef) return undefined;
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
      isCustom:        isCustomAttr,
      isPrimaryName,
      isAuditEnabled,
      maxLength,
      precision,
      lookupTarget,
      options,
      optionSetName:   osName,
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

function findFlowMatches(flow: Record<string, unknown>, candidates: string[]): string[] {
  if (candidates.length === 0) return [];
  const serialized = JSON.stringify(flow).toLowerCase();
  return Array.from(new Set(candidates.filter((candidate) => candidate && serialized.includes(candidate.toLowerCase()))));
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
): { steps: ProcessStep[]; connectors: string[]; triggerDescription: string; displayName: string | undefined } {
  const steps: ProcessStep[] = [];
  const connectors = new Set<string>();
  let triggerDescription = '';
  let displayName: string | undefined;

  try {
    // Real Power Platform solution exports wrap everything under 'properties'
    const props = (flowJson['properties'] ?? {}) as Record<string, unknown>;

    // Prefer display name from properties
    const propDisplayName = xmlStr(props, 'displayName') || xmlStr(props, 'DisplayName');
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
    Object.entries(triggers).forEach(([triggerName, triggerDef]) => {
      const td = triggerDef as Record<string, unknown>;
      const triggerType = xmlStr(td, 'type');
      const apiId = ((td['inputs'] as Record<string, unknown>)?.['host'] as Record<string, unknown>)?.['apiId'] as string | undefined;
      if (apiId) connectors.add(connectorDisplayFromId(apiId));
      triggerDescription = `${humanizeIdentifier(triggerName)} (${triggerType || '–'})`;
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
        entities.push(parseEntityNode(e as Record<string, unknown>, warnings));
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
          const extractControls = (node: Record<string, unknown>) => {
            const ctrl = node['control'];
            if (ctrl) {
              const ctrls = Array.isArray(ctrl) ? ctrl : [ctrl];
              ctrls.forEach((c: unknown) => {
                const cn = c as Record<string, unknown>;
                const id = xmlStr(cn, '@_id') || xmlStr(cn, 'id');
                if (id && !fields.some((x) => x.attributeName === id)) {
                  fields.push({ attributeName: id });
                }
              });
            }
            // Recurse into child elements
            Object.values(node).forEach((v) => {
              if (typeof v === 'object' && v !== null && !Array.isArray(v)) {
                extractControls(v as Record<string, unknown>);
              } else if (Array.isArray(v)) {
                v.forEach((item) => {
                  if (typeof item === 'object' && item !== null) {
                    extractControls(item as Record<string, unknown>);
                  }
                });
              }
            });
          };
          extractControls(formNode);
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
        displayName:   displayN,
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

    const envValuesBySchema = new Map<string, string>();
    const evValuesRoot = (root['environmentvariablevalues'] ?? root['EnvironmentVariableValues'] ?? {}) as Record<string, unknown>;
    const rawEnvValues: unknown[] = (() => {
      const ev = evValuesRoot['environmentvariablevalue'] ?? evValuesRoot['EnvironmentVariableValue'];
      if (!ev) return [];
      return Array.isArray(ev) ? ev : [ev];
    })();
    rawEnvValues.forEach((ev) => {
      const evn = ev as Record<string, unknown>;
      const schemaName =
        xmlStr(evn, 'schemaname') ||
        xmlStr(evn, '@_schemaname') ||
        xmlStr(evn, 'EnvironmentVariableDefinitionSchemaName') ||
        xmlStr(evn, 'environmentvariabledefinitionschemaname') ||
        xmlStr(evn, 'Name');
      const value = xmlStr(evn, 'value') || xmlStr(evn, 'Value') || xmlStr(evn, '@_value');
      if (schemaName && value) envValuesBySchema.set(schemaName, value);
    });

    rawEVDs.forEach((ev) => {
      const evn = ev as Record<string, unknown>;
      const schemaName = xmlStr(evn, 'schemaname') || xmlStr(evn, '@_schemaname');
      const currentValue = envValuesBySchema.get(schemaName);
      environmentVariables.push({
        name:          schemaName || xmlStr(evn, 'Name'),
        schemaName,
        displayName:   getLocalizedLabel(evn, xmlStr(evn, 'displayname') || xmlStr(evn, 'DisplayName') || schemaName),
        description:   xmlStr(evn, 'description') || undefined,
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

      const sitemapAreas = Array.from(new Set([
        ...extractSitemapAreas(xmlStr(amn, 'SiteMapXml') || xmlStr(amn, 'sitemapxml') || ''),
        ...extractSitemapAreas(xmlStr(amn, 'AppModuleXml') || xmlStr(amn, 'appmodulexml') || ''),
        xmlStr(amn, 'SiteMapUniqueName') || '',
      ].filter((value) => !!value)));

      apps.push({
        name:         uniqueName,
        displayName,
        appType:      AppType.ModelDriven,
        uniqueName,
        isEnabled:    xmlStr(amn, 'IsEnabled') !== 'false',
        version:      xmlStr(amn, 'ClientVersion') || xmlStr(amn, 'Version') || undefined,
        entities:     Array.from(appEntities),
        sitemapAreas,
        connectors:   Array.from(appConnectors),
      } as AppDefinition);
    });
  }

  // ── Step 4: Canvas apps from CanvasApps/ folder ─────────────────────────
  report(70);
  const canvasEntries = getEntriesWithPrefix(zip, 'CanvasApps/');
  const discoveredCanvasApps = new Set<string>();

  // Actual canvas/custom-page packages are typically exported as *.msapp
  for (const [path] of canvasEntries) {
    if (path.endsWith('.msapp')) {
      const appName = path.split('/').pop()?.replace(/\.(msapp|json)$/, '') ?? path;
      const cleanName = stripTrailingGuid(appName);
      const displayName = humanizeIdentifier(cleanName || appName);
      // Avoid duplicates with model-driven apps already found
      if (!apps.some((a) => a.uniqueName === appName)) {
        apps.push({
          name:        appName,
          displayName,
          appType:     AppType.Canvas,
          uniqueName:  appName,
          isEnabled:   true,
          entities:    [],
          connectors:  [],
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
        xmlStr(metadata, 'displayName') ||
        xmlStr(metadata, 'name') ||
        xmlStr((metadata['properties'] ?? {}) as Record<string, unknown>, 'displayName') ||
        xmlStr((metadata['properties'] ?? {}) as Record<string, unknown>, 'name') ||
        path.split('/').slice(-2)[0] ||
        path.split('/').pop()?.replace(/\.json$/, '') ||
        path;
      const uniqueName = stripTrailingGuid(rawName) || rawName;
      if (discoveredCanvasApps.has(uniqueName.toLowerCase()) || apps.some((a) => a.uniqueName.toLowerCase() === uniqueName.toLowerCase())) {
        continue;
      }

      const appEntities = new Set<string>();
      extractReferencedEntities(metadata, knownCanvasEntities, appEntities);

      const appConnectors = new Set<string>();
      extractConnectorNames(metadata, appConnectors);

      const version =
        xmlStr(metadata, 'version') ||
        xmlStr((metadata['properties'] ?? {}) as Record<string, unknown>, 'version') ||
        xmlStr((metadata['publishInfo'] ?? {}) as Record<string, unknown>, 'version') ||
        undefined;

      apps.push({
        name: uniqueName,
        displayName: humanizeIdentifier(uniqueName),
        appType,
        uniqueName,
        isEnabled: true,
        version,
        entities: Array.from(appEntities),
        connectors: Array.from(appConnectors),
      } as AppDefinition);
      discoveredCanvasApps.add(uniqueName.toLowerCase());
    } catch {
      // best-effort only
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

  // ── Step 5: Power Automate flows from Workflows/ folder ─────────────────
  report(75);
  const workflowEntries = getEntriesWithPrefix(zip, 'Workflows/');
  const flowJsonMap = new Map<string, Record<string, unknown>>();

  for (const [path, entry] of workflowEntries) {
    if (path.endsWith('.json')) {
      try {
        const jsonStr  = await entry.async('string');
        const flowJson = JSON.parse(jsonStr) as Record<string, unknown>;
        const flowName = path.split('/').pop()?.replace(/\.json$/, '') ?? path;
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

  flowJsonMap.forEach((flowJson, flowName) => {
    const { steps, connectors, triggerDescription, displayName: jsonDisplayName } = parseFlowDefinition(flowJson);
    const usedConnectionRefs = findFlowMatches(flowJson, connectionRefCandidates);
    const usedEnvVars = findFlowMatches(flowJson, envVarCandidates);

    // Prefer the display name embedded in the JSON properties (most human-readable)
    const flowNameWithoutGuid = stripTrailingGuid(flowName);
    const resolvedDisplayName = jsonDisplayName || humanizeIdentifier(flowNameWithoutGuid || flowName);

    const existingIdx = processes.findIndex((p) => {
      const pUnique  = p.uniqueName.toLowerCase();
      const pUniqueTrimmed = stripTrailingGuid(p.uniqueName).toLowerCase();
      const pDisplay = (p.displayName || '').toLowerCase();
      const jFile    = flowName.toLowerCase();
      const jFileTrimmed = stripTrailingGuid(flowName).toLowerCase();
      const jDisplay = resolvedDisplayName.toLowerCase();
      return (
        pUnique === jFile ||
        pUniqueTrimmed === jFileTrimmed ||
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
      // Use the JSON display name if the existing one is just the internal unique name
      if (jsonDisplayName && (!existing.displayName || existing.displayName === existing.uniqueName)) {
        existing.displayName = jsonDisplayName;
      }
      if (existing.displayName) existing.displayName = stripTrailingGuid(existing.displayName);
      existing.flowDefinition = flowJson;
      existing.category       = ProcessCategory.PowerAutomateFlow;
      // Bubble up entity references from flow steps to relatedEntities
      const collectExistingStepEntities = (stps: ProcessStep[]): string[] =>
        stps.flatMap((s) => [
          ...(s.referencedEntities ?? []),
          ...collectExistingStepEntities(s.children ?? []),
        ]);
      const stepEntities = collectExistingStepEntities(steps);
      if (stepEntities.length > 0) {
        const merged = [...(existing.relatedEntities ?? []), ...stepEntities, existing.primaryEntity].filter(Boolean) as string[];
        existing.relatedEntities = Array.from(new Set(merged));
      }
    } else {
      processes.push({
        name:           flowName,
        displayName:    resolvedDisplayName,
        uniqueName:     flowName,
        category:       ProcessCategory.PowerAutomateFlow,
        isActivated:    parseProcessActivationStatus(flowJson),
        steps,
        flowDefinition: flowJson,
        flowTrigger:    triggerDescription,
        flowConnectors: connectors,
        flowConnectionReferences: usedConnectionRefs,
        flowEnvironmentVariables: usedEnvVars,
        relatedEntities: (() => {
          const collectNewEntities = (stps: ProcessStep[]): string[] =>
            stps.flatMap((s) => [
              ...(s.referencedEntities ?? []),
              ...collectNewEntities(s.children ?? []),
            ]);
          return Array.from(new Set(collectNewEntities(steps)));
        })(),
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

        const step: PluginStepDefinition = {
          name:                 xmlStr(sn, 'Name') || xmlStr(sn, '@_Name'),
          message:              xmlStr(messageNode, '@_Name') || xmlStr(messageNode, 'Name') || '',
          primaryEntity:        xmlStr(filterNode, 'PrimaryObjectTypeCode') || undefined,
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
      process.displayName = stripTrailingGuid(process.displayName);
    }
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
