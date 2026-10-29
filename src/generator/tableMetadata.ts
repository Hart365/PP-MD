import type { EntityDefinition } from '../types/solution';

export interface MetadataBlock {
  heading: string;
  items: [string, string][];
}

export const NOT_AVAILABLE_PREFIX = 'Not available';

const LOCAL_ONLY_REASON = 'not included in solution ZIP exports';
const UNKNOWN_SOURCE_REASON = 'not present in the analysed metadata';

function text(value: unknown): string {
  if (value === undefined || value === null || value === '') return '';
  return String(value);
}

export function isUnavailableValue(value: string): boolean {
  return value.startsWith(NOT_AVAILABLE_PREFIX);
}

function unavailable(reason: string): string {
  return `${NOT_AVAILABLE_PREFIX} – ${reason}`;
}

function missingReason(entity: EntityDefinition, localExportHasIt: boolean): string {
  if (entity.metadataSource === 'local') {
    return localExportHasIt
      ? 'not set in the solution customizations.xml for this table'
      : LOCAL_ONLY_REASON;
  }
  return UNKNOWN_SOURCE_REASON;
}

function yesNo(entity: EntityDefinition, value: boolean | undefined, localExportHasIt = true): string {
  if (value === undefined) return unavailable(missingReason(entity, localExportHasIt));
  return value ? 'Yes' : 'No';
}

function ownershipLabel(entity: EntityDefinition): string {
  switch (entity.ownershipType) {
    case 'User': return 'User or team';
    case 'Organization': return 'Organisation';
    case 'None': return 'None';
    default: return unavailable(missingReason(entity, true));
  }
}

function primaryColumnLabel(entity: EntityDefinition): string {
  const primaryName = entity.primaryAttributeName?.toLowerCase();
  const attribute = entity.attributes.find((item) =>
    (primaryName && item.name.toLowerCase() === primaryName)
    || item.isPrimaryName);
  if (attribute) {
    const display = text(attribute.displayName).trim();
    const schema = attribute.name;
    return display && display !== schema ? `${display} (${schema})` : schema;
  }
  if (entity.primaryAttributeName) return entity.primaryAttributeName;
  return unavailable(missingReason(entity, true));
}

function tableTypeLabel(entity: EntityDefinition): string {
  if (entity.tableType) {
    return entity.metadataSource === 'local' && entity.tableType === 'Standard'
      ? 'Standard (Elastic/Virtual types cannot be identified from a solution ZIP)'
      : entity.tableType;
  }
  if (entity.isActivity !== undefined) return entity.isActivity ? 'Activity' : 'Standard';
  return unavailable(missingReason(entity, true));
}

function customHelpLabel(entity: EntityDefinition): string {
  if (entity.entityHelpUrlEnabled === undefined) return yesNo(entity, undefined);
  if (!entity.entityHelpUrlEnabled) return 'No';
  return entity.entityHelpUrl ? `Yes – ${entity.entityHelpUrl}` : 'Yes';
}

/**
 * Builds the maker-portal style "table properties" sections for a table.
 * Values that could not be determined start with {@link NOT_AVAILABLE_PREFIX} and include the reason.
 */
export function buildTableMetadataBlocks(entity: EntityDefinition): MetadataBlock[] {
  const description = text(entity.description).trim();
  const displayName = text(entity.displayName).trim() || entity.logicalName;
  return [
    {
      heading: 'General',
      items: [
        ['Display name', displayName],
        ['Schema name', entity.logicalName],
        ['Table type', tableTypeLabel(entity)],
        ['Primary column', primaryColumnLabel(entity)],
        ['Description', description || '(No description provided)'],
        ['Total columns', String(entity.attributes.length)],
        ['Custom table', entity.isCustom ? 'Yes' : 'No'],
        ['Enable attachments (including notes and files)', yesNo(entity, entity.hasNotes, false)],
        ['Record ownership', ownershipLabel(entity)],
      ],
    },
    {
      heading: 'Data and tracking',
      items: [
        ['Apply duplicate detection rules', yesNo(entity, entity.isDuplicateDetectionEnabled)],
        ['Audit changes to its data', yesNo(entity, entity.isAuditEnabled)],
        ['Track changes', yesNo(entity, entity.changeTracking)],
        ['Leverage quick create form if available', yesNo(entity, entity.isQuickCreateEnabled)],
        ['Provide custom help', customHelpLabel(entity)],
        ['Enable long term retention', yesNo(entity, entity.isRetentionEnabled, false)],
      ],
    },
    {
      heading: 'Make this table an option when',
      items: [
        ['Creating a new activity', yesNo(entity, entity.hasActivities, false)],
        ['Sending email (if an email field does not exist, one will be created)', yesNo(entity, entity.isActivityParty)],
        ['Working with OneNote', yesNo(entity, entity.isOneNoteIntegrationEnabled)],
        ['Doing a mail merge', yesNo(entity, entity.isMailMergeEnabled)],
        ['Setting up SharePoint document management', yesNo(entity, entity.isDocumentManagementEnabled)],
      ],
    },
    {
      heading: 'Rows in this table',
      items: [
        ['Can have connections', yesNo(entity, entity.isConnectionsEnabled)],
        ['Can have a contact email', yesNo(entity, entity.hasEmailAddresses, false)],
        ['Have an access team', yesNo(entity, entity.autoCreateAccessTeams)],
        ['Can be linked to feedback', yesNo(entity, entity.hasFeedback, false)],
        ['Appear in search results', yesNo(entity, entity.syncToExternalSearchIndex)],
        ['Can be taken offline', yesNo(entity, entity.isAvailableOffline)],
        ['Can be added to a queue', yesNo(entity, entity.isValidForQueue, false)],
        ["When rows are created or assigned, move them to the owner's default queue", yesNo(entity, entity.autoRouteToOwnerQueue)],
      ],
    },
  ];
}

/** Returns the metadata sections with unavailable values removed, dropping any section left empty. */
export function availableTableMetadataBlocks(entity: EntityDefinition): MetadataBlock[] {
  return buildTableMetadataBlocks(entity)
    .map((block) => ({ ...block, items: block.items.filter(([, value]) => !isUnavailableValue(value)) }))
    .filter((block) => block.items.length > 0);
}
