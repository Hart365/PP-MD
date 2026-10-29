import * as XLSX from 'xlsx-js-style';
import type { ParsedSolution, EntityDefinition } from '../types/solution';
import { applyDocumentationScope, DEFAULT_DOCUMENTATION_SETTINGS, filterDocumentationPrivileges, selectDocumentationAttributes, type DocumentContext, type DocumentationSettings } from '../generator/markdownGenerator';
import { buildTableMetadataBlocks, NOT_AVAILABLE_PREFIX, type MetadataBlock } from '../generator/tableMetadata';
import { collectMissingDependencies, dependenciesForTable, missingDependenciesNotice, MISSING_DEPENDENCIES_NOTE } from '../generator/missingDependencies';

export interface ExcelExportOptions {
  generateMissingDependencies?: boolean;
  documentationSettings?: DocumentationSettings;
}

type CellValue = string | number;

interface SheetLink {
  /** Zero-based column index within each data row that carries the link. */
  column: number;
  /** Target worksheet name for each data row (same order as rows). */
  targets: string[];
}

interface SheetDefinition {
  name: string;
  headers: string[];
  rows: CellValue[][];
  link?: SheetLink;
  /** Optional key/value block rendered above the main table. */
  preamble?: { title: string; rows: CellValue[][] };
  /** Optional heading row with a "back" link to another worksheet. */
  title?: { text: string; backLinkSheet?: string };
  /** Optional property sections (label/value pairs) rendered between the title and the main table. */
  metadataBlocks?: MetadataBlock[];
  dependencies?: { headers: string[]; rows: CellValue[][]; notice?: string };
}

/** Zero-based column index used for the "Back to ..." link on titled worksheets (column E). */
const BACK_LINK_COLUMN = 4;

export { NOT_AVAILABLE_PREFIX, buildTableMetadataBlocks };

const metadataBorder = {
  top: { style: 'thin', color: { rgb: '808080' } },
  bottom: { style: 'thin', color: { rgb: '808080' } },
  left: { style: 'thin', color: { rgb: '808080' } },
  right: { style: 'thin', color: { rgb: '808080' } },
};

const metadataHeadingStyle = {
  fill: { fgColor: { rgb: '252835' } },
  font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
  alignment: { vertical: 'center', wrapText: true },
  border: metadataBorder,
};

const metadataLabelStyle = {
  fill: { fgColor: { rgb: 'D6DCE4' } },
  font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: '000000' } },
  alignment: { vertical: 'top', wrapText: true },
  border: metadataBorder,
};

const metadataValueStyle = {
  fill: { fgColor: { rgb: 'FFFFFF' } },
  font: { name: 'Calibri', sz: 11, color: { rgb: '000000' } },
  alignment: { vertical: 'top', wrapText: true },
  border: metadataBorder,
};

const metadataUnavailableStyle = {
  fill: { fgColor: { rgb: 'FEF3C7' } },
  font: { name: 'Calibri', sz: 11, italic: true, color: { rgb: '92400E' } },
  alignment: { vertical: 'top', wrapText: true },
  border: metadataBorder,
};

const backButtonStyle = {
  fill: { fgColor: { rgb: 'DBEAFE' } },
  font: { bold: true, color: { rgb: '1D4ED8' }, underline: true },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: {
    top: { style: 'thin', color: { rgb: '93C5FD' } },
    bottom: { style: 'thin', color: { rgb: '93C5FD' } },
    left: { style: 'thin', color: { rgb: '93C5FD' } },
    right: { style: 'thin', color: { rgb: '93C5FD' } },
  },
};

const sheetTitleStyle = {
  ...metadataHeadingStyle,
  font: { ...metadataHeadingStyle.font, sz: 14 },
};

const tableNameStyle = {
  ...metadataValueStyle,
  font: { ...metadataValueStyle.font, bold: true, sz: 14 },
  alignment: { vertical: 'center', wrapText: true },
};

const wrapStyle = { alignment: { wrapText: true, vertical: 'top' } };

const DESCRIPTION_COLUMN_WIDTH = 60;

function isDescriptionHeader(header: string): boolean {
  return header.trim().toLowerCase() === 'description';
}

const linkStyle = {
  font: { color: { rgb: '1D4ED8' }, underline: true },
};

const sectionTitleStyle = {
  font: { bold: true, sz: 12, color: { rgb: '1E3A8A' } },
};

const MAX_SHEET_NAME_LENGTH = 31;

function compareFirstColumn(a: CellValue[], b: CellValue[]): number {
  const left = a[0];
  const right = b[0];
  if (typeof left === 'number' && typeof right === 'number') return left - right;
  return text(left).localeCompare(text(right), undefined, { sensitivity: 'base', numeric: true });
}

/** Sorts rows on the first column, keeping any per-row link targets aligned. */
function sortDefinition(definition: SheetDefinition): SheetDefinition {
  const indexed = definition.rows.map((row, index) => ({ row, target: definition.link?.targets[index] }));
  indexed.sort((left, right) => compareFirstColumn(left.row, right.row));
  return {
    ...definition,
    rows: indexed.map((entry) => entry.row),
    link: definition.link
      ? { column: definition.link.column, targets: indexed.map((entry) => entry.target ?? '') }
      : undefined,
  };
}

/** Produces a valid, unique Excel worksheet name (max 31 chars, no []:*?/\ ). */
function uniqueSheetName(raw: string, used: Set<string>): string {
  let base = raw.replace(/[[\]:*?/\\]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '');
  if (!base || base.toLowerCase() === 'history') base = base ? `${base} Table` : 'Table';
  base = base.slice(0, MAX_SHEET_NAME_LENGTH).trim();

  let candidate = base;
  let counter = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = ` (${counter})`;
    candidate = `${base.slice(0, MAX_SHEET_NAME_LENGTH - suffix.length).trim()}${suffix}`;
    counter += 1;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function entityDisplayName(entity: EntityDefinition): string {
  return text(entity.displayName).trim() || entity.logicalName;
}

function hasDocumentContext(context: DocumentContext | undefined): context is DocumentContext {
  if (!context) return false;
  return Object.values(context).some((value) => typeof value === 'string' && value.trim().length > 0);
}

function documentContextRows(context: DocumentContext): CellValue[][] {
  return [
    ['Client', context.client],
    ['Contract', context.contract],
    ['Contract ID/SoW', context.sow],
    ['Project', context.project],
    ['Sprint', context.sprint],
    ['Release Date', context.releaseDate],
  ]
    .map(([label, value]) => [label, text(value).trim()])
    .filter(([, value]) => value.length > 0);
}

const headerStyle = {
  fill: { fgColor: { rgb: '1D4ED8' } },
  font: { bold: true, color: { rgb: 'FFFFFF' } },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: { bottom: { style: 'medium', color: { rgb: '93C5FD' } } },
};

function text(value: unknown): string {
  if (value === undefined || value === null || value === '') return '';
  return String(value);
}

function internalLink(sheetName: string): string {
  return `#'${sheetName.replace(/'/g, "''")}'!A1`;
}

function titleRowValues(title: NonNullable<SheetDefinition['title']>): CellValue[] {
  if (!title.backLinkSheet) return [title.text];
  const row: CellValue[] = Array.from({ length: BACK_LINK_COLUMN + 1 }, () => '');
  row[0] = 'Table Name(s)';
  row[1] = title.text;
  row[BACK_LINK_COLUMN] = `← Back to ${title.backLinkSheet}`;
  return row;
}

function makeSheet(definition: SheetDefinition): XLSX.WorkSheet {
  const aoa: CellValue[][] = [];
  const titleRow = definition.title ? titleRowValues(definition.title) : [];
  if (definition.title) {
    aoa.push(titleRow);
    aoa.push([]);
  }
  const metadataCells: { row: number; column: number; span: number; kind: 'heading' | 'label' | 'value' }[] = [];
  const metadataStartRow = aoa.length;
  let metadataColumnCount = 0;
  (definition.metadataBlocks ?? []).forEach((block) => {
    const pairs = Math.max(1, Math.ceil(block.items.length / 6));
    const rowsPerPair = Math.ceil(block.items.length / pairs);
    const column = metadataColumnCount;
    metadataColumnCount += pairs * 2;
    aoa[metadataStartRow] ??= [];
    aoa[metadataStartRow][column] = block.heading;
    metadataCells.push({ row: metadataStartRow, column, span: pairs * 2, kind: 'heading' });
    block.items.forEach(([label, value], index) => {
      const row = metadataStartRow + 1 + index % rowsPerPair;
      const labelColumn = column + Math.floor(index / rowsPerPair) * 2;
      aoa[row] ??= [];
      aoa[row][labelColumn] = label;
      aoa[row][labelColumn + 1] = value;
      metadataCells.push(
        { row, column: labelColumn, span: 1, kind: 'label' },
        { row, column: labelColumn + 1, span: 1, kind: 'value' },
      );
    });
  });
  if (metadataCells.length > 0) aoa.push([]);
  const titleRows = aoa.length;
  const hasPreamble = Boolean(definition.preamble && definition.preamble.rows.length > 0);
  if (definition.preamble && hasPreamble) {
    aoa.push([definition.preamble.title]);
    aoa.push(...definition.preamble.rows);
    aoa.push([]);
  }
  const headerRow = aoa.length;
  aoa.push(definition.headers, ...definition.rows);
  let dependencyTitleRow = -1;
  if (definition.dependencies) {
    aoa.push([]);
    dependencyTitleRow = aoa.length;
    aoa.push(['Missing Dependencies'], [
      MISSING_DEPENDENCIES_NOTE,
    ], definition.dependencies.headers);
    aoa.push(...(definition.dependencies.rows.length > 0
      ? definition.dependencies.rows
      : [[definition.dependencies.notice ? 'Dependency information is unavailable or incomplete.'
        : 'No missing installation/import prerequisites were declared for this table.']]));
  }

  const sheet = XLSX.utils.aoa_to_sheet(aoa);

  if (definition.title) {
    const titleCell = sheet[XLSX.utils.encode_cell({ r: 0, c: 0 })];
    if (titleCell) titleCell.s = sheetTitleStyle;
    if (definition.title.backLinkSheet) {
      for (let column = 1; column < BACK_LINK_COLUMN; column += 1) {
        const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: column })];
        if (cell) cell.s = tableNameStyle;
      }
      sheet['!merges'] = [{ s: { r: 0, c: 1 }, e: { r: 0, c: BACK_LINK_COLUMN - 1 } }];
      const backCell = sheet[XLSX.utils.encode_cell({ r: 0, c: BACK_LINK_COLUMN })];
      if (backCell) {
        backCell.l = { Target: internalLink(definition.title.backLinkSheet), Tooltip: `Return to the ${definition.title.backLinkSheet} worksheet` };
        backCell.s = backButtonStyle;
      }
    }
    sheet['!rows'] = [{ hpt: 24 }];
  }

  if (metadataCells.length > 0) {
    const merges: XLSX.Range[] = sheet['!merges'] ?? [];
    metadataCells.forEach(({ row, column, span, kind }) => {
      if (kind === 'heading') {
        for (let c = column; c < column + span; c += 1) {
          const ref = XLSX.utils.encode_cell({ r: row, c });
          sheet[ref] = sheet[ref] ?? { t: 's', v: '' };
          const heading = text(aoa[row][column]);
          sheet[ref].s = heading === 'General' || heading === 'Data and tracking'
            ? metadataLabelStyle : metadataHeadingStyle;
        }
        merges.push({ s: { r: row, c: column }, e: { r: row, c: column + span - 1 } });
        return;
      }
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
      if (cell) {
        cell.s = kind === 'label' ? metadataLabelStyle
          : text(cell.v).startsWith(NOT_AVAILABLE_PREFIX) ? metadataUnavailableStyle : metadataValueStyle;
      }
    });
    sheet['!merges'] = merges;
  }

  if (hasPreamble) {
    const titleCell = sheet[XLSX.utils.encode_cell({ r: titleRows, c: 0 })];
    if (titleCell) titleCell.s = sectionTitleStyle;
    for (let r = titleRows + 1; r < headerRow - 1; r += 1) {
      const labelCell = sheet[XLSX.utils.encode_cell({ r, c: 0 })];
      if (labelCell) labelCell.s = { font: { bold: true } };
    }
  }

  definition.headers.forEach((header, column) => {
    if (!isDescriptionHeader(header)) return;
    definition.rows.forEach((_, index) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: headerRow + 1 + index, c: column })];
      if (cell) cell.s = { ...(cell.s ?? {}), ...wrapStyle };
    });
  });

  definition.headers.forEach((_, column) => {
    const headerCell = sheet[XLSX.utils.encode_cell({ r: headerRow, c: column })];
    if (headerCell) headerCell.s = headerStyle;
  });

  if (definition.link) {
    const { column, targets } = definition.link;
    targets.forEach((target, index) => {
      if (!target) return;
      const cell = sheet[XLSX.utils.encode_cell({ r: headerRow + 1 + index, c: column })];
      if (!cell) return;
      cell.l = { Target: internalLink(target), Tooltip: `Go to ${target}` };
      cell.s = linkStyle;
    });
  }

  const allRows = [titleRow, ...(definition.preamble?.rows ?? []), ...definition.rows];
  const columnCount = Math.max(definition.headers.length, metadataColumnCount);
  const columns = Array.from({ length: columnCount }, (_, index) => {
    const header = definition.headers[index] ?? '';
    if (isDescriptionHeader(header)) return { wch: DESCRIPTION_COLUMN_WIDTH };
    const longest = Math.max(header.length, ...allRows.map((row) => text(row[index]).length));
    const minimum = index < metadataColumnCount ? (index % 2 === 0 ? 28 : 24) : 12;
    return { wch: Math.min(Math.max(longest + 2, minimum), 48) };
  });
  sheet['!cols'] = columns;

  // Explicit heights keep wrapped metadata readable in viewers that do not auto-fit rows.
  const rowHeights: XLSX.RowInfo[] = sheet['!rows'] ?? [];
  metadataCells.forEach(({ row, column, span, kind }) => {
    const width = columns.slice(column, column + span).reduce((sum, item) => sum + item.wch, 0) - 2;
    const value = text(sheet[XLSX.utils.encode_cell({ r: row, c: column })]?.v);
    const lines = value.split(/\r\n|\r|\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / width)), 0);
    const height = Math.max(kind === 'heading' ? 24 : 20, lines * 15 + 4);
    rowHeights[row] = { hpt: Math.max(rowHeights[row]?.hpt ?? 0, height) };
  });
  if (metadataCells.length > 0) sheet['!rows'] = rowHeights;
  if (definition.dependencies) {
    const lastColumn = definition.headers.length - 1;
    const merges: XLSX.Range[] = sheet['!merges'] ?? [];
    for (const row of [dependencyTitleRow, dependencyTitleRow + 1]) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
      cell.s = row === dependencyTitleRow ? sectionTitleStyle : wrapStyle;
      merges.push({ s: { r: row, c: 0 }, e: { r: row, c: lastColumn } });
      const width = columns.slice(0, lastColumn + 1).reduce((sum, item) => sum + item.wch, 0) - 2;
      const lines = text(cell.v).split(/\r\n|\r|\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / width)), 0);
      rowHeights[row] = { hpt: Math.max(24, lines * 15 + 4) };
    }
    definition.dependencies.headers.forEach((_, column) => {
      sheet[XLSX.utils.encode_cell({ r: dependencyTitleRow + 2, c: column })].s = headerStyle;
    });
    const rows = definition.dependencies.rows.length > 0 ? definition.dependencies.rows : [aoa[dependencyTitleRow + 3]];
    rows.forEach((values, index) => {
      const row = dependencyTitleRow + 3 + index;
      let lines = 1;
      values.forEach((value, column) => {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
        cell.s = metadataValueStyle;
        lines = Math.max(lines, ...text(value).split(/\r\n|\r|\n/).map((line) => Math.ceil(line.length / (columns[column].wch - 2))));
      });
      rowHeights[row] = { hpt: lines * 15 + 4 };
    });
    if (definition.dependencies.rows.length === 0) {
      merges.push({ s: { r: dependencyTitleRow + 3, c: 0 }, e: { r: dependencyTitleRow + 3, c: lastColumn } });
      rowHeights[dependencyTitleRow + 3] = { hpt: 24 };
    }
    sheet['!merges'] = merges;
    sheet['!rows'] = rowHeights;
  }

  const tableRange = {
    s: { r: headerRow, c: 0 },
    e: { r: headerRow + definition.rows.length, c: Math.max(definition.headers.length - 1, 0) },
  };
  if (definition.rows.length > 0) sheet['!autofilter'] = { ref: XLSX.utils.encode_range(tableRange) };
  sheet['!freeze'] = { xSplit: 0, ySplit: headerRow + 1 };
  return sheet;
}

const COLUMN_HEADERS = ['Display Name', 'Schema Name', 'Type', 'Required', 'Custom', 'Primary Name', 'Lookup Target', 'Option Set', 'Description'];

function optionSetSummary(attribute: EntityDefinition['attributes'][number]): string {
  if (attribute.optionSetName) return attribute.optionSetName;
  if (!attribute.options || attribute.options.length === 0) return '';
  return `Local: ${attribute.options.map((option) => `${option.label} (${option.value})`).join(', ')}`;
}

/**
 * Flattens the appendix into Section/Item/Value/Custom (confirmed)/Unclassified/Details rows. Only the consolidated
 * "All Selected Solutions" result carries an appendix, so other workbooks omit this sheet.
 */
function buildSheets(solution: ParsedSolution, documentContext: DocumentContext | undefined, options: ExcelExportOptions): SheetDefinition[] {
  const settings = options.documentationSettings;
  const metadata = { ...DEFAULT_DOCUMENTATION_SETTINGS.metadata, ...settings?.metadata };
  const selectedColumns = settings ? [
    true, true, metadata.includeTypeColumn, metadata.includeRequiredLevelInfo, metadata.includeCustomColumn,
    metadata.includeNotesColumn, metadata.includeNotesColumn, metadata.includeNotesColumn, metadata.includeDescriptionColumn,
  ] : COLUMN_HEADERS.map(() => true);
  const extraHeaders = settings ? [
    ...(metadata.includeAuditInfo ? ['Audited'] : []),
    ...(metadata.includeFieldSecurityFlags ? ['Field Security'] : []),
    ...(metadata.includeValidForAdvancedFindInfo ? ['Advanced Find'] : []),
    ...(metadata.includeMetadataDiagnosticInfo ? ['Metadata Source'] : []),
  ] : [];
  const usedNames = new Set(
    ['Overview', 'Tables', 'Relationships', 'Processes', 'Apps', 'Integration', 'Security', 'Automation', 'Global Option Sets', 'Solution Components', 'Environment']
      .map((name) => name.toLowerCase()),
  );

  const tableSheets = [...solution.entities]
    .sort((a, b) => entityDisplayName(a).localeCompare(entityDisplayName(b), undefined, { sensitivity: 'base', numeric: true }))
    .map((entity) => ({ entity, sheetName: uniqueSheetName(entityDisplayName(entity), usedNames) }));
  const sheetNameByEntity = new Map(tableSheets.map(({ entity, sheetName }) => [entity, sheetName]));

  const sheets: SheetDefinition[] = [
    {
      name: 'Overview',
      headers: ['Property', 'Value'],
      preamble: hasDocumentContext(documentContext)
        ? { title: 'Document Information', rows: documentContextRows(documentContext) }
        : undefined,
      rows: [
        ['Solution', solution.metadata.displayName],
        ['Unique Name', solution.metadata.uniqueName],
        ['Version', solution.metadata.version],
        ['Publisher', solution.metadata.publisherName],
        ['Managed', solution.metadata.isManaged ? 'Yes' : 'No'],
        ['Tables', solution.entities.length],
        ['Global Option Sets', solution.optionSets.length],
        ['Additional Solution Components', solution.solutionComponents?.length ?? 0],
        ['Processes', solution.processes.length],
        ['Apps', solution.apps.length],
        ['Environment Variables', solution.environmentVariables.length],
        ['Connection References', solution.connectionReferences.length],
      ],
    },
    {
      name: 'Tables',
      headers: ['Display Name', 'Table', 'Description', 'Custom', 'Ownership', 'Primary Attribute', 'Columns', 'Relationships'],
      rows: solution.entities.map((entity) => [
        entityDisplayName(entity),
        entity.logicalName,
        text(entity.description),
        entity.isCustom ? 'Yes' : 'No',
        text(entity.ownershipType),
        text(entity.primaryAttributeName),
        entity.attributes.length,
        entity.relationships.length,
      ]),
      link: { column: 0, targets: solution.entities.map((entity) => sheetNameByEntity.get(entity) ?? '') },
    },
  ];

  const supportingSheets: SheetDefinition[] = [
    {
      name: 'Global Option Sets',
      headers: ['Schema Name', 'Display Name', 'Metadata ID', 'Description', 'Option Label', 'Option Value', 'Colour', 'Option Description'],
      rows: solution.optionSets.flatMap((choice) => choice.options.length > 0
        ? choice.options.map((option) => [choice.name, text(choice.displayName), text(choice.metadataId), text(choice.description), option.label, option.value, text(option.color), text(option.description)])
        : [[choice.name, text(choice.displayName), text(choice.metadataId), text(choice.description), '', '', '', '']]),
    },
    {
      name: 'Solution Components',
      headers: ['Type', 'Type Code', 'Name', 'Object ID', 'Table', 'Property', 'Value'],
      rows: [
        ...(solution.solutionComponents ?? []).flatMap((component) => (component.properties.length > 0 ? component.properties : [{ name: '', value: '' }])
          .map((property) => [component.componentTypeName, component.componentType, component.name, component.objectId, text(component.entityLogicalName), property.name, property.value])),
        ...(solution.unmappedComponents ?? []).map((component) => [text(component.componentTypeName) || 'Unknown', component.componentType, '', component.objectId, '', '', '']),
      ],
    },
    {
      name: 'Relationships',
      headers: ['Name', 'Type', 'Referenced Table', 'Referencing Table', 'Referencing Column', 'Referenced Column', 'Delete', 'Assign', 'Reparent'],
      rows: solution.entities.flatMap((entity) => entity.relationships.map((relationship) => [
        relationship.name,
        relationship.type,
        relationship.referencedEntity,
        relationship.referencingEntity,
        text(relationship.referencingAttribute),
        text(relationship.referencedAttribute),
        text(relationship.cascadeDelete),
        text(relationship.cascadeAssign),
        text(relationship.cascadeReparent),
      ])),
    },
    {
      name: 'Processes',
      headers: ['Name', 'Display Name', 'Category', 'Primary Table', 'Status', 'Trigger', 'Connectors', 'Environment Variables'],
      rows: solution.processes.map((process) => [
        process.name,
        text(process.displayName),
        process.category,
        text(process.primaryEntity),
        process.isActivated ? 'Activated' : 'Draft/Unknown',
        text(process.flowTrigger || process.triggerType),
        (process.flowConnectors ?? []).join(', '),
        (process.flowEnvironmentVariables ?? []).join(', '),
      ]),
    },
    {
      name: 'Apps',
      headers: ['Name', 'Display Name', 'Type', 'Tables', 'Sitemap Areas', 'Connectors', 'Enabled'],
      rows: solution.apps.map((app) => [
        app.uniqueName || app.name,
        text(app.displayName),
        app.appType,
        (app.entities ?? []).join(', '),
        (app.sitemapAreas ?? []).join(', '),
        (app.connectors ?? []).join(', '),
        app.isEnabled === false ? 'No' : 'Yes',
      ]),
    },
    {
      name: 'Integration',
      headers: ['Kind', 'Name', 'Display Name', 'Type', 'Connector', 'Has Current Value', 'Default Value'],
      rows: [
        ...solution.connectionReferences.map((reference) => ['Connection Reference', reference.name, text(reference.displayName), '', text(reference.connectorDisplayName || reference.connectorId), '', '']),
        ...solution.environmentVariables.map((variable) => ['Environment Variable', variable.schemaName, text(variable.displayName), text(variable.type), '', variable.hasCurrentValue === undefined ? 'Unknown' : variable.hasCurrentValue ? 'Yes' : 'No', text(variable.defaultValue)]),
      ],
    },
    {
      name: 'Security',
      headers: ['Role/Profile', 'Kind', 'Privilege or Attribute', 'Depth/Permission'],
      rows: [
        ...solution.securityRoles.flatMap((role) => role.privileges.map((privilege) => [role.name, 'Role', privilege.privilegeName, privilege.depth])),
        ...solution.fieldSecurityProfiles.flatMap((profile) => profile.permissions.map((permission) => [profile.name, 'Field Security Profile', permission.attributeName, `Read:${permission.canRead ? 'Y' : 'N'} Update:${permission.canUpdate ? 'Y' : 'N'} Create:${permission.canCreate ? 'Y' : 'N'}`])),
      ],
    },
    {
      name: 'Automation',
      headers: ['Kind', 'Name', 'Display Name', 'Description', 'Source', 'State/Mode'],
      rows: [
        ...solution.agents.map((item) => ['Agent', item.name, text(item.displayName), text(item.description), item.sourcePath, text(item.agentType)]),
        ...solution.aiModels.map((item) => ['AI Model', item.name, text(item.displayName), text(item.description), item.sourcePath, text(item.provider)]),
        ...solution.desktopFlows.map((item) => ['Desktop Flow', item.name, text(item.displayName), text(item.description), item.sourcePath, item.isEnabled ? 'Enabled' : 'Disabled']),
        ...solution.dataflows.map((item) => ['Dataflow', item.name, text(item.displayName), text(item.description), item.sourcePath, text(item.refreshMode)]),
        ...solution.customApis.map((item) => ['Custom API', item.name, text(item.displayName), text(item.description), item.sourcePath, item.isFunction ? 'Function' : 'Action']),
        ...solution.offlineProfiles.map((item) => ['Offline Profile', item.name, text(item.displayName), text(item.description), item.sourcePath, text(item.profileType)]),
      ],
    },
  ];

  const missingDependencies = options.generateMissingDependencies ? collectMissingDependencies(solution) : [];
  const perTableSheets: SheetDefinition[] = settings?.detailLevel === 'summary' ? [] : tableSheets.map(({ entity, sheetName }) => ({
    name: sheetName,
    title: { text: `${entityDisplayName(entity)} (${entity.logicalName})`, backLinkSheet: 'Tables' },
    metadataBlocks: buildTableMetadataBlocks(entity),
    dependencies: options.generateMissingDependencies ? {
      notice: missingDependenciesNotice(solution),
      headers: ['Dependent Item', 'Item Type', 'Required Component', 'Required Type', 'Solution'],
      rows: dependenciesForTable(missingDependencies, entity.logicalName).map((dependency) => [
        dependency.dependentName, dependency.dependentType,
        dependency.requiredParent ? `${dependency.requiredParent}.${dependency.requiredName}` : dependency.requiredName,
        dependency.requiredType, dependency.solutionUniqueName ?? solution.metadata.uniqueName,
      ]),
    } : undefined,
    headers: [...COLUMN_HEADERS.filter((_, index) => selectedColumns[index]), ...extraHeaders],
    rows: entity.attributes.map((attribute) => [
      text(attribute.displayName).trim() || attribute.name,
      attribute.name,
      attribute.type,
      attribute.required ? 'Yes' : 'No',
      attribute.isCustom ? 'Yes' : 'No',
      attribute.isPrimaryName ? 'Yes' : 'No',
      text(attribute.lookupTarget),
      optionSetSummary(attribute),
      text(attribute.description),
    ].filter((_, index) => selectedColumns[index]).concat(settings ? [
      ...(metadata.includeAuditInfo ? [attribute.isAuditEnabled === undefined ? 'Unknown' : attribute.isAuditEnabled ? 'Yes' : 'No'] : []),
      ...(metadata.includeFieldSecurityFlags ? [attribute.isSecured ? 'Yes' : 'No'] : []),
      ...(metadata.includeValidForAdvancedFindInfo ? [attribute.isValidForAdvancedFind === undefined ? 'Unknown' : attribute.isValidForAdvancedFind ? 'Yes' : 'No'] : []),
      ...(metadata.includeMetadataDiagnosticInfo ? [Object.values(attribute.metadataSources ?? {}).join(', ')] : []),
    ] : [])),
  }));

  // Per-table tabs go last so the summary sheets stay together at the front of the workbook.
  return [
    ...sheets.filter((sheet) => sheet.rows.length > 0 || sheet.name === 'Overview'),
    ...supportingSheets.filter((sheet) => sheet.rows.length > 0),
    ...perTableSheets,
  ].map(sortDefinition);
}

export function createSolutionWorkbook(solution: ParsedSolution, documentContext?: DocumentContext, options: ExcelExportOptions = {}): Blob {
  const settings = options.documentationSettings;
  if (settings) {
    const scoped = applyDocumentationScope(solution, settings);
    solution = {
      ...scoped,
      entities: scoped.entities.map((entity) => ({ ...entity, attributes: selectDocumentationAttributes(entity, scoped.forms, settings.metadata) })),
      securityRoles: scoped.securityRoles.map((role) => ({ ...role, privileges: filterDocumentationPrivileges(role, scoped.entities, settings.securityRoleFilters) })),
    };
  }
  const workbook = XLSX.utils.book_new();
  buildSheets(solution, documentContext, options).forEach((definition) => {
    const sheet = makeSheet(definition);
    XLSX.utils.book_append_sheet(workbook, sheet, definition.name);
  });
  const data = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', cellStyles: true });
  return new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
