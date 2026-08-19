/**
 * @file tableExport.ts
 * @description Converts Markdown tables into an Excel workbook with one sheet
 * per table and a front Contents sheet with links to each table tab.
 */

import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx-js-style';

export interface MarkdownTable {
  title: string;
  heading2: string;
  heading3: string;
  heading4: string;
  headingPath: string;
  headers: string[];
  rows: string[][];
}

interface SheetBlock {
  title: string;
  headers: string[];
  rows: string[][];
}

interface SheetDefinition {
  sheetName: string;
  title: string;
  sectionTitle: string;
  blocks: SheetBlock[];
  rowCount: number;
}

interface DocumentHeaderEntry {
  field: string;
  value: string;
}

const HEADER_FILL = { patternType: 'solid', fgColor: { rgb: '1F4E78' } };
const ALT_FILL = { patternType: 'solid', fgColor: { rgb: 'F5F9FF' } };
const BORDER = {
  top: { style: 'thin', color: { rgb: '000000' } },
  bottom: { style: 'thin', color: { rgb: '000000' } },
  left: { style: 'thin', color: { rgb: '000000' } },
  right: { style: 'thin', color: { rgb: '000000' } },
};
const MULTI_ITEM_HEADER_PATTERN =
  /(tables?|columns?|connectors?|references?|dependencies|members|owners|roles?|permissions?|steps?|components?|apps?|flows?)/i;
const PERMISSION_HEADER_PATTERN =
  /(read|write|create|delete|append|assign|share|privilege|permission|access|scope|allowed|internal|status)/i;

function stripMarkdownDecorators(value: string): string {
  return value
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\*([^*]+)\*/g, '$1')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '$1');
}

const HTML_ENTITY_DECODE_MAP: Readonly<Record<string, string>> = {
  '&nbsp;': ' ',
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  "&#39;": "'",
};
const HTML_ENTITY_PATTERN = /&(?:nbsp|amp|lt|gt|quot|#39);/gi;

/** Decodes known HTML entities in a single pass to avoid double-unescaping. */
function decodeHtmlEntities(value: string): string {
  return value.replace(HTML_ENTITY_PATTERN, (match) => HTML_ENTITY_DECODE_MAP[match.toLowerCase()] ?? match);
}

function stripHtml(value: string): string {
  return decodeHtmlEntities(
    value
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, ''),
  );
}

function normalizeTableCell(cell: string): string {
  return stripHtml(stripMarkdownDecorators(cell.replace(/\\\|/g, '|').trim()));
}

function normalizeHeaderName(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeHeaderForFilter(value: string): string {
  return value
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function enforceVisibleHeaders(headers: string[], rows: string[][]): { headers: string[]; rows: string[][] } {
  const visibleIndexes = headers
    .map((header, idx) => ({ idx, header: normalizeHeaderForFilter(header) }))
    .filter((entry) => entry.header.length > 0)
    .map((entry) => entry.idx);

  if (visibleIndexes.length === headers.length) {
    return { headers, rows };
  }

  return {
    headers: visibleIndexes.map((idx) => headers[idx]),
    rows: rows.map((row) => visibleIndexes.map((idx) => row[idx] ?? '')),
  };
}

function splitMultiItemValue(value: string, header: string): string {
  const normalized = value.replace(/\r\n/g, '\n').trim();
  if (!normalized || normalized.includes('\n')) return normalized;

  const hasSemicolonList = normalized.includes(';') && normalized.split(';').length > 1;
  const hasCommaList = MULTI_ITEM_HEADER_PATTERN.test(header) && normalized.includes(',') && normalized.split(',').length > 1;

  if (hasSemicolonList) {
    return normalized
      .split(/\s*;\s*/)
      .map((item) => item.trim())
      .filter(Boolean)
      .join('\n');
  }

  if (hasCommaList) {
    return normalized
      .split(/\s*,\s*/)
      .map((item) => item.trim())
      .filter(Boolean)
      .join('\n');
  }

  return normalized;
}

function applyPermissionIcon(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) return trimmed;
  if (/^[⚫🔵🟢🟡🟠🔴✅❌]/u.test(trimmed)) return trimmed;

  const lower = trimmed.toLowerCase();
  if (lower === 'none') return '⚫ None';
  if (lower === 'user') return '🔵 User';
  if (lower === 'business unit') return '🟢 Business Unit';
  if (lower === 'parent-child bu') return '🟡 Parent-Child BU';
  if (lower === 'org' || lower === 'organization') return '🟠 Organization';
  if (lower === 'unknown') return '🔴 Unknown';
  if (lower === 'allowed') return '✅ Allowed';
  if (lower === 'not allowed') return '❌ Not Allowed';
  return trimmed;
}

function formatCellValueForExcel(header: string, rawValue: string): string {
  const withNewlines = splitMultiItemValue(rawValue, header);
  if (!withNewlines) return withNewlines;

  const shouldDecorate = PERMISSION_HEADER_PATTERN.test(header)
    || /^(none|user|business unit|parent-child bu|org|organization|unknown|allowed|not allowed)$/i.test(withNewlines.trim());
  if (!shouldDecorate) return withNewlines;

  return withNewlines
    .split('\n')
    .map((line) => applyPermissionIcon(line))
    .join('\n');
}

function removeEmptyColumns(headers: string[], rows: string[][]): { headers: string[]; rows: string[][] } {
  const keepIndexes = headers
    .map((header, index) => ({
      index,
      header: normalizeHeaderForFilter(header),
      hasData: rows.some((row) => (row[index] ?? '').trim().length > 0),
    }))
    .filter((entry) => entry.header.length > 0 && entry.hasData)
    .map((entry) => entry.index);

  if (keepIndexes.length === 0) {
    const fallback = headers
      .map((header, index) => ({ index, header: normalizeHeaderForFilter(header) }))
      .filter((entry) => entry.header.length > 0)
      .map((entry) => entry.index);

    if (fallback.length === 0) {
      return { headers: [], rows: [] };
    }

    return {
      headers: fallback.map((idx) => headers[idx]),
      rows: rows.map((row) => fallback.map((idx) => row[idx] ?? '')),
    };
  }

  return {
    headers: keepIndexes.map((idx) => headers[idx]),
    rows: rows.map((row) => keepIndexes.map((idx) => row[idx] ?? '')),
  };
}

function extractVersionFromSolutionName(value: string): string {
  const matches = [...value.matchAll(/\(([^()]+)\)/g)].map((match) => match[1].trim());
  for (let i = matches.length - 1; i >= 0; i -= 1) {
    const candidate = matches[i];
    if (/v?\d+(?:\.\d+){1,5}(?:[-+][A-Za-z0-9._-]+)?/i.test(candidate)) {
      return candidate;
    }
  }
  return '';
}

function stripVersionFromSolutionName(value: string): string {
  return value.replace(/\s*\(([^()]*\d[^()]*)\)\s*$/, '').trim();
}

function transformSheetBlockForExport(sheetTitle: string, block: SheetBlock): SheetBlock {
  const cleaned = removeEmptyColumns(block.headers, block.rows);
  const visible = enforceVisibleHeaders(cleaned.headers, cleaned.rows);
  if (visible.headers.length === 0) {
    return {
      title: block.title,
      headers: [],
      rows: [],
    };
  }

  const normalizedSheet = normalizeHeaderName(sheetTitle);
  const isDependencySheet = normalizedSheet.includes('solution dependencies');

  if (!isDependencySheet) {
    return {
      title: block.title,
      headers: visible.headers,
      rows: visible.rows.map((row) => row.map((value, idx) => formatCellValueForExcel(visible.headers[idx], value))),
    };
  }

  const normalizedHeaders = visible.headers.map((header) => normalizeHeaderName(header));
  const solutionNameIdx = normalizedHeaders.findIndex((header) => header === 'solution name');
  const versionIdx = normalizedHeaders.findIndex((header) => header === 'version');
  const indexesToKeep = visible.headers
    .map((header, idx) => ({ header: normalizeHeaderName(header), idx }))
    .filter((entry) => entry.header !== 'internal')
    .map((entry) => entry.idx);

  const nextHeaders = indexesToKeep.map((idx) => visible.headers[idx]);
  const nextRows = visible.rows.map((row) => {
    const projected = indexesToKeep.map((idx) => row[idx] ?? '');
    const projectedHeaders = nextHeaders.map((header) => normalizeHeaderName(header));
    const projectedSolutionIdx = projectedHeaders.findIndex((header) => header === 'solution name');
    const projectedVersionIdx = projectedHeaders.findIndex((header) => header === 'version');

    if (
      solutionNameIdx >= 0
      && versionIdx >= 0
      && projectedSolutionIdx >= 0
      && projectedVersionIdx >= 0
    ) {
      const sourceSolution = row[solutionNameIdx] ?? '';
      const sourceVersion = row[versionIdx] ?? '';
      const extractedVersion = extractVersionFromSolutionName(sourceSolution);
      const cleanedSolution = stripVersionFromSolutionName(sourceSolution) || sourceSolution;

      projected[projectedSolutionIdx] = cleanedSolution;
      projected[projectedVersionIdx] = sourceVersion.trim() || extractedVersion;
    }

    return projected.map((value, idx) => formatCellValueForExcel(nextHeaders[idx], value));
  });

  return {
    title: block.title,
    headers: nextHeaders,
    rows: nextRows,
  };
}

function parseTableRow(line: string): string[] {
  const trimmed = line.trim();
  if (!trimmed.includes('|')) return [];

  let content = trimmed;
  if (content.startsWith('|')) content = content.slice(1);
  if (content.endsWith('|')) content = content.slice(0, -1);

  const cells = content
    .split('|')
    .map((cell) => normalizeTableCell(cell));

  while (cells.length > 0 && cells[cells.length - 1].trim().length === 0) {
    cells.pop();
  }

  return cells;
}

function isMarkdownSeparatorLine(line: string): boolean {
  const cells = parseTableRow(line);
  if (cells.length === 0) return false;
  return cells.every((cell) => /^:?-{3,}:?$/.test(cell.replace(/\s+/g, '')));
}

function headingFromLine(line: string): string | null {
  const match = line.match(/^#{2,6}\s+(.+)$/);
  return match ? stripMarkdownDecorators(match[1].trim()) : null;
}

function headingLevelFromLine(line: string): number | null {
  const match = line.match(/^(#{2,6})\s+/);
  return match ? match[1].length : null;
}

function makeSheetName(base: string, used: Set<string>): string {
  const normalizedBase = base
    .replace(/[\\/*?:[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim() || 'Table';

  const maxLength = 31;
  const root = normalizedBase.slice(0, maxLength);
  let candidate = root;
  let index = 2;

  while (used.has(candidate.toLowerCase())) {
    const suffix = ` (${index})`;
    const trimmedRoot = root.slice(0, Math.max(1, maxLength - suffix.length));
    candidate = `${trimmedRoot}${suffix}`;
    index += 1;
  }

  used.add(candidate.toLowerCase());
  return candidate;
}

export function extractMarkdownTables(markdown: string): MarkdownTable[] {
  const lines = markdown.split(/\r?\n/);
  const tables: MarkdownTable[] = [];
  const headingsByLevel = new Map<number, string>();

  for (let i = 0; i < lines.length; i += 1) {
    const heading = headingFromLine(lines[i]);
    const headingLevel = headingLevelFromLine(lines[i]);
    if (heading && headingLevel) {
      headingsByLevel.set(headingLevel, heading);
      for (let level = headingLevel + 1; level <= 6; level += 1) {
        headingsByLevel.delete(level);
      }
      continue;
    }

    const headerCells = parseTableRow(lines[i]);
    if (headerCells.length === 0) continue;
    if (i + 1 >= lines.length || !isMarkdownSeparatorLine(lines[i + 1])) continue;

    const rows: string[][] = [];
    let cursor = i + 2;

    while (cursor < lines.length) {
      const rowLine = lines[cursor];
      if (!rowLine.trim().includes('|')) break;

      const rowCells = parseTableRow(rowLine);
      if (rowCells.length === 0) break;

      const normalizedRow = headerCells.map((_, colIndex) => rowCells[colIndex] ?? '');
      rows.push(normalizedRow);
      cursor += 1;
    }

    const heading2 = headingsByLevel.get(2) || 'Document';
    const heading3 = headingsByLevel.get(3) || '';
    const heading4 = headingsByLevel.get(4) || '';
    const headingPath = [heading2, heading3, heading4].filter((item) => item.length > 0).join(' > ');

    tables.push({
      title: heading4 || heading3 || heading2,
      heading2,
      heading3,
      heading4,
      headingPath,
      headers: headerCells,
      rows,
    });

    i = cursor - 1;
  }

  return tables;
}

function buildSheetDefinitions(tables: MarkdownTable[]): SheetDefinition[] {
  const usedSheetNames = new Set<string>();
  const grouped = new Map<string, SheetDefinition>();
  const order: string[] = [];

  const getEntitySheetKey = (table: MarkdownTable): { key: string; title: string; sectionTitle: string } | null => {
    if (!/tables\s*&\s*columns/i.test(table.heading2)) return null;
    if (!table.heading3) return null;
    return {
      key: `entity:${table.heading3.toLowerCase()}`,
      title: table.heading3,
      sectionTitle: table.heading2,
    };
  };

  const getRoleSheetKey = (table: MarkdownTable): { key: string; title: string; sectionTitle: string } | null => {
    if (!/security roles/i.test(table.heading2) || !table.heading3) return null;
    return {
      key: `role:${table.heading3.toLowerCase()}`,
      title: `Security Role - ${table.heading3}`,
      sectionTitle: table.heading2,
    };
  };

  tables.forEach((table, index) => {
    const entityGroup = getEntitySheetKey(table);
    const roleGroup = getRoleSheetKey(table);
    const group = entityGroup || roleGroup || {
      key: `table:${index}`,
      title: table.heading3 || table.heading2 || table.title,
      sectionTitle: table.headingPath || table.heading2,
    };

    if (!grouped.has(group.key)) {
      const sheetName = makeSheetName(group.title, usedSheetNames);
      grouped.set(group.key, {
        sheetName,
        title: group.title,
        sectionTitle: group.sectionTitle,
        blocks: [],
        rowCount: 0,
      });
      order.push(group.key);
    }

    const entry = grouped.get(group.key);
    if (!entry) return;

    const blockTitle = table.heading4 || table.title;
    entry.blocks.push({
      title: blockTitle,
      headers: table.headers,
      rows: table.rows,
    });
    entry.rowCount += table.rows.length;
  });

  return order.map((key) => grouped.get(key)).filter((sheet): sheet is SheetDefinition => !!sheet);
}

function setCell(ws: XLSX.WorkSheet, row: number, col: number, value: string, style?: Record<string, unknown>): void {
  const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  ws[addr] = {
    t: 's',
    v: value,
    s: style,
  };

  const nextRefCell = { r: row - 1, c: col - 1 };
  if (!ws['!ref']) {
    ws['!ref'] = `${addr}:${addr}`;
    return;
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  range.s.r = Math.min(range.s.r, nextRefCell.r);
  range.s.c = Math.min(range.s.c, nextRefCell.c);
  range.e.r = Math.max(range.e.r, nextRefCell.r);
  range.e.c = Math.max(range.e.c, nextRefCell.c);
  ws['!ref'] = XLSX.utils.encode_range(range);
}

function updateColumnWidth(columnWidths: number[], col: number, value: string): void {
  const idx = col - 1;
  const base = Math.min(60, Math.max(10, value.split('\n').reduce((max, line) => Math.max(max, line.length), 0) + 2));
  columnWidths[idx] = Math.max(columnWidths[idx] || 0, base);
}

function writeStyledTable(
  ws: XLSX.WorkSheet,
  startRow: number,
  startCol: number,
  headers: string[],
  rows: string[][],
  columnWidths: number[],
): number {
  if (headers.length === 0) {
    return startRow;
  }

  headers.forEach((header, idx) => {
    const col = startCol + idx;
    setCell(ws, startRow, col, header, {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: HEADER_FILL,
      border: BORDER,
      alignment: { vertical: 'center', horizontal: 'left', wrapText: true },
    });
    updateColumnWidth(columnWidths, col, header);
  });

  rows.forEach((rowData, rowIdx) => {
    headers.forEach((_, colIdx) => {
      const col = startCol + colIdx;
      const value = rowData[colIdx] ?? '';
      setCell(ws, startRow + 1 + rowIdx, col, value, {
        fill: rowIdx % 2 === 1 ? ALT_FILL : undefined,
        border: BORDER,
        alignment: { vertical: 'top', horizontal: 'left', wrapText: true },
      });
      updateColumnWidth(columnWidths, col, value);
    });
  });

  ws['!autofilter'] = {
    ref: `${XLSX.utils.encode_cell({ r: startRow - 1, c: startCol - 1 })}:${XLSX.utils.encode_cell({
      r: startRow + Math.max(1, rows.length) - 1,
      c: startCol + headers.length - 1,
    })}`,
  };

  return startRow + 1 + rows.length;
}

function extractDocumentTitle(markdown: string): string {
  const firstHeading = markdown.match(/^#\s+(.+)$/m)?.[1]?.trim() || 'Power Platform Solution';
  if (/^power platform solution\s*:/i.test(firstHeading)) {
    return firstHeading.replace(/^power platform solution\s*:/i, 'Power Platform Solution -').trim();
  }
  if (!/^power platform solution/i.test(firstHeading)) {
    return `Power Platform Solution - ${firstHeading}`;
  }
  return firstHeading;
}

function extractGeneratedOn(markdown: string): string {
  const line = markdown.match(/Generated on:\s*(.+)$/im)?.[1]?.trim();
  if (line) return line;
  return new Date().toLocaleString();
}

function extractDocumentHeaderEntries(markdown: string): DocumentHeaderEntry[] {
  const lines = markdown.split(/\r?\n/);
  const entries: DocumentHeaderEntry[] = [];
  for (const line of lines) {
    if (/^##\s+table of contents/i.test(line)) break;
    const match = line.match(/^-\s+\*\*(.+?)\*\*:\s*(.+)$/);
    if (!match) continue;
    entries.push({ field: stripHtml(match[1].trim()), value: stripHtml(match[2].trim()) });
  }
  return entries;
}

export function buildWorkbookFromMarkdownTables(markdown: string): XLSX.WorkBook | null {
  const tables = extractMarkdownTables(markdown);
  if (tables.length === 0) return null;

  const sheetDefinitions = buildSheetDefinitions(tables);
  const workbook = XLSX.utils.book_new();

  const contentsWs: XLSX.WorkSheet = {};
  const contentsColWidths: number[] = [];
  const title = extractDocumentTitle(markdown);
  const generatedOn = extractGeneratedOn(markdown);
  const docHeaderEntries = extractDocumentHeaderEntries(markdown);

  setCell(contentsWs, 1, 1, title, { font: { bold: true, sz: 22, color: { rgb: '2F62D6' } } });
  setCell(contentsWs, 2, 1, 'Generated by PP-MD - Power Platform Solution Documentation', {
    font: { bold: true, sz: 16, color: { rgb: 'B0005A' } },
  });
  setCell(contentsWs, 3, 1, `Generated on: ${generatedOn}`, {
    font: { bold: true, sz: 14, color: { rgb: 'B0005A' } },
  });

  let row = 5;
  const headerRows = docHeaderEntries.length > 0
    ? docHeaderEntries.map((entry) => [entry.field, entry.value])
    : [['Document', title]];

  headerRows.forEach((entry) => {
    const field = entry[0] || '';
    const value = formatCellValueForExcel(field, entry[1] || '');

    setCell(contentsWs, row, 1, field, {
      font: { bold: true, color: { rgb: '1F4E78' } },
      alignment: { vertical: 'top', horizontal: 'left', wrapText: true },
    });
    setCell(contentsWs, row, 2, value, {
      alignment: { vertical: 'top', horizontal: 'left', wrapText: true },
    });
    updateColumnWidth(contentsColWidths, 1, field);
    updateColumnWidth(contentsColWidths, 2, value);
    row += 1;
  });
  row += 1;

  const tocHeaderRow = row;
  const tocRows = sheetDefinitions.map((sheet) => [
    sheet.title,
    sheet.title,
  ]);
  writeStyledTable(contentsWs, tocHeaderRow, 1, ['Section', 'Link to table'], tocRows, contentsColWidths);

  sheetDefinitions.forEach((sheet, index) => {
    const linkRow = tocHeaderRow + index + 1;
    const cellAddress = XLSX.utils.encode_cell({ r: linkRow - 1, c: 1 });
    const linkText = sheet.title;
    contentsWs[cellAddress] = {
      t: 's',
      v: linkText,
      l: {
        Target: `#'${sheet.sheetName.replace(/'/g, "''")}'!A1`,
        Tooltip: linkText,
      },
      s: {
        font: { color: { rgb: '1F4E78' }, underline: true },
        border: BORDER,
        alignment: { vertical: 'center', horizontal: 'left', wrapText: true },
      },
    };
    updateColumnWidth(contentsColWidths, 2, linkText);
  });

  contentsWs['!cols'] = contentsColWidths.map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(workbook, contentsWs, 'Contents');

  sheetDefinitions.forEach((sheet) => {
    const ws: XLSX.WorkSheet = {};
    const colWidths: number[] = [];
    let sheetRow = 1;

    setCell(ws, sheetRow, 1, sheet.title, {
      font: { bold: true, sz: 15, color: { rgb: '1F4E78' } },
    });
    sheetRow += 2;

    sheet.blocks.forEach((block) => {
      const transformedBlock = transformSheetBlockForExport(sheet.title, block);
      if (transformedBlock.headers.length === 0) {
        return;
      }

      if (normalizeHeaderName(transformedBlock.title) !== normalizeHeaderName(sheet.title)) {
        setCell(ws, sheetRow, 1, transformedBlock.title, {
          font: { bold: true, sz: 12, color: { rgb: '334155' } },
          alignment: { vertical: 'center', horizontal: 'left', wrapText: true },
        });
        sheetRow += 1;
      }

      sheetRow = writeStyledTable(
        ws,
        sheetRow,
        1,
        transformedBlock.headers,
        transformedBlock.rows,
        colWidths,
      ) + 2;
    });

    ws['!cols'] = colWidths.map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(workbook, ws, sheet.sheetName);
  });

  return workbook;
}

function toXlsxFileName(fileName: string): string {
  const trimmed = fileName.trim();
  if (!trimmed) return 'PP-MD Tables.xlsx';
  if (/\.xlsx$/i.test(trimmed)) return trimmed;
  if (/\.[a-z0-9]{1,8}$/i.test(trimmed)) {
    return trimmed.replace(/\.[a-z0-9]{1,8}$/i, '.xlsx');
  }
  return `${trimmed}.xlsx`;
}

export function exportMarkdownTablesToExcel(markdown: string, sourceFileName: string): {
  ok: boolean;
  message: string;
} {
  const workbook = buildWorkbookFromMarkdownTables(markdown);
  if (!workbook) {
    return {
      ok: false,
      message: 'No Markdown tables were found to export.',
    };
  }

  const data = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', cellStyles: true });
  const blob = new Blob([data], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  saveAs(blob, toXlsxFileName(sourceFileName));
  return {
    ok: true,
    message: 'Excel export completed.',
  };
}
