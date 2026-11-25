import JSZip from 'jszip';
import * as XLSX from 'xlsx-js-style';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006';
const tables = new WeakMap<XLSX.WorkSheet, XLSX.Range[]>();
const charts = new WeakMap<XLSX.WorkSheet, WorkbookChart[]>();

export interface WorkbookChart {
  title: string;
  description: string;
  type: 'bar' | 'column';
  categoryTitle: string;
  valueTitle: string;
  anchor: XLSX.Range;
  range: XLSX.Range;
}

function xml(value: string | number): string {
  return String(value).replace(/[&<>"'\t\r\n]/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;',
    '\t': '&#9;', '\r': '&#13;', '\n': '&#10;',
  })[character]!);
}

export function registerExcelTable(sheet: XLSX.WorkSheet, range: XLSX.Range): void {
  if (range.e.r <= range.s.r) return;
  const used = new Set<string>();
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    const ref = XLSX.utils.encode_cell({ r: range.s.r, c: column });
    const cell: XLSX.CellObject = sheet[ref] ?? { t: 's', v: '' };
    const base = (String(cell.v ?? '').trim() || `Column ${column - range.s.c + 1}`).slice(0, 255);
    let name = base;
    for (let suffix = 2; used.has(name.toLowerCase()); suffix += 1) {
      name = `${base.slice(0, 254 - String(suffix).length)} ${suffix}`;
    }
    used.add(name.toLowerCase());
    cell.t = 's';
    cell.v = name;
    delete cell.w;
    sheet[ref] = cell;
  }
  tables.set(sheet, [...(tables.get(sheet) ?? []), range]);
}

export function registerWorkbookChart(sheet: XLSX.WorkSheet, chart: WorkbookChart): void {
  charts.set(sheet, [...(charts.get(sheet) ?? []), chart]);
}

async function requiredXml(zip: JSZip, path: string): Promise<string> {
  const file = zip.file(path);
  if (!file) throw new Error(`Excel export is missing package part: ${path}`);
  return file.async('string');
}

function tableXml(sheet: XLSX.WorkSheet, range: XLSX.Range, id: number, name: string): string {
  const columns = Array.from({ length: range.e.c - range.s.c + 1 }, (_, index) =>
    `<tableColumn id="${index + 1}" name="${xml(String(sheet[XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c + index })]?.v ?? ''))}"/>`,
  ).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<table xmlns="${MAIN}" id="${id}" name="${name}" displayName="${name}" ref="${XLSX.utils.encode_range(range)}" totalsRowShown="0">`
    + `<autoFilter ref="${XLSX.utils.encode_range(range)}"/><tableColumns count="${range.e.c - range.s.c + 1}">${columns}</tableColumns>`
    + '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>';
}

function chartXml(sheetName: string, sheet: XLSX.WorkSheet, chart: WorkbookChart): string {
  const { range } = chart;
  const count = range.e.r - range.s.r;
  const formula = (column: number) => xml(`'${sheetName.replace(/'/g, "''")}'!$${XLSX.utils.encode_col(column)}$${range.s.r + 2}:$${XLSX.utils.encode_col(column)}$${range.e.r + 1}`);
  const points = (column: number) => Array.from({ length: count }, (_, index) =>
    `<c:pt idx="${index}"><c:v>${xml(sheet[XLSX.utils.encode_cell({ r: range.s.r + index + 1, c: column })]?.v ?? '')}</c:v></c:pt>`).join('');
  const title = (label: string) => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-GB" sz="1200" b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr><a:t>${xml(label)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
  const text = '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-GB"/></a:p></c:txPr>';
  const maximum = Math.max(0, ...Array.from({ length: count }, (_, index) =>
    Number(sheet[XLSX.utils.encode_cell({ r: range.s.r + index + 1, c: range.s.c + 1 })]?.v ?? 0)));
  const seriesName = xml(sheet[XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c + 1 })]?.v ?? 'Count');
  const seriesFormula = xml(`'${sheetName.replace(/'/g, "''")}'!$${XLSX.utils.encode_col(range.s.c + 1)}$${range.s.r + 1}`);
  const horizontal = chart.type === 'bar';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<c:chartSpace xmlns:c="${DRAWING}/chart" xmlns:a="${DRAWING}/main"><c:lang val="en-GB"/><c:chart>${title(chart.title)}<c:autoTitleDeleted val="0"/>`
    + `<c:plotArea><c:layout/><c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>`
    + `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>${seriesFormula}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${seriesName}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
    + '<c:spPr><a:solidFill><a:srgbClr val="1D4ED8"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>'
    + `<c:cat><c:strRef><c:f>${formula(range.s.c)}</c:f><c:strCache><c:ptCount val="${count}"/>${points(range.s.c)}</c:strCache></c:strRef></c:cat>`
    + `<c:val><c:numRef><c:f>${formula(range.s.c + 1)}</c:f><c:numCache><c:formatCode>0</c:formatCode><c:ptCount val="${count}"/>${points(range.s.c + 1)}</c:numCache></c:numRef></c:val></c:ser>`
    + `<c:dLbls>${text}<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/></c:dLbls>`
    + '<c:gapWidth val="80"/><c:axId val="1"/><c:axId val="2"/></c:barChart>'
    + `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="${horizontal ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/>${title(chart.categoryTitle)}<c:tickLblPos val="nextTo"/>${text}<c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`
    + `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/>${title(chart.valueTitle)}<c:numFmt formatCode="0" sourceLinked="0"/><c:tickLblPos val="nextTo"/>${text}<c:crossAx val="1"/><c:crosses val="${horizontal ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/><c:majorUnit val="${Math.max(1, Math.ceil(maximum / 5))}"/></c:valAx>`
    + `</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
    + `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></c:spPr>${text}</c:chartSpace>`;
}

export async function packageExcelWorkbook(workbook: XLSX.WorkBook, data: ArrayBuffer): Promise<Blob> {
  const zip = await JSZip.loadAsync(data);
  const fontOrder = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
  const styles = (await requiredXml(zip, 'xl/styles.xml'))
    .replace(/\brgb="([A-Fa-f0-9]{6})"/g, 'rgb="FF$1"')
    .replace(/<font>([\s\S]*?)<\/font>/g, (_, content: string) => {
      const children = content.match(/<[^>]+\/>/g) ?? [];
      if (content.replace(/<[^>]+\/>/g, '').trim()) {
        throw new Error('Excel export encountered an unsupported font definition.');
      }
      children.sort((left, right) => fontOrder.indexOf(left.match(/^<(\w+)/)![1])
        - fontOrder.indexOf(right.match(/^<(\w+)/)![1]));
      return `<font>${children.join('')}</font>`;
    });
  zip.file('xl/styles.xml', styles);
  let contentTypes = await requiredXml(zip, '[Content_Types].xml');
  let tableId = 0;
  let chartId = 0;
  const tableNames = new Set<string>();
  const override = (path: string, type: string) => {
    contentTypes = contentTypes.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/></Types>`);
  };
  for (const [index, sheetName] of workbook.SheetNames.entries()) {
    const ranges = tables.get(workbook.Sheets[sheetName]) ?? [];
    const path = `xl/worksheets/sheet${index + 1}.xml`;
    let source = await requiredXml(zip, path);
    const relationshipsPath = `xl/worksheets/_rels/sheet${index + 1}.xml.rels`;
    let relationships = zip.file(relationshipsPath) ? await requiredXml(zip, relationshipsPath)
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const parts = ranges.map((range, partIndex) => {
      tableId += 1;
      const base = `table_${sheetName.replace(/[^A-Za-z0-9_]/g, '_')}_${partIndex + 1}`;
      let name = base;
      for (let suffix = 2; tableNames.has(name.toLowerCase()); suffix += 1) name = `${base}_${suffix}`;
      tableNames.add(name.toLowerCase());
      const tablePath = `xl/tables/table${tableId}.xml`;
      zip.file(tablePath, tableXml(workbook.Sheets[sheetName], range, tableId, name));
      override(tablePath, 'spreadsheetml.table');
      const id = `ppmdTable${tableId}`;
      relationships = relationships.replace('</Relationships>', `<Relationship Id="${id}" Type="${REL}/table" Target="../tables/table${tableId}.xml"/></Relationships>`);
      return `<tablePart r:id="${id}"/>`;
    }).join('');
    if (parts) {
      source = source.replace('</worksheet>', `<tableParts count="${ranges.length}">${parts}</tableParts></worksheet>`);
      zip.file(relationshipsPath, relationships);
    }
    const freeze: { ySplit?: number } | undefined = workbook.Sheets[sheetName]['!freeze'];
    if (freeze?.ySplit) {
      const view = `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${freeze.ySplit}" topLeftCell="A${freeze.ySplit + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${freeze.ySplit + 1}" sqref="A${freeze.ySplit + 1}"/></sheetView></sheetViews>`;
      source = /<sheetViews\b/.test(source)
        ? source.replace(/<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/, view)
        : source.replace(/(<dimension\b[^>]*\/>)/, `$1${view}`);
    }
    const sheetCharts = charts.get(workbook.Sheets[sheetName]) ?? [];
    if (sheetCharts.length) {
      const drawingId = index + 1;
      const drawingPath = `xl/drawings/drawing${drawingId}.xml`;
      const drawingParts: string[] = [];
      const drawingRelationships: string[] = [];
      for (const chart of sheetCharts) {
        chartId += 1;
        const chartPath = `xl/charts/chart${chartId}.xml`;
        zip.file(chartPath, chartXml(sheetName, workbook.Sheets[sheetName], chart));
        override(chartPath, 'drawingml.chart');
        drawingRelationships.push(`<Relationship Id="chart${chartId}" Type="${REL}/chart" Target="../charts/chart${chartId}.xml"/>`);
        const anchor = (row: number, column: number) => `<xdr:col>${column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
        drawingParts.push(`<xdr:twoCellAnchor><xdr:from>${anchor(chart.anchor.s.r, chart.anchor.s.c)}</xdr:from><xdr:to>${anchor(chart.anchor.e.r, chart.anchor.e.c)}</xdr:to>`
          + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${chartId}" name="${xml(chart.title)}" title="${xml(chart.title)}" descr="${xml(`${chart.description} Data table: '${sheetName}'!${XLSX.utils.encode_range(chart.range)}.`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
          + `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${DRAWING}/chart"><c:chart xmlns:c="${DRAWING}/chart" xmlns:r="${REL}" r:id="chart${chartId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`);
      }
      zip.file(drawingPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${DRAWING}/spreadsheetDrawing" xmlns:a="${DRAWING}/main">${drawingParts.join('')}</xdr:wsDr>`);
      zip.file(`xl/drawings/_rels/drawing${drawingId}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRelationships.join('')}</Relationships>`);
      override(drawingPath, 'drawing');
      const id = `ppmdDrawing${drawingId}`;
      relationships = relationships.replace('</Relationships>', `<Relationship Id="${id}" Type="${REL}/drawing" Target="../drawings/drawing${drawingId}.xml"/></Relationships>`);
      const drawing = `<drawing r:id="${id}"/>`;
      source = source.includes('<tableParts')
        ? source.replace('<tableParts', `${drawing}<tableParts`)
        : source.replace('</worksheet>', `${drawing}</worksheet>`);
      zip.file(relationshipsPath, relationships);
    }
    zip.file(path, source);
  }
  zip.file('[Content_Types].xml', contentTypes);
  return new Blob([await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
