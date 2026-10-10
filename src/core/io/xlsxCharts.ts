// Reads and writes native Excel charts (DrawingML) inside .xlsx packages.
import JSZip from 'jszip';
import { chartData } from '../charts';
import type { Doc } from '../doc';
import type { ChartDef, ChartSeries, ChartType, Sheet, Workbook } from '../workbook';

const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const EMU = 9525;

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function unesc(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_m, d) => String.fromCharCode(+d))
    .replace(/&amp;/g, '&');
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

function relsPathFor(path: string): string {
  const i = path.lastIndexOf('/');
  return `${path.slice(0, i)}/_rels/${path.slice(i + 1)}.rels`;
}

interface Rel {
  id: string;
  type: string;
  target: string;
}

function parseRels(xml: string): Rel[] {
  const out: Rel[] = [];
  const re = /<Relationship\b([^>]*)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const attr = (n: string) => new RegExp(`\\b${n}="([^"]*)"`).exec(m![1])?.[1] ?? '';
    out.push({ id: attr('Id'), type: attr('Type'), target: unesc(attr('Target')) });
  }
  return out;
}

async function sheetPaths(zip: JSZip): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!wbXml || !relsXml) return out;
  const rels = parseRels(relsXml);
  const re = /<sheet\b([^>]*)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(wbXml))) {
    const name = unesc(/\bname="([^"]*)"/.exec(m[1])?.[1] ?? '');
    const rid = /\br:id="([^"]*)"/.exec(m[1])?.[1] ?? /\bid="([^"]*)"/.exec(m[1])?.[1];
    const rel = rels.find((r) => r.id === rid);
    if (rel) out.set(name, resolvePath('xl/workbook.xml', rel.target));
  }
  return out;
}

// ---------- writing ----------

function colsToEmu(sheet: Sheet, startCol: number, startOff: number, widthPx: number): { col: number; off: number } {
  let col = startCol;
  let x = startOff + widthPx;
  for (let guard = 0; guard < 16384; guard++) {
    const w = sheet.isColHidden(col) ? 0 : sheet.colWidth(col);
    if (w > 0 && x <= w) break;
    x -= w;
    col++;
  }
  return { col, off: Math.max(0, Math.round(x)) };
}

function rowsToEmu(sheet: Sheet, startRow: number, startOff: number, heightPx: number): { row: number; off: number } {
  let row = startRow;
  let y = startOff + heightPx;
  for (let guard = 0; guard < 1_048_576; guard++) {
    const h = sheet.isRowHidden(row) ? 0 : sheet.rowHeight(row);
    if (h > 0 && y <= h) break;
    y -= h;
    row++;
  }
  return { row, off: Math.max(0, Math.round(y)) };
}

function anchorXml(sheet: Sheet, chart: ChartDef, relId: string, idx: number): string {
  const a = chart.anchor;
  const to = colsToEmu(sheet, a.col, a.dx, a.w);
  const toRow = rowsToEmu(sheet, a.row, a.dy, a.h);
  return `<xdr:twoCellAnchor editAs="oneCell">
<xdr:from><xdr:col>${a.col}</xdr:col><xdr:colOff>${Math.round(a.dx * EMU)}</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>${Math.round(a.dy * EMU)}</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.off * EMU}</xdr:colOff><xdr:row>${toRow.row}</xdr:row><xdr:rowOff>${toRow.off * EMU}</xdr:rowOff></xdr:to>
<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${idx + 2}" name="Chart ${idx + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relId}"/></a:graphicData></a:graphic>
</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

function richText(text: string, size = 1400, bold = false): string {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="${bold ? 1 : 0}"/></a:pPr><a:r><a:rPr lang="en-US" sz="${size}" b="${bold ? 1 : 0}"/><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
}

function titleXml(text: string | undefined, size = 1400): string {
  if (!text) return '';
  return `<c:title>${richText(text, size)}<c:overlay val="0"/></c:title>`;
}

function refXml(kind: 'num' | 'str', ref: string | undefined, cache: (string | number | null)[]): string {
  if (!ref) return '';
  const pts = cache
    .map((v, i) => (v === null || v === '' ? '' : `<c:pt idx="${i}"><c:v>${esc(String(v))}</c:v></c:pt>`))
    .join('');
  if (kind === 'num') {
    return `<c:numRef><c:f>${esc(ref)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${cache.length}"/>${pts}</c:numCache></c:numRef>`;
  }
  return `<c:strRef><c:f>${esc(ref)}</c:f><c:strCache><c:ptCount val="${cache.length}"/>${pts}</c:strCache></c:strRef>`;
}

function seriesNameXml(se: ChartSeries, resolvedName: string): string {
  if (!se.name) return '';
  if (se.name.startsWith('=')) return `<c:tx>${refXml('str', se.name.slice(1), [resolvedName])}</c:tx>`;
  return `<c:tx><c:v>${esc(se.name)}</c:v></c:tx>`;
}

function fill(color: string, line = false): string {
  const c = color.replace('#', '').toUpperCase();
  if (line) return `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:round/></a:ln></c:spPr>`;
  return `<c:spPr><a:solidFill><a:srgbClr val="${c}"/></a:solidFill></c:spPr>`;
}

function chartXml(doc: Doc, sheet: Sheet, chart: ChartDef): string {
  const data = chartData(doc, chart, sheet);
  const t = chart.type;
  const catAxId = 50010;
  const valAxId = 50020;
  const sers = chart.series.map((se, i) => {
    const d = data.series[i];
    const head = `<c:idx val="${i}"/><c:order val="${i}"/>${seriesNameXml(se, d?.name ?? '')}`;
    const color = d?.color ?? '#4ea1ff';
    const cat = se.categories ? `<c:cat>${refXml(data.xs.some((x) => x !== null) && t === 'scatter' ? 'num' : 'str', se.categories, data.categories)}</c:cat>` : '';
    const val = `<c:val>${refXml('num', se.values, d?.values ?? [])}</c:val>`;
    switch (t) {
      case 'column':
      case 'bar':
        return `<c:ser>${head}${fill(color)}<c:invertIfNegative val="0"/>${cat}${val}</c:ser>`;
      case 'line':
        return `<c:ser>${head}${fill(color, true)}<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>${cat}${val}<c:smooth val="0"/></c:ser>`;
      case 'area':
        return `<c:ser>${head}${fill(color)}${cat}${val}</c:ser>`;
      case 'pie':
      case 'doughnut': {
        const pts = (d?.values ?? []).map((_, k) => {
          const pc = ['#4ea1ff', '#ff9f43', '#3fcf8e', '#e8618c', '#a78bfa', '#ffd43b', '#22b8cf', '#ff6b6b', '#94d82d', '#c4a7ff'][k % 10].replace('#', '').toUpperCase();
          return `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${pc}"/></a:solidFill></c:spPr></c:dPt>`;
        });
        return `<c:ser>${head}${pts.join('')}${cat}${val}</c:ser>`;
      }
      case 'scatter': {
        const x = se.categories ? `<c:xVal>${refXml('num', se.categories, data.xs)}</c:xVal>` : '';
        return `<c:ser>${head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="${color.replace('#', '').toUpperCase()}"/></a:solidFill></c:spPr></c:marker>${x}<c:yVal>${refXml('num', se.values, d?.values ?? [])}</c:yVal><c:smooth val="0"/></c:ser>`;
      }
    }
  });
  const grouping = chart.stacked ? 'stacked' : t === 'line' || t === 'area' ? 'standard' : 'clustered';
  let plot = '';
  const axes = (horizontal: boolean, scatter: boolean) => {
    const catPos = horizontal ? 'l' : 'b';
    const valPos = horizontal ? 'b' : 'l';
    const catTitle = titleXml(chart.xTitle, 1000);
    const valTitle = titleXml(chart.yTitle, 1000);
    const catAx = scatter
      ? `<c:valAx><c:axId val="${catAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${catTitle}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${valAxId}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
      : `<c:catAx><c:axId val="${catAxId}"/><c:scaling><c:orientation val="${horizontal ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/>${catTitle}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${valAxId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = `<c:valAx><c:axId val="${valAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/><c:majorGridlines/>${valTitle}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${catAxId}"/><c:crosses val="${horizontal ? 'max' : 'autoZero'}"/><c:crossBetween val="${scatter ? 'midCat' : 'between'}"/></c:valAx>`;
    return catAx + valAx;
  };
  const axIds = `<c:axId val="${catAxId}"/><c:axId val="${valAxId}"/>`;
  switch (t) {
    case 'column':
    case 'bar':
      plot = `<c:barChart><c:barDir val="${t === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers.join('')}<c:gapWidth val="80"/>${chart.stacked ? '<c:overlap val="100"/>' : ''}${axIds}</c:barChart>${axes(t === 'bar', false)}`;
      break;
    case 'line':
      plot = `<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers.join('')}<c:marker val="1"/>${axIds}</c:lineChart>${axes(false, false)}`;
      break;
    case 'area':
      plot = `<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers.join('')}${axIds}</c:areaChart>${axes(false, false)}`;
      break;
    case 'pie':
      plot = `<c:pieChart><c:varyColors val="1"/>${sers.slice(0, 1).join('')}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case 'doughnut':
      plot = `<c:doughnutChart><c:varyColors val="1"/>${sers.slice(0, 1).join('')}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
      break;
    case 'scatter':
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers.join('')}${axIds}</c:scatterChart>${axes(false, true)}`;
      break;
  }
  const legendPos = { right: 'r', bottom: 'b', top: 't', none: '' }[chart.legend];
  const legend = legendPos ? `<c:legend><c:legendPos val="${legendPos}"/><c:overlay val="0"/></c:legend>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_REL}"><c:roundedCorners val="0"/><c:chart>${titleXml(chart.title)}<c:autoTitleDeleted val="${chart.title ? 0 : 1}"/><c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

const DRAWING_SLOT_BEFORE = ['<legacyDrawing ', '<legacyDrawing>', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst'];

export async function writeCharts(buf: Uint8Array, doc: Doc): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(buf);
  const paths = await sheetPaths(zip);
  let ct = (await zip.file('[Content_Types].xml')!.async('string')) ?? '';
  let drawingN = Object.keys(zip.files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f)).length;
  let chartN = Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f)).length;
  for (const sheet of doc.wb.sheets) {
    if (!sheet.charts.length) continue;
    const path = paths.get(sheet.name);
    if (!path) continue;
    let sheetXml = await zip.file(path)!.async('string');
    if (/<drawing\b/.test(sheetXml)) continue; // sheet already has a drawing part (pictures) — don't corrupt it
    drawingN++;
    const drawingPath = `xl/drawings/drawing${drawingN}.xml`;
    const relsPath = relsPathFor(path);
    let rels = (await zip.file(relsPath)?.async('string')) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const used = new Set(parseRels(rels).map((r) => r.id));
    let k = 1;
    while (used.has(`rId${k}`)) k++;
    const drawingRelId = `rId${k}`;
    rels = rels.replace('</Relationships>', `<Relationship Id="${drawingRelId}" Type="${REL_DRAWING}" Target="../drawings/drawing${drawingN}.xml"/></Relationships>`);
    zip.file(relsPath, rels);

    const tag = `<drawing r:id="${drawingRelId}"/>`;
    let pos = -1;
    for (const marker of DRAWING_SLOT_BEFORE) {
      const i = sheetXml.indexOf(marker);
      if (i >= 0 && (pos < 0 || i < pos)) pos = i;
    }
    if (pos < 0) pos = sheetXml.lastIndexOf('</worksheet>');
    sheetXml = sheetXml.slice(0, pos) + tag + sheetXml.slice(pos);
    if (!/xmlns:r=/.test(sheetXml.slice(0, 600))) sheetXml = sheetXml.replace('<worksheet ', `<worksheet xmlns:r="${NS_REL}" `);
    zip.file(path, sheetXml);

    const anchors: string[] = [];
    const drawingRels: string[] = [];
    sheet.charts.forEach((chart, i) => {
      chartN++;
      const rid = `rId${i + 1}`;
      drawingRels.push(`<Relationship Id="${rid}" Type="${REL_CHART}" Target="../charts/chart${chartN}.xml"/>`);
      anchors.push(anchorXml(sheet, chart, rid, i));
      zip.file(`xl/charts/chart${chartN}.xml`, chartXml(doc, sheet, chart));
      ct = ct.replace('</Types>', `<Override PartName="/xl/charts/chart${chartN}.xml" ContentType="${CT_CHART}"/></Types>`);
    });
    zip.file(
      drawingPath,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_REL}">${anchors.join('')}</xdr:wsDr>`,
    );
    zip.file(
      relsPathFor(drawingPath),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRels.join('')}</Relationships>`,
    );
    ct = ct.replace('</Types>', `<Override PartName="/${drawingPath}" ContentType="${CT_DRAWING}"/></Types>`);
  }
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

// ---------- reading ----------

/**
 * ExcelJS crashes on drawing and comment parts written by some tools (other Excel generators use
 * different part names). Charts and notes are read separately, so remove those parts from the copy
 * given to ExcelJS. Returns null when there's nothing to remove.
 */
export async function prepareForExcelJS(data: Uint8Array): Promise<Uint8Array | null> {
  const zip = await JSZip.loadAsync(data);
  const extra = Object.keys(zip.files).filter((f) => /^xl\/(drawings|comments|threadedComments|persons)\//i.test(f) || /^xl\/comments\d*\.xml$/i.test(f));
  if (!extra.length) return null;
  for (const f of extra) zip.remove(f);
  for (const f of Object.keys(zip.files)) {
    if (/^xl\/worksheets\/sheet[^/]*\.xml$/.test(f)) {
      const xml = await zip.file(f)!.async('string');
      const next = xml.replace(/<(drawing|legacyDrawing|legacyDrawingHF)\b[^>]*\/>/g, '');
      if (next !== xml) zip.file(f, next);
    } else if (/^xl\/worksheets\/_rels\/.*\.rels$/.test(f)) {
      const xml = await zip.file(f)!.async('string');
      const next = xml.replace(/<Relationship\b[^>]*relationships\/(drawing|vmlDrawing|comments|threadedComment)"[^>]*\/>/g, '');
      if (next !== xml) zip.file(f, next);
    }
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
}

/** Reads cell notes (legacy comments) for every sheet. */
export async function readNotes(data: Uint8Array, wb: Workbook): Promise<void> {
  const zip = await JSZip.loadAsync(data);
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    if (!path) continue;
    const relsXml = await zip.file(relsPathFor(path))?.async('string');
    if (!relsXml) continue;
    for (const rel of parseRels(relsXml).filter((r) => r.type.endsWith('/comments'))) {
      const xml = await zip.file(resolvePath(path, rel.target))?.async('string');
      if (!xml) continue;
      const re = /<comment\b[^>]*\bref="([A-Z]+)(\d+)"[^>]*>([\s\S]*?)<\/comment>/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(xml))) {
        const col = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
        const row = parseInt(m[2], 10) - 1;
        const parts: string[] = [];
        const tre = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
        let t: RegExpExecArray | null;
        while ((t = tre.exec(m[3]))) parts.push(unesc(t[1]));
        const note = parts.join('').trim();
        if (!note) continue;
        const key = row * 16384 + col;
        sheet.cells.set(key, { ...(sheet.cells.get(key) ?? {}), note });
      }
    }
  }
}

/** Removes namespace prefixes from element names so files with and without prefixes parse the same. */
function unprefix(xml: string): string {
  return xml.replace(/<(\/?)[A-Za-z][\w.-]*:/g, '<$1');
}

function textOf(xml: string): string {
  const parts: string[] = [];
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) parts.push(unesc(m[1]));
  return parts.join('');
}

function refIn(xml: string | undefined): string | undefined {
  if (!xml) return undefined;
  const m = /<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/.exec(xml);
  return m ? unesc(m[1]).trim() : undefined;
}

function block(xml: string, tag: string): string | undefined {
  const m = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`).exec(xml);
  return m?.[1];
}

export function sizeFromAnchor(sheet: Sheet, xml: string): ChartDef['anchor'] | null {
  const num = (s: string | undefined) => (s ? parseInt(s, 10) : 0);
  const from = /<from>([\s\S]*?)<\/from>/.exec(xml)?.[1];
  const to = /<to>([\s\S]*?)<\/to>/.exec(xml)?.[1];
  const ext = /<ext\s+cx="(\d+)"\s+cy="(\d+)"/.exec(xml);
  if (!from) {
    const pos = /<pos\s+x="(\d+)"\s+y="(\d+)"/.exec(xml);
    if (!pos || !ext) return null;
    return { row: 0, col: 0, dx: +pos[1] / EMU, dy: +pos[2] / EMU, w: +ext[1] / EMU, h: +ext[2] / EMU };
  }
  const get = (src: string, tag: string) => num(new RegExp(`<${tag}>(-?\\d+)</${tag}>`).exec(src)?.[1]);
  const col = get(from, 'col');
  const row = get(from, 'row');
  const dx = get(from, 'colOff') / EMU;
  const dy = get(from, 'rowOff') / EMU;
  let w = 480;
  let h = 300;
  if (to) {
    const tc = get(to, 'col');
    const tr = get(to, 'row');
    let x = 0;
    for (let c = col; c < tc; c++) x += sheet.isColHidden(c) ? 0 : sheet.colWidth(c);
    w = x - dx + get(to, 'colOff') / EMU;
    let y = 0;
    for (let r = row; r < tr; r++) y += sheet.isRowHidden(r) ? 0 : sheet.rowHeight(r);
    h = y - dy + get(to, 'rowOff') / EMU;
  } else if (ext) {
    w = +ext[1] / EMU;
    h = +ext[2] / EMU;
  }
  return { row, col, dx: Math.round(dx), dy: Math.round(dy), w: Math.max(160, Math.round(w)), h: Math.max(120, Math.round(h)) };
}

export function parseChart(raw: string): Omit<ChartDef, 'anchor' | 'id'> | null {
  const xml = unprefix(raw);
  const typeMatch = /<(barChart|bar3DChart|lineChart|line3DChart|areaChart|area3DChart|pieChart|pie3DChart|doughnutChart|scatterChart)>/.exec(xml);
  if (!typeMatch) return null;
  const kind = typeMatch[1];
  let type: ChartType;
  if (kind.startsWith('bar')) type = /<barDir val="bar"\/>/.test(xml) ? 'bar' : 'column';
  else if (kind.startsWith('line')) type = 'line';
  else if (kind.startsWith('area')) type = 'area';
  else if (kind.startsWith('pie')) type = 'pie';
  else if (kind === 'doughnutChart') type = 'doughnut';
  else type = 'scatter';
  const stacked = /<grouping val="(stacked|percentStacked)"\/>/.test(xml);
  const series: ChartSeries[] = [];
  const serRe = /<ser>([\s\S]*?)<\/ser>/g;
  let m: RegExpExecArray | null;
  while ((m = serRe.exec(xml))) {
    const s = m[1];
    const tx = block(s, 'tx');
    let name: string | undefined;
    if (tx) {
      const r = refIn(tx);
      if (r) name = '=' + r;
      else {
        const v = /<v>([\s\S]*?)<\/v>/.exec(tx)?.[1];
        if (v) name = unesc(v);
      }
    }
    const cats = refIn(block(s, type === 'scatter' ? 'xVal' : 'cat'));
    const vals = refIn(block(s, type === 'scatter' ? 'yVal' : 'val'));
    if (!vals) continue;
    const head = s.split(/<(?:cat|val|xVal|yVal|dPt)>/)[0];
    const spPr = /<spPr>([\s\S]*?)<\/spPr>/.exec(head)?.[1];
    const color = spPr ? /<srgbClr val="([0-9A-Fa-f]{6})"/.exec(spPr)?.[1] : undefined;
    series.push({ name, categories: cats, values: vals, color: color && type !== 'pie' && type !== 'doughnut' ? '#' + color.toLowerCase() : undefined });
  }
  if (!series.length) return null;
  const chartBody = /<chart>([\s\S]*)<\/chart>/.exec(xml)?.[1] ?? xml;
  const plotStart = chartBody.indexOf('<plotArea>');
  const head = plotStart >= 0 ? chartBody.slice(0, plotStart) : '';
  const titleXmlPart = /<title>([\s\S]*?)<\/title>/.exec(head)?.[1];
  let title = titleXmlPart ? textOf(titleXmlPart) : undefined;
  if (!title && titleXmlPart && series.length === 1 && series[0].name && !series[0].name.startsWith('=')) title = series[0].name;
  const legendPos = /<legend>[\s\S]*?<legendPos val="(\w)"\/>/.exec(chartBody)?.[1];
  const legend: ChartDef['legend'] = !/<legend>/.test(chartBody) ? 'none' : legendPos === 't' ? 'top' : legendPos === 'b' ? 'bottom' : 'right';
  const axisTitle = (tag: string) => {
    const ax = new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`, 'g');
    const found: string[] = [];
    let am: RegExpExecArray | null;
    while ((am = ax.exec(chartBody))) {
      const t = /<title>([\s\S]*?)<\/title>/.exec(am[1]);
      found.push(t ? textOf(t[1]) : '');
    }
    return found;
  };
  const catTitles = axisTitle('catAx');
  const valTitles = axisTitle('valAx');
  const xTitle = type === 'scatter' ? valTitles[0] : catTitles[0];
  const yTitle = type === 'scatter' ? valTitles[1] : valTitles[0];
  return { type, title: title || undefined, series, stacked, legend, xTitle: xTitle || undefined, yTitle: yTitle || undefined };
}

export async function readCharts(data: Uint8Array, wb: Workbook): Promise<{ unsupported: boolean; pictures: boolean }> {
  const zip = await JSZip.loadAsync(data);
  const paths = await sheetPaths(zip);
  let unsupported = false;
  let pictures = false;
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    if (!path) continue;
    const relsXml = await zip.file(relsPathFor(path))?.async('string');
    if (!relsXml) continue;
    for (const rel of parseRels(relsXml).filter((r) => r.type.endsWith('/drawing'))) {
      const drawingPath = resolvePath(path, rel.target);
      const rawDrawing = await zip.file(drawingPath)?.async('string');
      const drawingRels = parseRels((await zip.file(relsPathFor(drawingPath))?.async('string')) ?? '');
      if (!rawDrawing) continue;
      const drawingXml = unprefix(rawDrawing);
      if (/<pic\b/.test(drawingXml)) pictures = true;
      const anchorRe = /<(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/\1>/g;
      let am: RegExpExecArray | null;
      while ((am = anchorRe.exec(drawingXml))) {
        const anchorXmlText = am[0];
        const rid = /<chart\b[^>]*r:id="([^"]+)"/.exec(anchorXmlText)?.[1];
        if (!rid) continue;
        const chartRel = drawingRels.find((r) => r.id === rid);
        if (!chartRel) continue;
        const chartXmlText = await zip.file(resolvePath(drawingPath, chartRel.target))?.async('string');
        if (!chartXmlText) continue;
        const parsed = parseChart(chartXmlText);
        const anchor = sizeFromAnchor(sheet, anchorXmlText);
        if (!parsed || !anchor) {
          unsupported = true;
          continue;
        }
        sheet.charts.push({ id: 'chart-' + Math.random().toString(36).slice(2, 10), anchor, ...parsed });
      }
    }
  }
  return { unsupported, pictures };
}
