// Tables in .xlsx files (xl/tables/tableN.xml, one part per table, listed in each sheet's
// <tableParts>). Excel opens what Gridline writes as real tables, sort order and filters included.
import JSZip from 'jszip';
import { formatRange, parseRange, type Range } from '../address';
import { structuredFromFile } from '../tableRefs';
import { TABLE_COLOR_IDS } from '../tables';
import type { FilterCond, FilterOp, Sheet, TableColor, TableDef, Workbook } from '../workbook';
import { esc, parseRels, relsPathFor, resolvePath, sheetPaths, topLevelSlot, unesc } from './xlsxCharts';

const REL_TABLE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Excel's built-in styles in each color (the accent colors of the default theme). */
/** Excel's table style for each color ('' = no style: the cells keep their own formats). */
const STYLE_FOR: Record<TableColor, string> = {
  none: '',
  blue: 'TableStyleMedium2',
  orange: 'TableStyleMedium3',
  gray: 'TableStyleMedium4',
  gold: 'TableStyleMedium5',
  green: 'TableStyleMedium7',
};

/** The nearest color for one of Excel's table styles (TableStyleLight9, TableStyleMedium2…). */
export function colorForStyle(name: string | undefined): TableColor {
  if (!name) return 'none';
  const m = /^TableStyle(?:Light|Medium|Dark)(\d+)$/.exec(name);
  if (!m) return 'blue';
  // Each style family cycles through: none/black, then accents 1–6
  const accent = (parseInt(m[1], 10) - 1) % 7;
  return (['gray', 'blue', 'orange', 'gray', 'gold', 'blue', 'green'] as TableColor[])[accent];
}

function attrs(s: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of s.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = unesc(m[2]);
  return out;
}

/** Column names in files use _x000a_ for line breaks. */
const fromXmlName = (s: string) => s.replace(/_x([0-9a-f]{4})_/gi, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
const toXmlName = (s: string) => s.replace(/[\n\r\t]/g, (ch) => `_x${ch.charCodeAt(0).toString(16).padStart(4, '0')}_`);

// ---------- filters ----------

const OP_TO_XML: Partial<Record<FilterOp, string>> = { equals: 'equal', notEquals: 'notEqual', gt: 'greaterThan', gte: 'greaterThanOrEqual', lt: 'lessThan', lte: 'lessThanOrEqual' };

function filterXml(colId: number, cond: FilterCond): string {
  if (cond.values) {
    const blank = cond.values.includes('');
    const vals = cond.values.filter((v) => v !== '').map((v) => `<filter val="${esc(v)}"/>`);
    return `<filterColumn colId="${colId}"><filters${blank ? ' blank="1"' : ''}>${vals.join('')}</filters></filterColumn>`;
  }
  const op = cond.op;
  if (!op) return '';
  const custom = (operator: string | undefined, val: string) => `<customFilter${operator ? ` operator="${operator}"` : ''} val="${esc(val)}"/>`;
  switch (op.type) {
    case 'contains':
      return `<filterColumn colId="${colId}"><customFilters>${custom(undefined, `*${op.value}*`)}</customFilters></filterColumn>`;
    case 'notContains':
      return `<filterColumn colId="${colId}"><customFilters>${custom('notEqual', `*${op.value}*`)}</customFilters></filterColumn>`;
    case 'beginsWith':
      return `<filterColumn colId="${colId}"><customFilters>${custom(undefined, `${op.value}*`)}</customFilters></filterColumn>`;
    case 'endsWith':
      return `<filterColumn colId="${colId}"><customFilters>${custom(undefined, `*${op.value}`)}</customFilters></filterColumn>`;
    case 'between':
      return `<filterColumn colId="${colId}"><customFilters and="1">${custom('greaterThanOrEqual', op.value)}${custom('lessThanOrEqual', op.value2 ?? op.value)}</customFilters></filterColumn>`;
    case 'top10':
      return `<filterColumn colId="${colId}"><top10 val="${Math.max(1, parseInt(op.value, 10) || 10)}"/></filterColumn>`;
    case 'aboveAvg':
    case 'belowAvg':
      return `<filterColumn colId="${colId}"><dynamicFilter type="${op.type === 'aboveAvg' ? 'aboveAverage' : 'belowAverage'}"/></filterColumn>`;
    default: {
      const xmlOp = OP_TO_XML[op.type];
      return xmlOp ? `<filterColumn colId="${colId}"><customFilters>${custom(xmlOp === 'equal' ? undefined : xmlOp, op.value)}</customFilters></filterColumn>` : '';
    }
  }
}

function filterFromXml(body: string): FilterCond | null {
  const filters = /<filters\b([^>]*)>([\s\S]*?)<\/filters>|<filters\b([^>]*)\/>/.exec(body);
  if (filters) {
    const a = attrs(filters[1] ?? filters[3] ?? '');
    const values = [...(filters[2] ?? '').matchAll(/<filter\b[^>]*\bval="([^"]*)"/g)].map((m) => unesc(m[1]));
    if (a.blank === '1' || a.blank === 'true') values.push('');
    return { values };
  }
  const top = /<top10\b([^>]*)\/?>/.exec(body);
  if (top) {
    const a = attrs(top[1]);
    if (a.percent === '1' || a.top === '0') return null;
    return { op: { type: 'top10', value: a.val ?? '10' } };
  }
  const dyn = /<dynamicFilter\b[^>]*\btype="(aboveAverage|belowAverage)"/.exec(body);
  if (dyn) return { op: { type: dyn[1] === 'aboveAverage' ? 'aboveAvg' : 'belowAvg', value: '' } };
  const customs = [...body.matchAll(/<customFilter\b([^>]*)\/?>/g)].map((m) => attrs(m[1]));
  if (customs.length === 2 && customs[0].operator === 'greaterThanOrEqual' && customs[1].operator === 'lessThanOrEqual') return { op: { type: 'between', value: customs[0].val ?? '', value2: customs[1].val ?? '' } };
  if (customs.length !== 1) return null;
  const { operator = 'equal', val = '' } = customs[0];
  const wild = /^\*(.*)\*$/.exec(val);
  if (wild) return { op: { type: operator === 'notEqual' ? 'notContains' : 'contains', value: wild[1] } };
  if (operator === 'equal' && val.endsWith('*')) return { op: { type: 'beginsWith', value: val.slice(0, -1) } };
  if (operator === 'equal' && val.startsWith('*')) return { op: { type: 'endsWith', value: val.slice(1) } };
  const type = (Object.keys(OP_TO_XML) as FilterOp[]).find((k) => OP_TO_XML[k] === operator);
  return type ? { op: { type, value: val } } : null;
}

// ---------- reading ----------

/** Reads the tables of every sheet into `wb`. Returns warnings about tables that couldn't be kept. */
export async function readTables(data: Uint8Array, wb: Workbook): Promise<string[]> {
  const zip = await JSZip.loadAsync(data);
  if (!Object.keys(zip.files).some((f) => /^xl\/tables\//i.test(f))) return [];
  const warnings = new Set<string>();
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    const relsXml = path ? await zip.file(relsPathFor(path))?.async('string') : undefined;
    if (!path || !relsXml) continue;
    for (const rel of parseRels(relsXml).filter((r) => r.type === REL_TABLE || r.type.endsWith('/table'))) {
      const xml = await zip.file(resolvePath(path, rel.target))?.async('string');
      if (!xml) continue;
      const t = tableFromXml(xml, sheet, wb);
      if (typeof t === 'string') warnings.add(t);
      else if (t) sheet.tables.push(t);
    }
    // Rows the file hides because of a table's filter are filtered rows, not hidden ones
    for (const t of sheet.tables) {
      if (!Object.keys(t.filter).length) continue;
      for (let r = t.range.r1 + 1; r <= t.range.r2; r++)
        if (sheet.hiddenRows.delete(r)) sheet.filterHidden.add(r);
    }
  }
  // Excel writes [@Col] as Table1[[#This Row],[Col]], and Table1 as Table1[]
  if (wb.hasTables)
    for (const sheet of wb.sheets)
      for (const [key, cell] of sheet.cells) {
        if (cell.f === undefined || !cell.f.includes('[')) continue;
        const f = structuredFromFile(cell.f, sheet, Math.floor(key / 16384), key % 16384);
        if (f !== cell.f) sheet.cells.set(key, { ...cell, f });
      }
  return [...warnings];
}

function tableFromXml(xml: string, sheet: Sheet, wb: Workbook): TableDef | string | null {
  const head = /<table\b([^>]*)>/.exec(xml);
  if (!head) return null;
  const a = attrs(head[1]);
  const full = parseRange(a.ref ?? '');
  if (!full) return null;
  if (a.headerRowCount === '0') return 'A table without a header row was kept as ordinary cells.';
  const totals = parseInt(a.totalsRowCount ?? '0', 10) || 0;
  const range: Range = { ...full, r2: Math.max(full.r1 + 1, full.r2 - totals) };
  const columns = [...xml.matchAll(/<tableColumn\b([^>]*?)\/?>/g)].map((m) => fromXmlName(attrs(m[1]).name ?? ''));
  const name = a.displayName || a.name || `Table${sheet.tables.length + 1}`;
  if (wb.table(name)) return null;
  const style = /<tableStyleInfo\b([^>]*)\/?>/.exec(xml);
  const st = style ? attrs(style[1]) : {};
  const sort = [...(/<sortState\b[\s\S]*?<\/sortState>/.exec(xml)?.[0] ?? '').matchAll(/<sortCondition\b([^>]*)\/?>/g)]
    .map((m) => attrs(m[1]))
    .map((s) => ({ col: parseRange(s.ref ?? '')?.c1 ?? -1, desc: s.descending === '1' || s.descending === 'true' }))
    .filter((k) => k.col >= range.c1 && k.col <= range.c2);
  const filter: Record<number, FilterCond> = {};
  const af = /<autoFilter\b[^>]*>([\s\S]*?)<\/autoFilter>/.exec(xml)?.[1] ?? '';
  for (const m of af.matchAll(/<filterColumn\b([^>]*)>([\s\S]*?)<\/filterColumn>/g)) {
    const col = range.c1 + (parseInt(attrs(m[1]).colId ?? '-1', 10) || 0);
    const cond = filterFromXml(m[2]);
    if (cond && col >= range.c1 && col <= range.c2) filter[col] = cond;
  }
  const n = range.c2 - range.c1 + 1;
  return {
    id: 't' + Math.random().toString(36).slice(2, 10),
    name,
    range,
    columns: Array.from({ length: n }, (_, i) => columns[i] || `Column${i + 1}`),
    color: colorForStyle(st.name),
    banded: st.showRowStripes !== '0',
    sort,
    filter,
    fileStyle: st.name || undefined,
  };
}

// ---------- writing ----------

function tableXml(t: TableDef, id: number): string {
  const ref = formatRange(t.range);
  const top = { ...t.range, r1: t.range.r1 + 1 };
  const filters = Object.entries(t.filter)
    .map(([k, cond]) => filterXml(+k - t.range.c1, cond))
    .join('');
  const sort = t.sort.length
    ? `<sortState ref="${formatRange(top)}">${t.sort.map((k) => `<sortCondition${k.desc ? ' descending="1"' : ''} ref="${formatRange({ ...top, c1: k.col, c2: k.col })}"/>`).join('')}</sortState>`
    : '';
  const cols = t.columns.map((c, i) => `<tableColumn id="${i + 1}" name="${esc(toXmlName(c))}"/>`).join('');
  const style = t.fileStyle ?? STYLE_FOR[TABLE_COLOR_IDS.includes(t.color) ? t.color : 'blue'];
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${esc(t.name)}" displayName="${esc(t.name)}" ref="${ref}" totalsRowShown="0">` +
    `<autoFilter ref="${ref}">${filters}</autoFilter>${sort}` +
    `<tableColumns count="${t.columns.length}">${cols}</tableColumns>` +
    `<tableStyleInfo${style ? ` name="${esc(style)}"` : ''} showFirstColumn="0" showLastColumn="0" showRowStripes="${t.banded ? 1 : 0}" showColumnStripes="0"/>` +
    `</table>`
  );
}

/** Adds the tables to a package written by ExcelJS. */
export async function writeTables(buf: Uint8Array, wb: Workbook): Promise<Uint8Array> {
  if (!wb.hasTables) return buf;
  const zip = await JSZip.loadAsync(buf);
  const paths = await sheetPaths(zip);
  let ct = (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
  let n = Object.keys(zip.files).filter((f) => /^xl\/tables\/table\d+\.xml$/.test(f)).length;
  for (const sheet of wb.sheets) {
    if (!sheet.tables.length) continue;
    const path = paths.get(sheet.name);
    let xml = path ? await zip.file(path)?.async('string') : undefined;
    if (!path || !xml) continue;
    const relsPath = relsPathFor(path);
    let rels = (await zip.file(relsPath)?.async('string')) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    const used = new Set(parseRels(rels).map((r) => r.id));
    const ids: string[] = [];
    for (const t of sheet.tables) {
      n++;
      let k = 1;
      while (used.has(`rId${k}`)) k++;
      used.add(`rId${k}`);
      ids.push(`rId${k}`);
      rels = rels.replace('</Relationships>', `<Relationship Id="rId${k}" Type="${REL_TABLE}" Target="../tables/table${n}.xml"/></Relationships>`);
      zip.file(`xl/tables/table${n}.xml`, tableXml(t, n));
      ct = ct.replace('</Types>', `<Override PartName="/xl/tables/table${n}.xml" ContentType="${CT_TABLE}"/></Types>`);
    }
    zip.file(relsPath, rels);
    xml = xml.replace(/<tableParts\b[\s\S]*?(<\/tableParts>|\/>)/, '');
    const parts = `<tableParts count="${ids.length}">${ids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>`;
    const at = topLevelSlot(xml, ['extLst']);
    xml = xml.slice(0, at) + parts + xml.slice(at);
    if (!/xmlns:r=/.test(xml.slice(0, 800))) xml = xml.replace('<worksheet ', `<worksheet xmlns:r="${NS_REL}" `);
    zip.file(path, xml);
  }
  zip.file('[Content_Types].xml', ct);
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}
