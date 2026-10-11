// Data validation in .xlsx files (<dataValidations> in each sheet, plus the x14 extension Excel uses
// for lists that point at other sheets).
import JSZip from 'jszip';
import { formatRange, parseRange, type Range } from '../address';
import { newValidationId, type DataValidation, type ValidationOp, type ValidationType } from '../validation';
import type { Workbook } from '../workbook';
import { esc, sheetPaths, topLevelSlot, unesc } from './xlsxCharts';

const TYPES: ValidationType[] = ['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'];
const OPS: ValidationOp[] = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];

function attrs(s: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of s.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = unesc(m[2]);
  return out;
}

const truthy = (v: string | undefined) => v === '1' || v === 'true';

function ranges(sqref: string): Range[] {
  return sqref
    .split(/\s+/)
    .map((p) => parseRange(p))
    .filter((r): r is Range => !!r);
}

function fromXml(a: Record<string, string>, f1: string | undefined, f2: string | undefined, sqref: string): DataValidation | null {
  const type = (a.type ?? 'none') as ValidationType | 'none';
  const rs = ranges(sqref);
  if (!rs.length) return null;
  const known = TYPES.includes(type as ValidationType);
  // "Any value" rules only carry an input message
  if (!known && !(truthy(a.showInputMessage) && (a.prompt || a.promptTitle))) return null;
  return {
    id: newValidationId(),
    ranges: rs,
    type: known ? (type as ValidationType) : 'any',
    operator: OPS.includes(a.operator as ValidationOp) ? (a.operator as ValidationOp) : 'between',
    formula1: (f1 ?? '').trim(),
    formula2: f2?.trim() || undefined,
    allowBlank: truthy(a.allowBlank),
    // In the file, showDropDown="1" means the arrow is hidden
    dropdown: !truthy(a.showDropDown),
    showError: truthy(a.showErrorMessage),
    errorStyle: a.errorStyle === 'warning' || a.errorStyle === 'information' ? a.errorStyle : 'stop',
    errorTitle: a.errorTitle || undefined,
    error: a.error || undefined,
    showPrompt: truthy(a.showInputMessage),
    promptTitle: a.promptTitle || undefined,
    prompt: a.prompt || undefined,
  };
}

/** Reads every sheet's validation rules into `wb`. */
export async function readValidations(data: Uint8Array, wb: Workbook): Promise<number> {
  const zip = await JSZip.loadAsync(data);
  const paths = await sheetPaths(zip);
  let count = 0;
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    const xml = path ? await zip.file(path)?.async('string') : undefined;
    if (!xml) continue;
    const list: DataValidation[] = [];
    const main = /<dataValidations\b[^>]*>([\s\S]*?)<\/dataValidations>/.exec(xml)?.[1] ?? '';
    for (const m of main.matchAll(/<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g)) {
      const a = attrs(m[1]);
      const body = m[2] ?? '';
      const f1 = /<formula1>([\s\S]*?)<\/formula1>/.exec(body)?.[1];
      const f2 = /<formula2>([\s\S]*?)<\/formula2>/.exec(body)?.[1];
      const dv = fromXml(a, f1 !== undefined ? unesc(f1) : undefined, f2 !== undefined ? unesc(f2) : undefined, a.sqref ?? '');
      if (dv) list.push(dv);
    }
    // Excel 2010+ extension (used for lists that refer to other sheets)
    const ext = /<x14:dataValidations\b[^>]*>([\s\S]*?)<\/x14:dataValidations>/.exec(xml)?.[1] ?? '';
    for (const m of ext.matchAll(/<x14:dataValidation\b([^>]*?)>([\s\S]*?)<\/x14:dataValidation>/g)) {
      const a = attrs(m[1]);
      const body = m[2];
      const f1 = /<x14:formula1>\s*<xm:f>([\s\S]*?)<\/xm:f>/.exec(body)?.[1];
      const f2 = /<x14:formula2>\s*<xm:f>([\s\S]*?)<\/xm:f>/.exec(body)?.[1];
      const sqref = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(body)?.[1] ?? '';
      const dv = fromXml(a, f1 !== undefined ? unesc(f1) : undefined, f2 !== undefined ? unesc(f2) : undefined, sqref);
      if (dv) list.push(dv);
    }
    sheet.validations = list;
    count += list.length;
  }
  return count;
}

function toXml(dv: DataValidation): string {
  const a: string[] = [];
  if (dv.type !== 'any') a.push(`type="${dv.type}"`);
  if (dv.errorStyle !== 'stop') a.push(`errorStyle="${dv.errorStyle}"`);
  if (dv.operator !== 'between' && dv.type !== 'list' && dv.type !== 'custom' && dv.type !== 'any') a.push(`operator="${dv.operator}"`);
  if (dv.allowBlank) a.push('allowBlank="1"');
  if (dv.type === 'list' && !dv.dropdown) a.push('showDropDown="1"');
  if (dv.showPrompt) a.push('showInputMessage="1"');
  if (dv.showError) a.push('showErrorMessage="1"');
  if (dv.errorTitle) a.push(`errorTitle="${esc(dv.errorTitle)}"`);
  if (dv.error) a.push(`error="${esc(dv.error)}"`);
  if (dv.promptTitle) a.push(`promptTitle="${esc(dv.promptTitle)}"`);
  if (dv.prompt) a.push(`prompt="${esc(dv.prompt)}"`);
  a.push(`sqref="${dv.ranges.map(formatRange).join(' ')}"`);
  const f1 = dv.formula1.replace(/^=/, '');
  const f2 = dv.formula2?.replace(/^=/, '');
  const body = (dv.type !== 'any' && f1 ? `<formula1>${esc(f1)}</formula1>` : '') + (f2 && (dv.operator === 'between' || dv.operator === 'notBetween') ? `<formula2>${esc(f2)}</formula2>` : '');
  return body ? `<dataValidation ${a.join(' ')}>${body}</dataValidation>` : `<dataValidation ${a.join(' ')}/>`;
}

/** Elements that come after <dataValidations> in a worksheet, in schema order. */
const AFTER = ['hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/** Writes validation rules into a package written by ExcelJS. */
export async function writeValidations(buf: Uint8Array, wb: Workbook): Promise<Uint8Array> {
  if (!wb.sheets.some((s) => s.validations.length)) return buf;
  const zip = await JSZip.loadAsync(buf);
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    if (!sheet.validations.length) continue;
    const path = paths.get(sheet.name);
    let xml = path ? await zip.file(path)?.async('string') : undefined;
    if (!path || !xml) continue;
    xml = xml.replace(/<dataValidations\b[\s\S]*?<\/dataValidations>/, '');
    const block = `<dataValidations count="${sheet.validations.length}">${sheet.validations.map(toXml).join('')}</dataValidations>`;
    const at = topLevelSlot(xml, AFTER);
    xml = xml.slice(0, at) + block + xml.slice(at);
    zip.file(path, xml);
  }
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}
