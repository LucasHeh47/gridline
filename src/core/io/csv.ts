// CSV / TSV import and export.
import { cellKey } from '../address';
import type { Doc } from '../doc';
import { numberToText } from '../formula/coerce';
import { parseInput } from '../input';
import { formatValue, isDateFormat } from '../numfmt';
import { ErrorValue } from '../values';
import { Sheet, Workbook } from '../workbook';

/** Decodes file bytes: UTF-8 (with or without BOM), UTF-16 with BOM, else Windows-1252. */
export function decodeText(bytes: Uint8Array): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return new TextDecoder('utf-8').decode(bytes.subarray(3));
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

/** Picks the most likely delimiter by checking which gives a consistent column count. */
export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 20000);
  const candidates = [',', ';', '\t', '|'];
  let best = ',';
  let bestScore = -1;
  for (const d of candidates) {
    const rows = parseCSV(sample, d).slice(0, 20);
    if (rows.length > 1) rows.pop(); // last row may be cut off
    const counts = rows.map((r) => r.length);
    if (!counts.length || counts[0] < 2) continue;
    const same = counts.filter((c) => c === counts[0]).length;
    const score = same * 10 + counts[0];
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

/** RFC 4180 parser (quoted fields, escaped quotes, newlines inside quotes). */
export function parseCSV(text: string, delimiter = ','): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let i = 0;
  const n = text.length;
  while (i < n) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === '') {
      quoted = true;
      i++;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      if (ch === '\r' && text[i + 1] === '\n') i++;
      i++;
      continue;
    }
    field += ch;
    i++;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function sheetNameFrom(fileName: string): string {
  const base = fileName.replace(/\.[^.]+$/, '').replace(/[\\/?*[\]:]/g, '_').trim();
  return (base || 'Sheet1').slice(0, 31);
}

export function csvToWorkbook(text: string, fileName: string, delimiter?: string): Workbook {
  const wb = new Workbook(true);
  const sheet = new Sheet(sheetNameFrom(fileName));
  wb.sheets.push(sheet);
  wb.activeSheetId = sheet.id;
  const d = delimiter ?? detectDelimiter(text);
  const rows = parseCSV(text, d);
  const widths: number[] = [];
  const fmtIds = new Map<string, number>();
  rows.forEach((fields, r) => {
    fields.forEach((raw, c) => {
      if (raw === '') return;
      const p = parseInput(raw);
      const cell: { v?: string | number | boolean | ErrorValue | null; f?: string; s?: number } = {};
      if (p.formula !== undefined) {
        // CSV files from other apps sometimes contain "=..." text; treat as formulas like Excel does.
        cell.f = p.formula;
      } else cell.v = p.value;
      if (p.numFmt) {
        let id = fmtIds.get(p.numFmt);
        if (id === undefined) fmtIds.set(p.numFmt, (id = wb.styles.intern({ numFmt: p.numFmt })));
        cell.s = id;
      }
      sheet.cells.set(cellKey(r, c), cell);
      if (r < 300) widths[c] = Math.max(widths[c] ?? 0, Math.min(raw.length, 60));
    });
  });
  // Size columns to their content so data is readable right away
  widths.forEach((len, c) => {
    const px = Math.round(len * 7.2 + 16);
    if (px > sheet.defaultColWidth) sheet.colWidths.set(c, Math.min(px, 420));
  });
  sheet.invalidateBounds();
  return wb;
}

function csvField(s: string, delimiter: string): string {
  if (s.includes(delimiter) || /["\r\n]/.test(s) || /^\s|\s$/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

/**
 * Exports a sheet as CSV. Numbers keep full precision (unlike Excel, which writes the rounded
 * display text); dates and times are written as formatted text so other apps can read them.
 */
export function sheetToCSV(doc: Doc, sheet: Sheet, delimiter = ','): string {
  const b = doc.engine.usedBounds(sheet);
  const lines: string[] = [];
  for (let r = 0; r <= b.maxRow; r++) {
    const fields: string[] = [];
    for (let c = 0; c <= b.maxCol; c++) {
      const v = doc.engine.getValue(sheet, r, c);
      let s: string;
      if (v === null) s = '';
      else if (typeof v === 'number') {
        const fmt = doc.numFmtAt(sheet, r, c);
        s = isDateFormat(fmt) ? formatValue(v, fmt).text : numberToText(v);
      } else if (typeof v === 'boolean') s = v ? 'TRUE' : 'FALSE';
      else if (v instanceof ErrorValue) s = v.code;
      else s = v;
      fields.push(csvField(s, delimiter));
    }
    while (fields.length && fields[fields.length - 1] === '') fields.pop();
    lines.push(fields.join(delimiter));
  }
  return lines.join('\r\n') + '\r\n';
}

export function encodeCSV(text: string): Uint8Array {
  // A UTF-8 byte order mark makes Excel and other tools read accented characters correctly.
  const body = new TextEncoder().encode(text);
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf]);
  out.set(body, 3);
  return out;
}
