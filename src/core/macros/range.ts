// The VBA Range object (and the Interior, Font and Borders objects hanging off it).
import { colToName, formatCell, isFullColumn, isFullRow, MAX_COLS, MAX_ROWS, nameToCol, quoteSheetName, rangeContains, type Range } from '../address';
import type { Transaction } from '../history';
import { applyStyle, clearRanges, forEachCellIn, mergeCells, setCellInput, setLink, setNote, type PatchFn } from '../ops/cells';
import { copyRange, pasteClip, type PasteMode } from '../ops/clipboard';
import { applyFilter, currentRegion, detectHeader, displayText, removeDuplicates, sortRange, type SortKey } from '../ops/data';
import { autofill, type FillMode } from '../ops/fill';
import { restructure, setColumnWidth, setHidden, setRowHeight } from '../ops/structure';
import { parseHex, toHex, type BorderSide, type BorderStyle, type HAlign, type Style, type StylePatch, type VAlign } from '../styles';
import type { Cell, FilterCond, FilterOp, Sheet } from '../workbook';
import type { MacroSession } from './host';
import { MISSING } from './vba/builtins';
import { EMPTY, toBool, toNum, toStr, VbArray, VbError, type Value, type VbObject } from './vba/values';

/** An optional argument by position or name, or undefined when it was left out. */
export function opt(args: Value[], named: Record<string, Value> | undefined, i: number, name: string): Value | undefined {
  const v = named?.[name] ?? args[i];
  return v === (MISSING as unknown) || v === undefined ? undefined : v;
}

const cap = (s: string) => s.charAt(0).toUpperCase() + s.slice(1);

export function unsupported(obj: string, member: string): VbError {
  return new VbError(438, `Gridline doesn't support ${obj}.${cap(member)} yet.`);
}

// ---------- colors ----------

/** VBA colors are BGR numbers (RGB(255, 0, 0) = 255). */
export function vbToHex(v: Value): string {
  const n = Math.round(toNum(v));
  return toHex(n & 255, (n >> 8) & 255, (n >> 16) & 255);
}

export function hexToVb(hex: string): number {
  const [r, g, b] = parseHex(hex);
  return r + g * 256 + b * 65536;
}

/** Excel's default 56-color palette for ColorIndex. */
const PALETTE = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080',
  'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF',
  '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC',
  '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

function paletteColor(v: Value): string | null {
  const n = Math.round(toNum(v));
  if (n === -4142 || n === -4105 || n === 0) return null; // xlNone / xlAutomatic
  const hex = PALETTE[n - 1];
  if (!hex) throw new VbError(9, `ColorIndex must be between 1 and 56 (got ${n}).`);
  return '#' + hex.toLowerCase();
}

function paletteIndex(hex: string | undefined): number {
  if (!hex) return -4142;
  const [r, g, b] = parseHex(hex);
  let best = 1;
  let bestD = Infinity;
  PALETTE.forEach((p, i) => {
    const [pr, pg, pb] = parseHex(p);
    const d = (pr - r) ** 2 + (pg - g) ** 2 + (pb - b) ** 2;
    if (d < bestD) {
      bestD = d;
      best = i + 1;
    }
  });
  return best;
}

// ---------- alignment & borders ----------

const H_ALIGN: [number, HAlign][] = [
  [-4131, 'left'],
  [-4108, 'center'],
  [-4152, 'right'],
  [1, 'general'],
  [-4130, 'justify'],
  [5, 'fill'],
  [7, 'centerContinuous'],
  [-4117, 'distributed'],
];
const V_ALIGN: [number, VAlign][] = [
  [-4160, 'top'],
  [-4108, 'middle'],
  [-4107, 'bottom'],
];

type Side = 'borderTop' | 'borderBottom' | 'borderLeft' | 'borderRight';
const ALL_SIDES: Side[] = ['borderTop', 'borderBottom', 'borderLeft', 'borderRight'];

function toBorderStyle(line: number, weight: number): BorderStyle | null {
  if (line === -4142) return null;
  const heavy = weight === -4138 || weight === 4;
  switch (line) {
    case -4115:
      return heavy ? 'mediumDashed' : 'dashed';
    case -4118:
      return 'dotted';
    case -4119:
      return 'double';
    case 4:
      return heavy ? 'mediumDashDot' : 'dashDot';
    case 5:
      return heavy ? 'mediumDashDotDot' : 'dashDotDot';
    case 13:
      return 'slantDashDot';
    default:
      return weight === 1 ? 'hair' : weight === -4138 ? 'medium' : weight === 4 ? 'thick' : 'thin';
  }
}

function fromBorderStyle(st: BorderStyle | undefined): [number, number] {
  switch (st) {
    case undefined:
      return [-4142, 2];
    case 'hair':
      return [1, 1];
    case 'medium':
      return [1, -4138];
    case 'thick':
      return [1, 4];
    case 'dashed':
      return [-4115, 2];
    case 'mediumDashed':
      return [-4115, -4138];
    case 'dotted':
      return [-4118, 2];
    case 'double':
      return [-4119, 4];
    case 'dashDot':
      return [4, 2];
    case 'mediumDashDot':
      return [4, -4138];
    case 'dashDotDot':
      return [5, 2];
    case 'mediumDashDotDot':
      return [5, -4138];
    case 'slantDashDot':
      return [13, -4138];
    default:
      return [1, 2];
  }
}

/** Which border sides of cell (row, col) an Excel border index refers to. 0 = every side of every cell. */
function sidesFor(idx: number, row: number, col: number, r: Range): Side[] {
  switch (idx) {
    case 7:
      return col === r.c1 ? ['borderLeft'] : [];
    case 8:
      return row === r.r1 ? ['borderTop'] : [];
    case 9:
      return row === r.r2 ? ['borderBottom'] : [];
    case 10:
      return col === r.c2 ? ['borderRight'] : [];
    case 11:
      return col < r.c2 ? ['borderRight'] : [];
    case 12:
      return row < r.r2 ? ['borderBottom'] : [];
    case 0:
      return ALL_SIDES;
    default:
      return []; // diagonals aren't drawn by Gridline
  }
}

// ---------- R1C1 formulas ----------

/** Converts an R1C1 formula (as written by Excel's macro recorder) to A1 style for the cell at row, col. */
export function r1c1ToA1(formula: string, row: number, col: number): string {
  const part = (p: string | undefined, here: number): { abs: boolean; n: number } => {
    if (!p) return { abs: false, n: here };
    if (p.startsWith('[')) return { abs: false, n: here + parseInt(p.slice(1, -1), 10) };
    return { abs: true, n: parseInt(p, 10) - 1 };
  };
  return formula.replace(/("(?:[^"]|"")*")|(?<![A-Za-z0-9_.])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(.])/g, (m, str, rp, cp) => {
    if (str) return str;
    const r = part(rp, row);
    const c = part(cp, col);
    if (r.n < 0 || c.n < 0 || r.n >= MAX_ROWS || c.n >= MAX_COLS) return '#REF!';
    return `${c.abs ? '$' : ''}${colToName(c.n)}${r.abs ? '$' : ''}${r.n + 1}`;
  });
}

// ---------- Range ----------

export type RangeMode = 'cells' | 'rows' | 'cols';

export class RangeObj implements VbObject {
  readonly vbType = 'Range';

  constructor(
    readonly s: MacroSession,
    readonly sheet: Sheet,
    readonly areas: Range[],
    readonly mode: RangeMode = 'cells',
  ) {}

  get r(): Range {
    return this.areas[0];
  }

  private get doc() {
    return this.s.doc;
  }

  sub(r: Range, mode: RangeMode = 'cells'): RangeObj {
    return new RangeObj(this.s, this.sheet, [r], mode);
  }

  private at(row: number, col: number): RangeObj {
    if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) throw new VbError(1004, 'That cell is outside the sheet.');
    return this.sub({ r1: row, c1: col, r2: row, c2: col });
  }

  private style(row = this.r.r1, col = this.r.c1): Style {
    return this.doc.wb.styles.get(this.sheet.styleIdAt(row, col));
  }

  private patch(p: PatchFn): void {
    this.s.edit((tx) => applyStyle(tx, this.doc.wb, this.sheet, this.areas, p));
  }

  /** Cells covered, with whole rows/columns limited to the used part of the sheet. */
  private used(r: Range): Range {
    return this.s.clamp(this.sheet, r);
  }

  get(name: string, args: Value[], named?: Record<string, Value>): Value | Promise<Value> {
    this.s.checkSheet(this.sheet);
    const o = (i: number, n: string) => opt(args, named, i, n);
    const r = this.r;
    switch (name) {
      case 'value':
      case 'value2':
        return this.readValue(name === 'value2');
      case 'formula':
      case 'formulalocal':
      case 'formula2':
      case 'formulaarray':
      case 'formular1c1':
        return this.readFormula();
      case 'text':
        this.s.settle();
        return displayText(this.doc, this.sheet, r.r1, r.c1);
      case 'numberformat':
      case 'numberformatlocal':
        return this.style().numFmt ?? 'General';
      case 'address':
      case 'addresslocal':
        return this.address(o(0, 'rowabsolute'), o(1, 'columnabsolute'), o(2, 'referencestyle'), o(3, 'external'));
      case 'row':
        return r.r1 + 1;
      case 'column':
        return r.c1 + 1;
      case 'rows':
        return args.length ? this.lineItem('rows', args[0]) : new RangeObj(this.s, this.sheet, this.areas, 'rows');
      case 'columns':
        return args.length ? this.lineItem('cols', args[0]) : new RangeObj(this.s, this.sheet, this.areas, 'cols');
      case 'count':
      case 'countlarge':
        return this.areas.reduce((n, a) => n + (this.mode === 'rows' ? a.r2 - a.r1 + 1 : this.mode === 'cols' ? a.c2 - a.c1 + 1 : (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1)), 0);
      case 'cells':
        return args.length ? this.cellItem(args) : new RangeObj(this.s, this.sheet, this.areas, 'cells');
      case 'item':
        return this.item(args);
      case 'range': {
        const inner = this.s.rangeArg(args[0], this.sheet, o(1, 'cell2'));
        return new RangeObj(this.s, this.sheet, inner.areas.map((a) => ({ r1: a.r1 + r.r1, c1: a.c1 + r.c1, r2: a.r2 + r.r1, c2: a.c2 + r.c1 })));
      }
      case 'offset': {
        const dr = Math.round(toNum(o(0, 'rowoffset') ?? 0));
        const dc = Math.round(toNum(o(1, 'columnoffset') ?? 0));
        const moved = this.areas.map((a) => ({ r1: a.r1 + dr, c1: a.c1 + dc, r2: a.r2 + dr, c2: a.c2 + dc }));
        if (moved.some((a) => a.r1 < 0 || a.c1 < 0 || a.r2 >= MAX_ROWS || a.c2 >= MAX_COLS)) throw new VbError(1004, 'Offset goes past the edge of the sheet.');
        return new RangeObj(this.s, this.sheet, moved);
      }
      case 'resize': {
        const rows = o(0, 'rowsize');
        const cols = o(1, 'columnsize');
        const h = rows === undefined ? r.r2 - r.r1 + 1 : Math.round(toNum(rows));
        const w = cols === undefined ? r.c2 - r.c1 + 1 : Math.round(toNum(cols));
        if (h < 1 || w < 1) throw new VbError(1004, 'Resize needs at least 1 row and 1 column.');
        if (r.r1 + h > MAX_ROWS || r.c1 + w > MAX_COLS) throw new VbError(1004, 'Resize goes past the edge of the sheet.');
        return this.sub({ r1: r.r1, c1: r.c1, r2: r.r1 + h - 1, c2: r.c1 + w - 1 });
      }
      case 'end':
        return this.end(toNum(o(0, 'direction') ?? -4121));
      case 'currentregion':
        this.s.settle();
        return this.sub(currentRegion(this.doc, this.sheet, r.r1, r.c1));
      case 'entirerow':
        return new RangeObj(this.s, this.sheet, this.areas.map((a) => ({ r1: a.r1, c1: 0, r2: a.r2, c2: MAX_COLS - 1 })));
      case 'entirecolumn':
        return new RangeObj(this.s, this.sheet, this.areas.map((a) => ({ r1: 0, c1: a.c1, r2: MAX_ROWS - 1, c2: a.c2 })));
      case 'mergearea':
        return this.sub(this.sheet.mergeAt(r.r1, r.c1) ?? { r1: r.r1, c1: r.c1, r2: r.r1, c2: r.c1 });
      case 'interior':
        return new InteriorObj(this);
      case 'font':
        return new FontObj(this);
      case 'borders':
        return args.length ? new BorderObj(this, toNum(args[0])) : new BordersObj(this);
      case 'borderaround': {
        const line = toNum(o(0, 'linestyle') ?? 1);
        const weight = toNum(o(1, 'weight') ?? 2);
        const ci = o(2, 'colorindex');
        const color = o(3, 'color');
        const hex = color !== undefined ? vbToHex(color) : ci !== undefined ? paletteColor(ci) ?? undefined : undefined;
        const st = toBorderStyle(line, weight);
        for (const idx of [7, 8, 9, 10]) this.setBorder(idx, () => (st ? { style: st, color: hex } : undefined));
        return true;
      }
      case 'horizontalalignment':
        return H_ALIGN.find(([, a]) => a === (this.style().hAlign ?? 'general'))?.[0] ?? 1;
      case 'verticalalignment':
        return V_ALIGN.find(([, a]) => a === (this.style().vAlign ?? 'bottom'))?.[0] ?? -4107;
      case 'wraptext':
        return !!this.style().wrap;
      case 'indentlevel':
        return this.style().indent ?? 0;
      case 'orientation':
        return this.style().rotation ?? 0;
      case 'columnwidth':
        return Math.round(((this.sheet.colWidth(r.c1) - 5) / 7) * 100) / 100;
      case 'rowheight':
        return Math.round(this.sheet.rowHeight(r.r1) * 0.75 * 100) / 100;
      case 'width':
      case 'height':
      case 'left':
      case 'top':
        return this.metrics()[name];
      case 'hidden':
        return isFullColumn(r) ? this.sheet.isColHidden(r.c1) : this.sheet.isRowHidden(r.r1);
      case 'mergecells':
        return !!this.sheet.mergeAt(r.r1, r.c1);
      case 'hasformula': {
        let all = true;
        let any = false;
        forEachCellIn(this.sheet, this.used(r), (_r, _c, cell) => {
          if (cell.f !== undefined) any = true;
        });
        for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 1000); row++)
          for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 100); c++) if (this.sheet.getCell(row, c)?.f === undefined) all = false;
        return all ? true : any ? EMPTY : false;
      }
      case 'worksheet':
      case 'parent':
        return this.s.sheetObj(this.sheet);
      case 'areas':
        return args.length ? this.areaItem(args[0]) : new AreasObj(this);
      case 'comment':
        return this.sheet.getCell(r.r1, r.c1)?.note ? new CommentObj(this) : null;
      case 'notetext':
        if (args.length) {
          this.s.edit((tx) => setNote(tx, this.sheet, r.r1, r.c1, toStr(args[0])));
          return EMPTY;
        }
        return this.sheet.getCell(r.r1, r.c1)?.note ?? '';
      case 'addcomment':
      case 'addcommentthreaded':
        this.s.edit((tx) => setNote(tx, this.sheet, r.r1, r.c1, toStr(o(0, 'text') ?? '')));
        return new CommentObj(this);
      case 'clearcomments':
      case 'clearnotes':
        this.s.edit((tx) => clearRanges(tx, this.sheet, this.areas, 'notes'));
        return EMPTY;
      case 'clear':
        this.s.edit((tx) => clearRanges(tx, this.sheet, this.areas, 'all'));
        return EMPTY;
      case 'clearcontents':
        this.s.edit((tx) => clearRanges(tx, this.sheet, this.areas, 'contents'));
        return EMPTY;
      case 'clearformats':
        this.s.edit((tx) => clearRanges(tx, this.sheet, this.areas, 'formats'));
        return EMPTY;
      case 'clearhyperlinks':
        this.s.edit((tx) => clearRanges(tx, this.sheet, this.areas, 'links'));
        return EMPTY;
      case 'delete':
        this.deleteCells(o(0, 'shift'));
        return true;
      case 'insert':
        this.insertCells(o(0, 'shift'));
        return true;
      case 'copy':
        this.copy(o(0, 'destination'), false);
        return true;
      case 'cut':
        this.copy(o(0, 'destination'), true);
        return true;
      case 'pastespecial':
        this.s.paste(this.sheet, r, pasteMode(o(0, 'paste'), o(3, 'transpose')));
        return true;
      case 'select':
        this.s.ui.select(this.sheet, this.areas, { row: r.r1, col: r.c1 });
        return true;
      case 'activate': {
        const cur = this.s.ui.selection();
        if (cur.sheet === this.sheet && cur.ranges.some((x) => rangeContains(x, r.r1, r.c1))) this.s.ui.select(this.sheet, cur.ranges, { row: r.r1, col: r.c1 });
        else this.s.ui.select(this.sheet, [{ r1: r.r1, c1: r.c1, r2: r.r1, c2: r.c1 }], { row: r.r1, col: r.c1 });
        return true;
      }
      case 'merge': {
        const across = toBool(o(0, 'across') ?? false);
        this.s.edit((tx) => {
          for (const a of this.areas) mergeCells(tx, this.doc, this.sheet, a, across ? 'across' : 'merge');
        });
        return EMPTY;
      }
      case 'unmerge':
        this.s.edit((tx) => {
          for (const a of this.areas) mergeCells(tx, this.doc, this.sheet, a, 'unmerge');
        });
        return EMPTY;
      case 'autofit':
        this.autofit();
        return true;
      case 'sort':
        this.sort(args, named);
        return true;
      case 'find':
        return this.s.find(this, {
          what: toStr(o(0, 'what') ?? ''),
          after: o(1, 'after'),
          lookIn: toNum(o(2, 'lookin') ?? -4163),
          lookAt: toNum(o(3, 'lookat') ?? 2),
          byColumns: toNum(o(4, 'searchorder') ?? 1) === 2,
          previous: toNum(o(5, 'searchdirection') ?? 1) === 2,
          matchCase: toBool(o(6, 'matchcase') ?? false),
        });
      case 'findnext':
      case 'findprevious':
        return this.s.findAgain(this, o(0, 'after'), name === 'findprevious');
      case 'replace':
        return this.replace(toStr(o(0, 'what') ?? ''), toStr(o(1, 'replacement') ?? ''), toNum(o(2, 'lookat') ?? 2) === 1, toBool(o(4, 'matchcase') ?? false));
      case 'filldown':
      case 'fillup':
      case 'fillright':
      case 'fillleft': {
        const src: Range =
          name === 'filldown' ? { ...r, r2: r.r1 } : name === 'fillup' ? { ...r, r1: r.r2 } : name === 'fillright' ? { ...r, c2: r.c1 } : { ...r, c1: r.c2 };
        this.s.edit((tx) => autofill(tx, this.doc, this.sheet, src, r, 'copy'));
        return true;
      }
      case 'autofill': {
        const dest = this.s.rangeArg(o(0, 'destination'), this.sheet);
        const type = toNum(o(1, 'type') ?? 0);
        const mode: FillMode = type === 1 ? 'copy' : type === 2 ? 'series' : type === 3 ? 'formats' : type === 4 ? 'values' : 'auto';
        this.s.edit((tx) => autofill(tx, this.doc, this.sheet, r, dest.r, mode));
        return true;
      }
      case 'specialcells':
        return this.specialCells(toNum(o(0, 'type') ?? 2));
      case 'removeduplicates': {
        const colsArg = o(0, 'columns') ?? 1;
        const rel = colsArg instanceof VbArray ? colsArg.data.map(toNum) : [toNum(colsArg)];
        const header = toNum(o(1, 'header') ?? 2) === 1;
        const area = this.used(r);
        this.s.edit((tx) => removeDuplicates(tx, this.doc, this.sheet, area, rel.map((c) => area.c1 + c - 1), header));
        return EMPTY;
      }
      case 'autofilter':
        this.autoFilter(o(0, 'field'), o(1, 'criteria1'), o(2, 'operator'));
        return true;
      case 'hyperlinks':
        return new RangeLinksObj(this);
      case 'calculate':
        this.s.settle();
        return EMPTY;
      case 'locked':
      case 'formulahidden':
        return false;
      case 'next':
        return this.at(r.r1, r.c1 + 1);
      case 'previous':
        return this.at(r.r1, Math.max(0, r.c1 - 1));
      case 'show':
        return EMPTY;
      case 'name': {
        const hit = [...this.doc.wb.names.values()].find((n) => n.ref.replace(/\$/g, '').toLowerCase() === `${quoteSheetName(this.sheet.name)}!${this.address(false, false)}`.toLowerCase());
        if (!hit) throw new VbError(1004, 'This range has no name.');
        return { vbType: 'Name', get: (n: string) => (n === 'name' ? hit.name : '=' + hit.ref), set: () => undefined, dflt: () => hit.name } as VbObject;
      }
    }
    throw unsupported('Range', name);
  }

  set(name: string, args: Value[], v: Value): void {
    this.s.checkSheet(this.sheet);
    const r = this.r;
    switch (name) {
      case 'value':
      case 'value2':
        if (args.length) return this.cellItem(args).writeValue(v);
        return this.writeValue(v);
      case 'formula':
      case 'formulalocal':
      case 'formula2':
      case 'formulaarray':
        return this.writeValue(v);
      case 'formular1c1':
        return this.writeR1C1(v);
      case 'numberformat':
      case 'numberformatlocal':
        return this.patch({ numFmt: toStr(v) || 'General' });
      case 'horizontalalignment': {
        const a = H_ALIGN.find(([n]) => n === toNum(v))?.[1];
        if (!a) throw new VbError(1004, 'Unknown alignment.');
        return this.patch({ hAlign: a });
      }
      case 'verticalalignment': {
        const n = toNum(v);
        const a = V_ALIGN.find(([x]) => x === n)?.[1] ?? (n === -4130 || n === -4117 ? 'middle' : undefined);
        if (!a) throw new VbError(1004, 'Unknown alignment.');
        return this.patch({ vAlign: a });
      }
      case 'wraptext':
        return this.patch({ wrap: toBool(v) || null });
      case 'indentlevel':
        return this.patch({ indent: Math.max(0, Math.round(toNum(v))) || null });
      case 'orientation':
        return this.patch({ rotation: Math.round(toNum(v)) || null });
      case 'columnwidth': {
        const px = Math.round(toNum(v) * 7 + 5);
        return this.s.edit((tx) => setColumnWidth(tx, this.sheet, this.indexes('col'), px));
      }
      case 'rowheight': {
        const px = Math.round(toNum(v) / 0.75);
        return this.s.edit((tx) => setRowHeight(tx, this.sheet, this.indexes('row'), px));
      }
      case 'hidden': {
        const hide = toBool(v);
        const axis = isFullColumn(r) || this.mode === 'cols' ? 'col' : 'row';
        return this.s.edit((tx) => setHidden(tx, this.sheet, axis, this.indexes(axis), hide));
      }
      case 'mergecells':
        return this.s.edit((tx) => {
          for (const a of this.areas) mergeCells(tx, this.doc, this.sheet, a, toBool(v) ? 'merge' : 'unmerge');
        });
      case 'name':
        return this.s.defineName(toStr(v), `${quoteSheetName(this.sheet.name)}!${this.address(true, true)}`);
      case 'locked':
      case 'formulahidden':
      case 'addindent':
      case 'shrinktofit':
        return;
    }
    throw unsupported('Range', name);
  }

  dflt(args: Value[]): Value | Promise<Value> {
    return args.length ? this.item(args) : this.readValue(false);
  }

  setDflt(args: Value[], v: Value): void {
    if (args.length) return this.item(args).writeValue(v);
    this.writeValue(v);
  }

  /** For Each: cells (or rows/columns), limited to the used part for whole rows/columns. */
  items(): Value[] {
    const out: Value[] = [];
    for (const a of this.areas) {
      const u = this.used(a);
      if (this.mode === 'rows') for (let row = u.r1; row <= u.r2; row++) out.push(this.sub({ r1: row, c1: a.c1, r2: row, c2: a.c2 }));
      else if (this.mode === 'cols') for (let c = u.c1; c <= u.c2; c++) out.push(this.sub({ r1: a.r1, c1: c, r2: a.r2, c2: c }));
      else {
        if ((u.r2 - u.r1 + 1) * (u.c2 - u.c1 + 1) > 2_000_000) throw new VbError(7, 'That range has too many cells to loop over.');
        for (let row = u.r1; row <= u.r2; row++) for (let c = u.c1; c <= u.c2; c++) out.push(this.at(row, c));
      }
    }
    return out;
  }

  // ---------- items ----------

  private item(args: Value[]): RangeObj {
    if (args.length === 1 && this.mode !== 'cells') return this.lineItem(this.mode, args[0]);
    return this.cellItem(args);
  }

  private cellItem(args: Value[]): RangeObj {
    const r = this.r;
    if (args.length >= 2) {
      const col = typeof args[1] === 'string' && /^[A-Za-z]{1,3}$/.test(args[1]) ? nameToCol(args[1]) + 1 : toNum(args[1]);
      return this.at(r.r1 + Math.round(toNum(args[0])) - 1, r.c1 + Math.round(col) - 1);
    }
    const n = Math.round(toNum(args[0])) - 1;
    const w = r.c2 - r.c1 + 1;
    return this.at(r.r1 + Math.floor(n / w), r.c1 + (((n % w) + w) % w));
  }

  /** Rows(n) / Columns(n), or Columns("B") / Columns("B:D") / Rows("2:5"). */
  private lineItem(mode: 'rows' | 'cols', key: Value): RangeObj {
    const r = this.r;
    if (typeof key === 'string' && !/^\d+$/.test(key.trim())) {
      const parts = key.split(':').map((p) => p.trim().replace(/\$/g, ''));
      if (mode === 'cols') {
        const c1 = nameToCol(parts[0]);
        const c2 = nameToCol(parts[1] ?? parts[0]);
        if (c1 < 0 || c2 < 0) throw new VbError(13, `"${key}" isn't a column.`);
        return this.sub({ r1: r.r1, c1: r.c1 + Math.min(c1, c2), r2: r.r2, c2: r.c1 + Math.max(c1, c2) }, 'cols');
      }
      const r1 = parseInt(parts[0], 10) - 1;
      const r2 = parseInt(parts[1] ?? parts[0], 10) - 1;
      if (!(r1 >= 0) || !(r2 >= 0)) throw new VbError(13, `"${key}" isn't a row.`);
      return this.sub({ r1: r.r1 + Math.min(r1, r2), c1: r.c1, r2: r.r1 + Math.max(r1, r2), c2: r.c2 }, 'rows');
    }
    const n = Math.round(toNum(key)) - 1;
    if (mode === 'rows') {
      if (r.r1 + n < 0 || r.r1 + n >= MAX_ROWS) throw new VbError(1004, 'That row is outside the sheet.');
      return this.sub({ r1: r.r1 + n, c1: r.c1, r2: r.r1 + n, c2: r.c2 }, 'rows');
    }
    if (r.c1 + n < 0 || r.c1 + n >= MAX_COLS) throw new VbError(1004, 'That column is outside the sheet.');
    return this.sub({ r1: r.r1, c1: r.c1 + n, r2: r.r2, c2: r.c1 + n }, 'cols');
  }

  private areaItem(i: Value): RangeObj {
    const a = this.areas[Math.round(toNum(i)) - 1];
    if (!a) throw new VbError(9, `This range has ${this.areas.length} area(s).`);
    return this.sub(a);
  }

  /** Row or column indexes covered (for widths, heights and hiding). */
  private indexes(axis: 'row' | 'col'): number[] {
    const out = new Set<number>();
    for (const a of this.areas) {
      if (axis === 'col') for (let c = a.c1; c <= a.c2; c++) out.add(c);
      else {
        const u = isFullColumn(a) ? this.used(a) : a;
        for (let row = u.r1; row <= u.r2; row++) out.add(row);
      }
    }
    return [...out];
  }

  // ---------- values ----------

  private readValue(value2: boolean): Value {
    this.s.settle();
    const r = this.r;
    if (r.r1 === r.r2 && r.c1 === r.c2) return this.s.cellValue(this.sheet, r.r1, r.c1, value2);
    const u = this.used(r);
    if ((u.r2 - u.r1 + 1) * (u.c2 - u.c1 + 1) > 5_000_000) throw new VbError(7, 'That range is too big to read at once.');
    const arr = new VbArray([
      { lb: 1, ub: u.r2 - u.r1 + 1 },
      { lb: 1, ub: u.c2 - u.c1 + 1 },
    ]);
    let i = 0;
    for (let row = u.r1; row <= u.r2; row++) for (let c = u.c1; c <= u.c2; c++) arr.data[i++] = this.s.cellValue(this.sheet, row, c, value2);
    return arr;
  }

  private readFormula(): Value {
    const r = this.r;
    const one = (row: number, c: number): string => {
      const cell = this.sheet.getCell(row, c);
      if (!cell) return '';
      if (cell.f !== undefined) return '=' + cell.f;
      const v = cell.v;
      return v === undefined || v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
    };
    if (r.r1 === r.r2 && r.c1 === r.c2) return one(r.r1, r.c1);
    const u = this.used(r);
    const arr = new VbArray([
      { lb: 1, ub: u.r2 - u.r1 + 1 },
      { lb: 1, ub: u.c2 - u.c1 + 1 },
    ]);
    let i = 0;
    for (let row = u.r1; row <= u.r2; row++) for (let c = u.c1; c <= u.c2; c++) arr.data[i++] = one(row, c);
    return arr;
  }

  writeValue(v: Value): void {
    this.s.edit((tx) => {
      for (const area of this.areas) {
        const a = this.s.clampWrite(this.sheet, area);
        if (v instanceof VbArray) {
          const get = arrayReader(v);
          for (let row = a.r1; row <= a.r2; row++) for (let c = a.c1; c <= a.c2; c++) this.s.writeCell(tx, this.sheet, row, c, get(row - a.r1, c - a.c1));
          continue;
        }
        // A formula typed into many cells adjusts its references for each cell, like Ctrl+Enter
        if (typeof v === 'string' && v.startsWith('=') && v.length > 1 && (a.r2 > a.r1 || a.c2 > a.c1)) {
          for (let row = a.r1; row <= a.r2; row++)
            for (let c = a.c1; c <= a.c2; c++) {
              if (row === a.r1 && c === a.c1) setCellInput(tx, this.doc, this.sheet, row, c, v);
              else {
                const base = this.sheet.getCell(a.r1, a.c1)?.f;
                setCellInput(tx, this.doc, this.sheet, row, c, base !== undefined ? '=' + this.s.translate(base, row - a.r1, c - a.c1) : v);
              }
            }
          continue;
        }
        for (let row = a.r1; row <= a.r2; row++) for (let c = a.c1; c <= a.c2; c++) this.s.writeCell(tx, this.sheet, row, c, v);
      }
    });
  }

  private writeR1C1(v: Value): void {
    const text = toStr(v);
    if (!text.startsWith('=')) return this.writeValue(v);
    this.s.edit((tx) => {
      for (const area of this.areas) {
        const a = this.s.clampWrite(this.sheet, area);
        for (let row = a.r1; row <= a.r2; row++) for (let c = a.c1; c <= a.c2; c++) setCellInput(tx, this.doc, this.sheet, row, c, r1c1ToA1(text, row, c));
      }
    });
  }

  address(rowAbs?: Value, colAbs?: Value, refStyle?: Value, external?: Value): string {
    const ra = rowAbs === undefined ? true : toBool(rowAbs);
    const ca = colAbs === undefined ? true : toBool(colAbs);
    const r1c1 = refStyle !== undefined && toNum(refStyle) === -4150;
    const cell = (row: number, col: number) => (r1c1 ? `R${row + 1}C${col + 1}` : formatCell(row, col, ra, ca));
    const one = (a: Range): string => {
      if (isFullColumn(a)) return r1c1 ? `C${a.c1 + 1}:C${a.c2 + 1}` : `${ca ? '$' : ''}${colToName(a.c1)}:${ca ? '$' : ''}${colToName(a.c2)}`;
      if (isFullRow(a)) return r1c1 ? `R${a.r1 + 1}:R${a.r2 + 1}` : `${ra ? '$' : ''}${a.r1 + 1}:${ra ? '$' : ''}${a.r2 + 1}`;
      if (a.r1 === a.r2 && a.c1 === a.c2) return cell(a.r1, a.c1);
      return `${cell(a.r1, a.c1)}:${cell(a.r2, a.c2)}`;
    };
    const text = this.areas.map(one).join(',');
    if (external !== undefined && toBool(external)) return `[${this.doc.fileName}]${quoteSheetName(this.sheet.name)}!${text}`;
    return text;
  }

  private metrics(): { width: number; height: number; left: number; top: number } {
    const r = this.r;
    let left = 0;
    let top = 0;
    let width = 0;
    let height = 0;
    for (let c = 0; c < r.c1; c++) if (!this.sheet.isColHidden(c)) left += this.sheet.colWidth(c);
    for (let row = 0; row < Math.min(r.r1, 100000); row++) if (!this.sheet.isRowHidden(row)) top += this.sheet.rowHeight(row);
    for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 1000); c++) if (!this.sheet.isColHidden(c)) width += this.sheet.colWidth(c);
    for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 100000); row++) if (!this.sheet.isRowHidden(row)) height += this.sheet.rowHeight(row);
    const pt = (px: number) => Math.round(px * 0.75 * 100) / 100;
    return { width: pt(width), height: pt(height), left: pt(left), top: pt(top) };
  }

  // ---------- navigation ----------

  private filled(row: number, col: number): boolean {
    const cell = this.sheet.getCell(row, col);
    if (cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''))) return true;
    return this.doc.engine.getValue(this.sheet, row, col) !== null;
  }

  /** Ctrl+Arrow: End(xlUp), End(xlDown), End(xlToLeft), End(xlToRight). */
  private end(dir: number): RangeObj {
    this.s.settle();
    const dr = dir === -4162 ? -1 : dir === -4121 ? 1 : 0;
    const dc = dir === -4159 ? -1 : dir === -4161 ? 1 : 0;
    if (!dr && !dc) throw new VbError(5, 'End needs xlUp, xlDown, xlToLeft or xlToRight.');
    const b = this.doc.engine.usedBounds(this.sheet);
    const maxR = MAX_ROWS - 1;
    const maxC = MAX_COLS - 1;
    let row = this.r.r1;
    let col = this.r.c1;
    const inside = (r: number, c: number) => r >= 0 && c >= 0 && r <= maxR && c <= maxC;
    if (!inside(row + dr, col + dc)) return this.at(row, col);
    if (this.filled(row, col) && this.filled(row + dr, col + dc)) {
      while (inside(row + dr, col + dc) && this.filled(row + dr, col + dc)) {
        row += dr;
        col += dc;
      }
      return this.at(row, col);
    }
    row += dr;
    col += dc;
    // Jump over the empty space beyond the data in one step
    if (dr < 0 && row > b.maxRow) row = Math.max(0, b.maxRow);
    if (dc < 0 && col > b.maxCol) col = Math.max(0, b.maxCol);
    while (inside(row, col) && !this.filled(row, col)) {
      if ((dr > 0 && row > b.maxRow) || (dc > 0 && col > b.maxCol)) return this.at(dr > 0 ? maxR : row, dc > 0 ? maxC : col);
      if (!inside(row + dr, col + dc)) break;
      row += dr;
      col += dc;
    }
    return this.at(row, col);
  }

  // ---------- editing ----------

  private copy(dest: Value | undefined, cut: boolean): void {
    this.s.settle();
    const clip = copyRange(this.doc, this.sheet, this.r, cut);
    if (dest === undefined) {
      this.s.clipboard = clip;
      return;
    }
    const d = this.s.rangeArg(dest, this.sheet);
    this.s.edit((tx) => pasteClip(tx, this.doc, clip, d.sheet, d.r, 'all'));
  }

  private deleteCells(shift: Value | undefined): void {
    this.s.edit((tx) => {
      // Bottom-most / right-most areas first so earlier deletions don't move later ones
      const areas = [...this.areas].sort((a, b) => b.r1 - a.r1 || b.c1 - a.c1);
      for (const a of areas) {
        if (isFullRow(a)) restructure(tx, this.doc.wb, this.sheet, 'row', a.r1, -(a.r2 - a.r1 + 1));
        else if (isFullColumn(a)) restructure(tx, this.doc.wb, this.sheet, 'col', a.c1, -(a.c2 - a.c1 + 1));
        else {
          const up = shift === undefined ? a.c2 - a.c1 >= a.r2 - a.r1 : toNum(shift) === -4162;
          shiftCells(tx, this.sheet, a, up ? 'up' : 'left');
        }
      }
    });
  }

  private insertCells(shift: Value | undefined): void {
    this.s.edit((tx) => {
      for (const a of this.areas) {
        if (isFullRow(a)) restructure(tx, this.doc.wb, this.sheet, 'row', a.r1, a.r2 - a.r1 + 1);
        else if (isFullColumn(a)) restructure(tx, this.doc.wb, this.sheet, 'col', a.c1, a.c2 - a.c1 + 1);
        else {
          const down = shift === undefined ? a.c2 - a.c1 >= a.r2 - a.r1 : toNum(shift) === -4121;
          shiftCells(tx, this.sheet, a, down ? 'down' : 'right');
        }
      }
    });
  }

  private autofit(): void {
    const r = this.r;
    const cols = this.mode === 'cols' || isFullColumn(r);
    const rows = this.mode === 'rows' || isFullRow(r);
    if (!cols && !rows) throw new VbError(1004, 'AutoFit works on columns or rows — use .Columns.AutoFit or .Rows.AutoFit.');
    this.s.settle();
    this.s.edit((tx) => {
      if (cols) {
        for (const c of this.indexes('col')) {
          const w = this.s.measure(this.sheet, 'col', c);
          setColumnWidth(tx, this.sheet, [c], w > 0 ? Math.max(24, Math.min(600, Math.ceil(w))) : undefined);
        }
      } else {
        for (const row of this.indexes('row')) {
          const h = this.s.measure(this.sheet, 'row', row);
          setRowHeight(tx, this.sheet, [row], h <= this.sheet.defaultRowHeight ? undefined : h);
        }
      }
    });
  }

  private sort(args: Value[], named: Record<string, Value> | undefined): void {
    this.s.settle();
    const o = (i: number, n: string) => opt(args, named, i, n);
    let area = this.used(this.r);
    if (area.r1 === area.r2 && area.c1 === area.c2) area = currentRegion(this.doc, this.sheet, area.r1, area.c1);
    const keys: SortKey[] = [];
    const add = (key: Value | undefined, order: Value | undefined) => {
      if (key === undefined) return;
      const k = this.s.rangeArg(key, this.sheet);
      keys.push({ col: k.r.c1, desc: order !== undefined && toNum(order) === 2 });
    };
    add(o(0, 'key1'), o(1, 'order1'));
    add(o(2, 'key2'), o(4, 'order2'));
    add(o(5, 'key3'), o(6, 'order3'));
    if (!keys.length) keys.push({ col: area.c1, desc: false });
    const h = toNum(o(7, 'header') ?? 2);
    const header = h === 1 || (h === 0 && detectHeader(this.doc, this.sheet, area));
    this.s.edit((tx) => sortRange(tx, this.doc, this.sheet, area, keys, header));
  }

  private replace(what: string, replacement: string, whole: boolean, matchCase: boolean): boolean {
    if (!what) return false;
    this.s.settle();
    const re = wildcard(what, whole, matchCase, true);
    this.s.edit((tx) => {
      for (const a of this.areas) {
        const hits: [number, number, Cell][] = [];
        forEachCellIn(this.sheet, this.used(a), (row, c, cell) => hits.push([row, c, cell]));
        for (const [row, c, cell] of hits) {
          const text = cell.f !== undefined ? '=' + cell.f : cell.v === undefined || cell.v === null ? '' : typeof cell.v === 'object' ? cell.v.code : String(cell.v);
          if (!text) continue;
          re.lastIndex = 0;
          if (!re.test(text)) continue;
          re.lastIndex = 0;
          const next = text.replace(re, replacement.replace(/\$/g, '$$$$'));
          if (next !== text) setCellInput(tx, this.doc, this.sheet, row, c, next);
        }
      }
    });
    return true;
  }

  private specialCells(type: number): RangeObj {
    this.s.settle();
    if (type === 11) {
      const b = this.sheet.bounds();
      return this.at(Math.max(0, b.maxRow), Math.max(0, b.maxCol));
    }
    const hits: Range[] = [];
    const push = (row: number, c: number) => {
      const last = hits[hits.length - 1];
      if (last && last.r1 === row && last.c2 === c - 1) last.c2 = c;
      else hits.push({ r1: row, c1: c, r2: row, c2: c });
    };
    for (const a of this.areas) {
      const u = this.used(isFullColumn(a) || isFullRow(a) || (a.r1 === a.r2 && a.c1 === a.c2) ? { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 } : a);
      const area = a.r1 === a.r2 && a.c1 === a.c2 ? u : { r1: Math.max(a.r1, u.r1), c1: Math.max(a.c1, u.c1), r2: Math.min(a.r2, u.r2), c2: Math.min(a.c2, u.c2) };
      for (let row = area.r1; row <= area.r2; row++) {
        if (type === 12 && this.sheet.isRowHidden(row)) continue;
        for (let c = area.c1; c <= area.c2; c++) {
          const cell = this.sheet.getCell(row, c);
          let ok = false;
          if (type === 2) ok = !!cell && cell.f === undefined && cell.v !== undefined && cell.v !== null && cell.v !== '';
          else if (type === -4123) ok = cell?.f !== undefined;
          else if (type === 4) ok = !this.filled(row, c);
          else if (type === 12) ok = !this.sheet.isColHidden(c);
          else if (type === -4144) ok = !!cell?.note;
          else throw new VbError(1004, 'Gridline supports SpecialCells for constants, formulas, blanks, visible cells, comments and the last cell.');
          if (ok) push(row, c);
        }
      }
    }
    if (!hits.length) throw new VbError(1004, 'No cells were found.');
    return new RangeObj(this.s, this.sheet, mergeRows(hits));
  }

  private autoFilter(field: Value | undefined, crit: Value | undefined, operator: Value | undefined): void {
    this.s.settle();
    const existing = this.sheet.filter;
    let range = existing?.range ?? this.used(this.r);
    if (!existing && range.r1 === range.r2 && range.c1 === range.c2) range = currentRegion(this.doc, this.sheet, range.r1, range.c1);
    if (field === undefined) {
      this.s.edit((tx) => applyFilter(tx, this.doc, this.sheet, existing ? null : { range, conds: {} }));
      return;
    }
    const col = range.c1 + Math.round(toNum(field)) - 1;
    const conds = { ...(existing?.conds ?? {}) };
    if (crit === undefined) delete conds[col];
    else if (crit instanceof VbArray || (operator !== undefined && toNum(operator) === 7)) {
      const list: Value[] = crit instanceof VbArray ? crit.data : [crit];
      conds[col] = { values: list.map((x) => toStr(x).replace(/^=/, '')) };
    } else conds[col] = criteriaCond(toStr(crit), operator !== undefined ? toNum(operator) : undefined);
    this.s.edit((tx) => applyFilter(tx, this.doc, this.sheet, { range, conds }));
  }

  // ---------- borders ----------

  /** Changes one border (by Excel index) across the range. */
  setBorder(idx: number, fn: (old: BorderSide | undefined) => BorderSide | undefined): void {
    for (const a of this.areas) {
      this.s.edit((tx) =>
        applyStyle(tx, this.doc.wb, this.sheet, [a], (cur, row, col) => {
          const sides = sidesFor(idx, row, col, a);
          if (!sides.length) return null;
          const p: StylePatch = {};
          for (const side of sides) p[side] = fn(cur[side]) ?? null;
          return p;
        }),
      );
    }
  }

  border(idx: number): BorderSide | undefined {
    const r = this.r;
    const row = idx === 9 ? r.r2 : r.r1;
    const col = idx === 10 ? r.c2 : r.c1;
    const side = sidesFor(idx, row, col, r)[0] ?? 'borderTop';
    return this.style(row, col)[side];
  }

  styleAt(): Style {
    return this.style();
  }

  patchStyle(p: StylePatch): void {
    this.patch(p);
  }
}

// ---------- helpers ----------

/** Joins single-row segments into rectangles where they line up. */
function mergeRows(list: Range[]): Range[] {
  const out: Range[] = [];
  for (const r of list) {
    const prev = out.find((o) => o.c1 === r.c1 && o.c2 === r.c2 && o.r2 === r.r1 - 1);
    if (prev) prev.r2 = r.r2;
    else out.push({ ...r });
  }
  return out;
}

/** #N/A, for cells beyond the end of an array written to a bigger range. */
const NA: VbObject = { vbType: 'Error', get: () => 2042, set: () => undefined, dflt: () => 2042 };

/** Reads an array assigned to a range: 2D arrays by position, 1D arrays as a row. */
function arrayReader(arr: VbArray): (i: number, j: number) => Value {
  if (arr.dims.length === 2) {
    const [d0, d1] = arr.dims;
    return (i, j) => (d0.lb + i <= d0.ub && d1.lb + j <= d1.ub ? arr.get([d0.lb + i, d1.lb + j]) : NA);
  }
  if (arr.dims.length === 1) return (_i, j) => (j < arr.data.length ? arr.data[j] : NA);
  return () => NA;
}

function pasteMode(paste: Value | undefined, transpose: Value | undefined): PasteMode {
  if (transpose !== undefined && toBool(transpose)) return 'transpose';
  switch (paste === undefined ? -4104 : toNum(paste)) {
    case -4163:
    case 12:
      return 'values';
    case -4122:
      return 'formats';
    case -4123:
      return 'formulas';
    case 7:
      return 'noBorders';
    default:
      return 'all';
  }
}

/** Moves cells to make room (insert) or close a gap (delete). References from other cells don't follow. */
function shiftCells(tx: Transaction, sheet: Sheet, a: Range, dir: 'up' | 'left' | 'down' | 'right'): void {
  const b = sheet.bounds();
  const h = a.r2 - a.r1 + 1;
  const w = a.c2 - a.c1 + 1;
  const move = (fr: number, fc: number, tr: number, tc: number) => tx.setCell(sheet, tr, tc, fr > MAX_ROWS - 1 || fc > MAX_COLS - 1 ? undefined : sheet.getCell(fr, fc));
  if (dir === 'up') {
    for (let c = a.c1; c <= a.c2; c++) for (let r = a.r1; r <= Math.max(a.r2, b.maxRow); r++) move(r + h, c, r, c);
  } else if (dir === 'left') {
    for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= Math.max(a.c2, b.maxCol); c++) move(r, c + w, r, c);
  } else if (dir === 'down') {
    if (b.maxRow + h >= MAX_ROWS) throw new VbError(1004, "There's no room to insert cells there.");
    for (let c = a.c1; c <= a.c2; c++) {
      for (let r = b.maxRow; r >= a.r1; r--) move(r, c, r + h, c);
      for (let r = a.r1; r <= a.r2; r++) tx.setCell(sheet, r, c, undefined);
    }
  } else {
    if (b.maxCol + w >= MAX_COLS) throw new VbError(1004, "There's no room to insert cells there.");
    for (let r = a.r1; r <= a.r2; r++) {
      for (let c = b.maxCol; c >= a.c1; c--) move(r, c, r, c + w);
      for (let c = a.c1; c <= a.c2; c++) tx.setCell(sheet, r, c, undefined);
    }
  }
}

/** Excel wildcards (* ? ~) to a regular expression. */
export function wildcard(what: string, whole: boolean, matchCase: boolean, global = false): RegExp {
  let src = '';
  for (let i = 0; i < what.length; i++) {
    const ch = what[i];
    if (ch === '~' && i + 1 < what.length) src += what[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') src += '.*';
    else if (ch === '?') src += '.';
    else src += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(whole ? `^${src}$` : src, (matchCase ? '' : 'i') + (global ? 'g' : ''));
}

/** AutoFilter Criteria1 like ">10", "=Paid", "*north*". */
function criteriaCond(crit: string, operator?: number): FilterCond {
  if (operator === 3) return { op: { type: 'top10', value: crit } };
  if (operator === 4) return { op: { type: 'top10', value: crit } };
  const m = /^(>=|<=|<>|>|<|=)?(.*)$/.exec(crit)!;
  const value = m[2];
  const ops: Record<string, FilterOp> = { '>=': 'gte', '<=': 'lte', '<>': 'notEquals', '>': 'gt', '<': 'lt' };
  if (m[1] && m[1] !== '=') return { op: { type: ops[m[1]], value } };
  if (value.startsWith('*') && value.endsWith('*') && value.length > 1) return { op: { type: 'contains', value: value.slice(1, -1) } };
  if (value.endsWith('*')) return { op: { type: 'beginsWith', value: value.slice(0, -1) } };
  if (value.startsWith('*')) return { op: { type: 'endsWith', value: value.slice(1) } };
  return { op: { type: 'equals', value } };
}

// ---------- Interior / Font / Borders ----------

class InteriorObj implements VbObject {
  readonly vbType = 'Interior';
  constructor(private rng: RangeObj) {}

  get(name: string): Value {
    const fill = this.rng.styleAt().fill;
    switch (name) {
      case 'color':
        return fill ? hexToVb(fill) : 16777215;
      case 'colorindex':
        return paletteIndex(fill);
      case 'pattern':
        return fill ? 1 : -4142;
      case 'themecolor':
      case 'tintandshade':
      case 'patterncolor':
      case 'patterncolorindex':
        return 0;
    }
    throw unsupported('Interior', name);
  }

  set(name: string, _args: Value[], v: Value): void {
    switch (name) {
      case 'color':
        return this.rng.patchStyle({ fill: vbToHex(v) });
      case 'colorindex':
        return this.rng.patchStyle({ fill: paletteColor(v) });
      case 'pattern':
        if (toNum(v) === -4142) this.rng.patchStyle({ fill: null });
        return;
      case 'themecolor':
      case 'tintandshade':
      case 'patterncolor':
      case 'patterncolorindex':
      case 'patterntintandshade':
        return;
    }
    throw unsupported('Interior', name);
  }
}

class FontObj implements VbObject {
  readonly vbType = 'Font';
  constructor(private rng: RangeObj) {}

  get(name: string): Value {
    const st = this.rng.styleAt();
    switch (name) {
      case 'bold':
        return !!st.bold;
      case 'italic':
        return !!st.italic;
      case 'underline':
        return st.underline ? 2 : -4142;
      case 'strikethrough':
        return !!st.strike;
      case 'size':
        return st.fontSize ?? 11;
      case 'name':
        return st.fontName ?? 'Calibri';
      case 'color':
        return st.color ? hexToVb(st.color) : 0;
      case 'colorindex':
        return st.color ? paletteIndex(st.color) : -4105;
      case 'themecolor':
      case 'tintandshade':
        return 0;
    }
    throw unsupported('Font', name);
  }

  set(name: string, _args: Value[], v: Value): void {
    switch (name) {
      case 'bold':
        return this.rng.patchStyle({ bold: toBool(v) || null });
      case 'italic':
        return this.rng.patchStyle({ italic: toBool(v) || null });
      case 'underline': {
        const on = typeof v === 'boolean' ? v : toNum(v) !== -4142 && toNum(v) !== 0;
        return this.rng.patchStyle({ underline: on || null });
      }
      case 'strikethrough':
        return this.rng.patchStyle({ strike: toBool(v) || null });
      case 'size': {
        const n = toNum(v);
        if (!(n >= 1 && n <= 409)) throw new VbError(1004, 'Font size must be between 1 and 409.');
        return this.rng.patchStyle({ fontSize: n });
      }
      case 'name':
        return this.rng.patchStyle({ fontName: toStr(v) || null });
      case 'color':
        return this.rng.patchStyle({ color: vbToHex(v) });
      case 'colorindex':
        return this.rng.patchStyle({ color: paletteColor(v) });
      case 'themecolor':
      case 'tintandshade':
      case 'fontstyle':
        return;
    }
    throw unsupported('Font', name);
  }
}

class BorderObj implements VbObject {
  readonly vbType = 'Border';
  constructor(
    private rng: RangeObj,
    private idx: number,
  ) {}

  get(name: string): Value {
    const side = this.rng.border(this.idx);
    const [line, weight] = fromBorderStyle(side?.style);
    switch (name) {
      case 'linestyle':
        return line;
      case 'weight':
        return weight;
      case 'color':
        return side?.color ? hexToVb(side.color) : 0;
      case 'colorindex':
        return side?.color ? paletteIndex(side.color) : -4105;
      case 'themecolor':
      case 'tintandshade':
        return 0;
    }
    throw unsupported('Border', name);
  }

  set(name: string, _args: Value[], v: Value): void {
    if (name === 'themecolor' || name === 'tintandshade') return;
    if (!['linestyle', 'weight', 'color', 'colorindex'].includes(name)) throw unsupported('Border', name);
    this.rng.setBorder(this.idx, (old) => {
      let [line, weight] = fromBorderStyle(old?.style);
      let color = old?.color;
      if (name === 'linestyle') {
        line = toNum(v);
        if (!old) weight = 2;
      } else if (name === 'weight') {
        weight = toNum(v);
        if (!old) line = 1;
      } else {
        color = name === 'color' ? vbToHex(v) : paletteColor(v) ?? undefined;
        if (!old) line = 1;
      }
      const style = toBorderStyle(line, weight);
      return style ? { style, color } : undefined;
    });
  }
}

/** Range.Borders without an index: applies to every edge, inside lines included. */
class BordersObj implements VbObject {
  readonly vbType = 'Borders';
  private all: BorderObj;
  constructor(private rng: RangeObj) {
    this.all = new BorderObj(rng, 0);
  }

  get(name: string, args: Value[]): Value {
    if (name === 'item') return new BorderObj(this.rng, toNum(args[0]));
    if (name === 'count') return 6;
    return this.all.get(name);
  }

  set(name: string, args: Value[], v: Value): void {
    this.all.set(name, args, v);
  }

  dflt(args: Value[]): Value {
    return new BorderObj(this.rng, toNum(args[0]));
  }

  items(): Value[] {
    return [7, 8, 9, 10, 11, 12].map((i) => new BorderObj(this.rng, i));
  }
}

class AreasObj implements VbObject {
  readonly vbType = 'Areas';
  constructor(private rng: RangeObj) {}

  get(name: string, args: Value[]): Value {
    if (name === 'count') return this.rng.areas.length;
    if (name === 'item') return this.dflt(args);
    throw unsupported('Areas', name);
  }

  set(name: string): void {
    throw unsupported('Areas', name);
  }

  dflt(args: Value[]): Value {
    const a = this.rng.areas[Math.round(toNum(args[0])) - 1];
    if (!a) throw new VbError(9, 'Subscript out of range.');
    return this.rng.sub(a);
  }

  items(): Value[] {
    return this.rng.areas.map((a) => this.rng.sub(a));
  }
}

class CommentObj implements VbObject {
  readonly vbType = 'Comment';
  constructor(private rng: RangeObj) {}

  get(name: string, args: Value[], named?: Record<string, Value>): Value {
    const r = this.rng.r;
    switch (name) {
      case 'text': {
        const t = opt(args, named, 0, 'text');
        if (t !== undefined) this.rng.s.edit((tx) => setNote(tx, this.rng.sheet, r.r1, r.c1, toStr(t)));
        return this.rng.sheet.getCell(r.r1, r.c1)?.note ?? '';
      }
      case 'delete':
        this.rng.s.edit((tx) => setNote(tx, this.rng.sheet, r.r1, r.c1, undefined));
        return EMPTY;
      case 'visible':
        return false;
      case 'parent':
        return this.rng;
      case 'author':
        return '';
    }
    throw unsupported('Comment', name);
  }

  set(name: string): void {
    if (name === 'visible') return;
    throw unsupported('Comment', name);
  }
}

class RangeLinksObj implements VbObject {
  readonly vbType = 'Hyperlinks';
  constructor(private rng: RangeObj) {}

  get(name: string, args: Value[], named?: Record<string, Value>): Value {
    const r = this.rng.r;
    if (name === 'delete') {
      this.rng.s.edit((tx) => clearRanges(tx, this.rng.sheet, this.rng.areas, 'links'));
      return EMPTY;
    }
    if (name === 'count') return this.rng.sheet.getCell(r.r1, r.c1)?.link ? 1 : 0;
    if (name === 'add') {
      const address = toStr(opt(args, named, 1, 'address') ?? '');
      this.rng.s.edit((tx) => setLink(tx, this.rng.sheet, r.r1, r.c1, address));
      return EMPTY;
    }
    throw unsupported('Hyperlinks', name);
  }

  set(name: string): void {
    throw unsupported('Hyperlinks', name);
  }
}

