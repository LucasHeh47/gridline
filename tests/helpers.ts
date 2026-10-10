import { parseCell } from '../src/core/address';
import { Doc } from '../src/core/doc';
import { setCellInput } from '../src/core/ops/cells';
import { ErrorValue, type Scalar } from '../src/core/values';
import type { Sheet } from '../src/core/workbook';

/** Creates a document and fills cells from a map like { A1: '5', B1: '=A1*2' }. */
export function makeDoc(cells: Record<string, string> = {}): Doc {
  const doc = new Doc();
  if (Object.keys(cells).length) set(doc, cells);
  return doc;
}

export function set(doc: Doc, cells: Record<string, string>, sheet?: Sheet): void {
  const s = sheet ?? doc.wb.activeSheet;
  doc.transact('test', (tx) => {
    for (const [addr, raw] of Object.entries(cells)) {
      const a = parseCell(addr)!;
      setCellInput(tx, doc, s, a.row, a.col, raw);
    }
  });
}

export function val(doc: Doc, addr: string, sheet?: Sheet): Scalar | string {
  const a = parseCell(addr)!;
  const v = doc.engine.getValue(sheet ?? doc.wb.activeSheet, a.row, a.col);
  return v instanceof ErrorValue ? v.code : v;
}

/** Evaluates a formula in a scratch cell (Z1000). */
export function calc(doc: Doc, formula: string): Scalar | string {
  set(doc, { Z1000: formula.startsWith('=') ? formula : '=' + formula });
  return val(doc, 'Z1000');
}
