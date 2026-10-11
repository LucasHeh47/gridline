import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseCell } from '../src/core/address';
import { Doc } from '../src/core/doc';
import { readXlsx, writeXlsx } from '../src/core/io/xlsx';
import { setCellInput } from '../src/core/ops/cells';
import { restructure } from '../src/core/ops/structure';
import { isValid, listItems, newValidationId, type DataValidation } from '../src/core/validation';
import { makeDoc } from './helpers';

function rule(p: Partial<DataValidation>): DataValidation {
  return { id: newValidationId(), ranges: [{ r1: 2, c1: 2, r2: 2, c2: 2 }], type: 'whole', operator: 'between', formula1: '1', allowBlank: true, dropdown: true, showError: true, errorStyle: 'stop', showPrompt: false, ...p };
}

/** Types `inputs` into C3 one by one and checks each against the rule, like Excel's Validation.Value. */
function check(doc: Doc, dv: DataValidation, inputs: (string | null)[]): boolean[] {
  const sheet = doc.wb.activeSheet;
  return inputs.map((text) => {
    doc.transact('t', (tx) => (text === null ? tx.setCell(sheet, 2, 2, undefined) : setCellInput(tx, doc, sheet, 2, 2, text)));
    return isValid(doc, sheet, dv, 2, 2, doc.engine.getValue(sheet, 2, 2));
  });
}

// Expected results were measured in Excel (Range.Validation.Value) with the same typed inputs.
describe('data validation matches Excel', () => {
  it('whole numbers', () => {
    const doc = makeDoc();
    expect(check(doc, rule({ formula1: '1', formula2: '500' }), ['1', '500', '0', '501', '2.5', 'abc', null, 'TRUE'])).toEqual([true, true, false, false, false, false, true, false]);
    expect(check(doc, rule({ formula1: '1', formula2: '500', allowBlank: false }), [null])).toEqual([false]);
    expect(check(doc, rule({ operator: 'notBetween', formula1: '1', formula2: '5' }), ['0', '3', '6'])).toEqual([true, false, true]);
  });

  it('decimals', () => {
    const doc = makeDoc();
    expect(check(doc, rule({ type: 'decimal', formula1: '0', formula2: '0.5' }), ['0', '0.5', '0.25', '0.51', '-0.1', 'x'])).toEqual([true, true, true, false, false, false]);
    expect(check(doc, rule({ type: 'decimal', operator: 'greaterThan', formula1: '10' }), ['10', '10.0001', 'x'])).toEqual([false, true, false]);
  });

  it('typed lists are case-sensitive, lists from cells are not', () => {
    const doc = makeDoc({ E1: 'a', E2: 'b', E3: 'c' });
    expect(check(doc, rule({ type: 'list', formula1: '"Web,Partner,Retail,Direct"' }), ['Web', 'web', 'WEB', ' Web', 'Web ', 'Partner', 'Other', '1', null])).toEqual([
      true,
      false,
      false,
      true,
      true,
      true,
      false,
      false,
      true,
    ]);
    expect(check(doc, rule({ type: 'list', formula1: '"1,2,3"' }), ['1', '2', '4', '1.0', "'2"])).toEqual([true, true, false, true, true]);
    expect(check(doc, rule({ type: 'list', formula1: '$E$1:$E$3' }), ['a', 'b', 'z', 'A'])).toEqual([true, true, false, true]);
    expect(listItems(doc, doc.wb.activeSheet, rule({ type: 'list', formula1: '$E$1:$E$3' }), 2, 2).items).toEqual(['a', 'b', 'c']);
  });

  it('text length, dates and custom formulas', () => {
    const doc = makeDoc();
    expect(check(doc, rule({ type: 'textLength', operator: 'lessThanOrEqual', formula1: '3' }), ['abc', 'abcd', '12', '1234', null])).toEqual([true, false, true, false, true]);
    expect(check(doc, rule({ type: 'date', formula1: 'DATE(2025,1,1)', formula2: 'DATE(2025,12,31)' }), ['3/1/2025', '1/1/2026', 'x', '45700'])).toEqual([true, false, false, true]);
    expect(check(doc, rule({ type: 'custom', formula1: 'C3>5' }), ['6', '4', 'x'])).toEqual([true, false, true]);
  });
});

describe('data validation in files', () => {
  it('saves and reads rules, and follows inserted rows', async () => {
    const doc = makeDoc({ A1: 'x' });
    const sheet = doc.wb.activeSheet;
    sheet.validations = [
      rule({ ranges: [{ r1: 4, c1: 4, r2: 1203, c2: 4 }], formula1: '1', formula2: '500', showError: false, allowBlank: false }),
      rule({ ranges: [{ r1: 4, c1: 6, r2: 1203, c2: 6 }], type: 'list', formula1: '"Web,Partner,Retail,Direct"', errorStyle: 'warning', error: 'Pick a channel', showPrompt: true, prompt: 'Choose one' }),
    ];
    const data = await writeXlsx(doc);
    const xml = await (await JSZip.loadAsync(data)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).toContain('<dataValidations count="2">');
    const { wb, warnings } = await readXlsx(data);
    expect(warnings).toEqual([]);
    const back = wb.sheets[0].validations;
    expect(back.map((v) => [v.type, v.formula1, v.formula2, v.showError, v.allowBlank, v.errorStyle, v.prompt])).toEqual([
      ['whole', '1', '500', false, false, 'stop', undefined],
      ['list', '"Web,Partner,Retail,Direct"', undefined, true, true, 'warning', 'Choose one'],
    ]);
    const d2 = new Doc(wb);
    d2.transact('ins', (tx) => restructure(tx, d2.wb, d2.wb.sheets[0], 'row', 0, 2));
    expect(d2.wb.sheets[0].validations[0].ranges[0]).toEqual({ r1: 6, c1: 4, r2: 1205, c2: 4 });
  });

  it('reads the x14 extension used for lists on other sheets', async () => {
    const doc = makeDoc();
    const data = await writeXlsx(doc);
    const zip = await JSZip.loadAsync(data);
    let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    xml = xml.replace(
      '</worksheet>',
      '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!$A$1:$A$5</xm:f></x14:formula1><xm:sqref>B2:B20</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>',
    );
    zip.file('xl/worksheets/sheet1.xml', xml);
    const { wb } = await readXlsx(new Uint8Array(await zip.generateAsync({ type: 'uint8array' })));
    expect(wb.sheets[0].validations[0]).toMatchObject({ type: 'list', formula1: 'Lists!$A$1:$A$5', ranges: [parseCellRange('B2:B20')] });
  });
});

function parseCellRange(s: string) {
  const [a, b] = s.split(':').map((x) => parseCell(x)!);
  return { r1: a.row, c1: a.col, r2: b.row, c2: b.col };
}
