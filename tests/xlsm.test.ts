import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Doc } from '../src/core/doc';
import { readVbaProject } from '../src/core/io/vbaProject';
import { readXlsx, writeXlsx } from '../src/core/io/xlsx';
import { MacroProject } from '../src/core/macros/project';
import { ensureDocumentModules } from '../src/core/macros/runner';
import { newStep } from '../src/core/macros/steps';
import { makeDoc, val } from './helpers';

function docWithMacros(): Doc {
  const doc = makeDoc({ A1: 'hello' });
  const project = (doc.wb.macros = new MacroProject());
  ensureDocumentModules(doc);
  project.addModule('standard', 'Module1', 'Sub Main()\r\n    Range("B1").Value = 42\r\nEnd Sub');
  project.setCode('ThisWorkbook', 'Private Sub Workbook_Open()\r\n    Range("C1").Value = "opened"\r\nEnd Sub');
  project.shortcuts.push({ macro: 'Module1.Main', keys: 'Ctrl+Shift+M' });
  project.stepMacros.push({ id: 'x', name: 'Color it', steps: [newStep('fill')], variables: [{ name: 'Price', from: 'cell', value: 'B2' }] });
  project.triggers.push({ id: 't', macro: 'Color it', event: 'open' });
  doc.wb.sheets[0].buttons.push({ id: 'b1', label: 'Run', macro: 'Module1.Main', anchor: { row: 2, col: 1, dx: 4, dy: 4, w: 120, h: 32 } });
  return doc;
}

describe('.xlsm files', () => {
  it('saves VBA, code names and Gridline macro data and reads them back', async () => {
    const data = await writeXlsx(docWithMacros(), { macroEnabled: true });
    const zip = await JSZip.loadAsync(data);
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('application/vnd.ms-excel.sheet.macroEnabled.main+xml');
    expect(ct).toContain('application/vnd.ms-office.vbaProject');
    expect(await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).toContain('relationships/vbaProject');
    expect(await zip.file('xl/workbook.xml')!.async('string')).toMatch(/<workbookPr\b[^>]*codeName="ThisWorkbook"/);
    expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toMatch(/<sheetPr\b[^>]*codeName="Sheet1"/);
    const vba = readVbaProject(await zip.file('xl/vbaProject.bin')!.async('uint8array'));
    expect(vba.modules.map((m) => m.name)).toEqual(['ThisWorkbook', 'Sheet1', 'Module1']);
    expect(vba.modules[2].source).toContain('Attribute Main.VB_ProcData.VB_Invoke_Func = "M\\n14"');

    const { wb, hasVba, warnings } = await readXlsx(data);
    expect(warnings).toEqual([]);
    expect(hasVba).toBe(true);
    expect(wb.sheets[0].codeName).toBe('Sheet1');
    const p = wb.macros!;
    expect(p.module('Module1')!.code).toBe('Sub Main()\r\n    Range("B1").Value = 42\r\nEnd Sub');
    expect(p.module('ThisWorkbook')!.code).toContain('Workbook_Open');
    expect(p.stepMacros[0].name).toBe('Color it');
    expect(p.stepMacros[0].variables).toEqual([{ name: 'Price', from: 'cell', value: 'B2' }]);
    expect(p.triggers[0].event).toBe('open');
    expect(p.shortcuts).toEqual([{ macro: 'Module1.Main', keys: 'Ctrl+Shift+M' }]);
    expect(wb.sheets[0].buttons[0]).toMatchObject({ label: 'Run', macro: 'Module1.Main' });
    expect(p.codeChanged).toBe(false);
    expect(p.original).not.toBeNull();
  });

  it('keeps the original vbaProject.bin when the code is unchanged', async () => {
    const first = await writeXlsx(docWithMacros(), { macroEnabled: true });
    const { wb } = await readXlsx(first);
    const doc = new Doc(wb);
    const second = await writeXlsx(doc, { macroEnabled: true });
    const bin = async (d: Uint8Array) => (await JSZip.loadAsync(d)).file('xl/vbaProject.bin')!.async('uint8array');
    expect(Array.from(await bin(second))).toEqual(Array.from(await bin(first)));
    // Editing code rebuilds it
    doc.wb.macros!.setCode('Module1', 'Sub Main()\r\n    Range("B1").Value = 7\r\nEnd Sub');
    const third = await writeXlsx(doc, { macroEnabled: true });
    const back = await readXlsx(third);
    expect(back.wb.macros!.module('Module1')!.code).toContain('= 7');
    expect(val(new Doc(back.wb), 'A1')).toBe('hello');
  });

  it('drops VBA but keeps step macros in a plain .xlsx', async () => {
    const data = await writeXlsx(docWithMacros());
    const zip = await JSZip.loadAsync(data);
    expect(zip.file('xl/vbaProject.bin')).toBeNull();
    expect(await zip.file('[Content_Types].xml')!.async('string')).not.toContain('macroEnabled');
    const { wb, hasVba } = await readXlsx(data);
    expect(hasVba).toBe(false);
    expect(wb.macros!.stepMacros).toHaveLength(1);
    expect(wb.sheets[0].buttons).toHaveLength(1);
  });

  it('adds document modules for new sheets', async () => {
    const doc = docWithMacros();
    const extra = doc.wb.sheets[0];
    doc.transact('add', (tx) => {
      tx.setProp(doc.wb, 'sheets', [...doc.wb.sheets, Object.assign(new (extra.constructor as new (n: string) => typeof extra)('Data'), { codeName: 'Sheet2' })], true);
    });
    const data = await writeXlsx(doc, { macroEnabled: true });
    const vba = readVbaProject(await (await JSZip.loadAsync(data)).file('xl/vbaProject.bin')!.async('uint8array'));
    expect(vba.modules.map((m) => m.name)).toEqual(['ThisWorkbook', 'Sheet1', 'Sheet2', 'Module1']);
  });
});
