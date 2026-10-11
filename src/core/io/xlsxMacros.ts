// Macros inside Excel packages:
//  • VBA code lives in xl/vbaProject.bin (only in macro-enabled .xlsm files).
//  • Sheet code names (Sheet1…) live in <sheetPr codeName> and <workbookPr codeName>.
//  • Gridline's own macro data (step macros, sheet buttons, shortcuts, triggers) is stored as JSON in a
//    custom XML part, which Excel keeps but ignores. It works in .xlsx files too.
import JSZip from 'jszip';
import { MacroProject, type MacroShortcut, type MacroTrigger, type ModuleKind, type SheetButton, type StepMacro, type VbaModule } from '../macros/project';
import type { Workbook } from '../workbook';
import { guid, invokeShortcuts, joinAttributes, readVbaProject, splitAttributes, writeVbaProject } from './vbaProject';
import { esc, parseRels, relsPathFor, resolvePath, sheetPaths, unesc } from './xlsxCharts';

export const GRIDLINE_NS = 'urn:gridline:macros:1';

const REL_VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
const REL_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const REL_CUSTOM_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';
const CT_VBA = 'application/vnd.ms-office.vbaProject';
const CT_WB = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const CT_WB_MACRO = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
const CT_CUSTOM_PROPS = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';

/** What Gridline stores in its custom XML part. */
interface GridlinePart {
  version: 1;
  stepMacros?: StepMacro[];
  shortcuts?: MacroShortcut[];
  triggers?: MacroTrigger[];
  buttons?: (SheetButton & { sheet: string })[];
}

// ---------- read ----------

export interface MacroReadResult {
  /** The file had VBA code. */
  hasVba: boolean;
  warnings: string[];
}

/** Reads code names, the VBA project and Gridline's macro data into `wb`. */
export async function readMacros(data: Uint8Array, wb: Workbook): Promise<MacroReadResult> {
  const zip = await JSZip.loadAsync(data);
  const warnings: string[] = [];
  const paths = await sheetPaths(zip);
  for (const sheet of wb.sheets) {
    const path = paths.get(sheet.name);
    const xml = path ? await zip.file(path)?.async('string') : undefined;
    const code = xml && /<sheetPr\b[^>]*\bcodeName="([^"]+)"/.exec(xml)?.[1];
    if (code) sheet.codeName = unesc(code);
  }
  const wbXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? '';
  const wbCode = unesc(/<workbookPr\b[^>]*\bcodeName="([^"]+)"/.exec(wbXml)?.[1] ?? 'ThisWorkbook');

  let project: MacroProject | null = null;
  let hasVba = false;
  const rels = parseRels((await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? '');
  const vbaRel = rels.find((r) => r.type === REL_VBA);
  const vbaPath = vbaRel ? resolvePath('xl/workbook.xml', vbaRel.target) : zip.file('xl/vbaProject.bin') ? 'xl/vbaProject.bin' : null;
  const vbaBytes = vbaPath ? await zip.file(vbaPath)?.async('uint8array') : undefined;
  if (vbaBytes) {
    try {
      const vba = readVbaProject(vbaBytes);
      project = new MacroProject();
      project.projectName = vba.name;
      project.original = vbaBytes;
      // Excel names the workbook module after <workbookPr codeName> (ThisWorkbook in English Excel)
      project.workbookModule = wbCode;
      for (const m of vba.modules) {
        const parts = splitAttributes(m.source);
        const kind: ModuleKind = m.kind;
        const mod: VbaModule = { name: m.name, kind, code: parts.code, attrs: { header: parts.header, procAttrs: parts.procAttrs } };
        project.modules.push(mod);
        for (const [proc, letter] of Object.entries(invokeShortcuts(parts.procAttrs))) {
          const procName = /^\s*(?:(?:Public|Private)\s+)?Sub\s+(\w+)/im.exec(parts.code.split(/\r?\n/).find((l) => new RegExp(`\\bSub\\s+${proc}\\b`, 'i').test(l)) ?? '')?.[1] ?? proc;
          project.shortcuts.push({ macro: `${m.name}.${procName}`, keys: letter === letter.toUpperCase() ? `Ctrl+Shift+${letter.toUpperCase()}` : `Ctrl+${letter.toUpperCase()}` });
        }
      }
      hasVba = project.hasCode;
      if (project.modules.some((m) => m.kind === 'form')) warnings.push("This file has VBA UserForms. They're kept when you save as .xlsm, but Gridline can't show them.");
    } catch (e) {
      warnings.push(`The macros in this file couldn't be read (${(e as Error).message}).`);
    }
  }

  const part = await readGridlinePart(zip);
  if (part) {
    project ??= new MacroProject();
    project.stepMacros = Array.isArray(part.stepMacros) ? part.stepMacros : [];
    if (Array.isArray(part.shortcuts)) {
      // Gridline's list is complete; it replaces shortcuts read from VBA attributes
      project.shortcuts = part.shortcuts.filter((s) => s && typeof s.macro === 'string' && typeof s.keys === 'string');
    }
    project.triggers = Array.isArray(part.triggers) ? part.triggers : [];
    for (const b of part.buttons ?? []) {
      const sheet = wb.sheetByName(b.sheet);
      if (!sheet || !b.anchor) continue;
      const { sheet: _ignored, ...button } = b;
      sheet.buttons.push(button);
    }
  }
  if (project) {
    project.codeChanged = false;
    wb.macros = project;
  }
  return { hasVba, warnings };
}

async function readGridlinePart(zip: JSZip): Promise<GridlinePart | null> {
  for (const f of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/i.test(f)) continue;
    const xml = await zip.file(f)!.async('string');
    if (!xml.includes(GRIDLINE_NS)) continue;
    const m = /<!\[CDATA\[([\s\S]*)\]\]>/.exec(xml);
    if (!m) continue;
    try {
      const json = JSON.parse(m[1].replace(/\]\]\]\]><!\[CDATA\[>/g, ']]>')) as GridlinePart;
      if (json && json.version === 1) return json;
    } catch {
      // ignore a damaged part
    }
  }
  return null;
}

// ---------- write ----------

/** The VBA module list as it must be saved: every sheet needs a document module named by its code name. */
export function modulesForSave(wb: Workbook): VbaModule[] {
  const project = wb.macros!;
  const out: VbaModule[] = [];
  const wbName = project.workbookModule;
  const wbModule = project.modules.find((m) => m.kind === 'document' && m.name.toLowerCase() === wbName.toLowerCase());
  out.push(wbModule ?? { name: wbName, kind: 'document', code: '' });
  for (const sheet of wb.sheets) {
    const name = sheet.codeName ?? '';
    const existing = project.modules.find((m) => m.kind === 'document' && m.name.toLowerCase() === name.toLowerCase());
    out.push(existing ?? { name, kind: 'document', code: '' });
  }
  for (const m of project.modules) if (m.kind !== 'document') out.push(m);
  return out;
}

/** Whether the saved module list differs from what's in the original vbaProject.bin. */
function sameModules(a: VbaModule[], b: VbaModule[]): boolean {
  if (a.length !== b.length) return false;
  const key = (m: VbaModule) => `${m.kind}:${m.name.toLowerCase()}`;
  const set = new Set(b.map(key));
  return a.every((m) => set.has(key(m)));
}

/** Shortcut letters to store as VB_Invoke_Func, per module: Ctrl+M → "m", Ctrl+Shift+M → "M". */
function vbaShortcuts(project: MacroProject): Map<string, Record<string, string>> {
  const out = new Map<string, Record<string, string>>();
  for (const s of project.shortcuts) {
    const m = /^Ctrl\+(Shift\+)?([A-Z])$/i.exec(s.keys);
    const dot = s.macro.indexOf('.');
    if (!m || dot < 0) continue;
    const mod = s.macro.slice(0, dot).toLowerCase();
    const proc = s.macro.slice(dot + 1).toLowerCase();
    const rec = out.get(mod) ?? {};
    rec[proc] = m[1] ? m[2].toUpperCase() : m[2].toLowerCase();
    out.set(mod, rec);
  }
  return out;
}

/**
 * Adds macros to a package written by ExcelJS. With `macroEnabled` the VBA project is included
 * (the file must be saved as .xlsm); otherwise only Gridline's own macro data is kept.
 */
export async function writeMacros(buf: Uint8Array, wb: Workbook, macroEnabled: boolean): Promise<Uint8Array> {
  const project = wb.macros;
  const hasButtons = wb.sheets.some((s) => s.buttons.length);
  const vba = macroEnabled && !!project?.hasCode;
  const hasPart = !!project && (project.stepMacros.length > 0 || project.triggers.length > 0 || project.shortcuts.length > 0);
  if (!vba && !hasPart && !hasButtons) return buf;

  const zip = await JSZip.loadAsync(buf);
  let ct = (await zip.file('[Content_Types].xml')!.async('string')) ?? '';
  let wbRels = (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? '';
  const rels = parseRels(wbRels);
  let nextRel = rels.reduce((n, r) => Math.max(n, parseInt(r.id.replace(/\D/g, ''), 10) || 0), 0) + 1;
  const addRel = (type: string, target: string) => {
    wbRels = wbRels.replace('</Relationships>', `<Relationship Id="rId${nextRel++}" Type="${type}" Target="${esc(target)}"/></Relationships>`);
  };

  if (vba && project) {
    const modules = modulesForSave(wb);
    const shortcuts = vbaShortcuts(project);
    const unchanged = !project.codeChanged && !!project.original && sameModules(modules, project.modules);
    const bin = unchanged
      ? project.original!
      : writeVbaProject(
          {
            name: project.projectName,
            modules: modules.map((m) => ({ name: m.name, kind: m.kind, source: joinAttributes(m.name, m.kind, m.code, m.attrs, shortcuts.get(m.name.toLowerCase())) })),
          },
          project.original,
        );
    zip.file('xl/vbaProject.bin', bin);
    if (!/Extension="bin"/i.test(ct)) ct = ct.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="bin" ContentType="${CT_VBA}"/>`);
    ct = ct.replace(CT_WB, CT_WB_MACRO);
    addRel(REL_VBA, 'vbaProject.bin');

    // Code names tie sheets to their VBA modules
    let wbXml = (await zip.file('xl/workbook.xml')!.async('string')) ?? '';
    const wbCode = esc(project.workbookModule);
    if (/<workbookPr\b/.test(wbXml)) {
      wbXml = /<workbookPr\b[^>]*\bcodeName=/.test(wbXml) ? wbXml : wbXml.replace(/<workbookPr\b/, `<workbookPr codeName="${wbCode}"`);
    } else {
      // workbookPr comes after fileVersion/fileSharing, before bookViews
      const pr = `<workbookPr codeName="${wbCode}"/>`;
      wbXml = /<\/fileSharing>|<fileSharing\b[^>]*\/>/.test(wbXml)
        ? wbXml.replace(/(<\/fileSharing>|<fileSharing\b[^>]*\/>)/, `$1${pr}`)
        : /<fileVersion\b[^>]*\/>/.test(wbXml)
          ? wbXml.replace(/(<fileVersion\b[^>]*\/>)/, `$1${pr}`)
          : wbXml.replace(/(<workbook\b[^>]*>)/, `$1${pr}`);
    }
    zip.file('xl/workbook.xml', wbXml);
    const paths = await sheetPaths(zip);
    for (const sheet of wb.sheets) {
      const path = paths.get(sheet.name);
      if (!path || !sheet.codeName) continue;
      let xml = await zip.file(path)!.async('string');
      const code = esc(sheet.codeName);
      if (/<sheetPr\b[^>]*\bcodeName=/.test(xml)) xml = xml.replace(/(<sheetPr\b[^>]*\bcodeName=")[^"]*"/, `$1${code}"`);
      else if (/<sheetPr\b/.test(xml)) xml = xml.replace(/<sheetPr\b/, `<sheetPr codeName="${code}"`);
      else xml = xml.replace(/(<worksheet\b[^>]*>)/, `$1<sheetPr codeName="${code}"/>`);
      zip.file(path, xml);
    }
  }

  if (project || hasButtons) {
    const part: GridlinePart = { version: 1 };
    if (project?.stepMacros.length) part.stepMacros = project.stepMacros;
    if (project?.shortcuts.length) part.shortcuts = project.shortcuts;
    if (project?.triggers.length) part.triggers = project.triggers;
    const buttons = wb.sheets.flatMap((s) => s.buttons.map((b) => ({ ...b, sheet: s.name })));
    if (buttons.length) part.buttons = buttons;
    if (Object.keys(part).length > 1) {
      let n = 1;
      while (zip.file(`customXml/item${n}.xml`)) n++;
      const json = JSON.stringify(part).replace(/]]>/g, ']]]]><![CDATA[>');
      zip.file(`customXml/item${n}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<gridline xmlns="${GRIDLINE_NS}"><![CDATA[${json}]]></gridline>`);
      zip.file(
        `customXml/itemProps${n}.xml`,
        `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="{${guid()}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${GRIDLINE_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
      );
      zip.file(
        relsPathFor(`customXml/item${n}.xml`),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_CUSTOM_PROPS}" Target="itemProps${n}.xml"/></Relationships>`,
      );
      if (!/Extension="xml"/i.test(ct)) ct = ct.replace(/(<Types\b[^>]*>)/, '$1<Default Extension="xml" ContentType="application/xml"/>');
      ct = ct.replace('</Types>', `<Override PartName="/customXml/itemProps${n}.xml" ContentType="${CT_CUSTOM_PROPS}"/></Types>`);
      addRel(REL_CUSTOM_XML, `../customXml/item${n}.xml`);
    }
  }

  zip.file('[Content_Types].xml', ct);
  zip.file('xl/_rels/workbook.xml.rels', wbRels);
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}
