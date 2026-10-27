// The suite: Baseline (home), Gridline (spreadsheets), Pixelline (images), Frameline (video) and
// Beatline (music). Each app notes where it runs from in a small shared file, so the others can find
// and start it, installed or portable. (This file is the same in all five apps.)
import { app } from 'electron';
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import fs from 'node:fs/promises';
import path from 'node:path';

export type SuiteAppId = 'baseline' | 'gridline' | 'pixelline' | 'frameline' | 'beatline';

export interface SuiteAppInfo {
  id: SuiteAppId;
  name: string;
  description: string;
  installed: boolean;
  exe: string | null;
  version: string | null;
  /** Where to download it. */
  download: string;
}

export const SUITE: { id: SuiteAppId; name: string; description: string; repo: string }[] = [
  { id: 'baseline', name: 'Baseline', description: 'Home for your apps and files', repo: 'LucasHeh47/baseline' },
  { id: 'gridline', name: 'Gridline', description: 'Spreadsheets', repo: 'LucasHeh47/gridline' },
  { id: 'pixelline', name: 'Pixelline', description: 'Images and pixel art', repo: 'LucasHeh47/pixelline' },
  { id: 'frameline', name: 'Frameline', description: 'Video editing', repo: 'LucasHeh47/frameline' },
  { id: 'beatline', name: 'Beatline', description: 'Music making and recording', repo: 'LucasHeh47/beatline' },
];

interface Entry {
  exe: string;
  version: string;
  seen: number;
}

/** Shared by the apps; it lives with Baseline's settings. */
export const suiteFile = () => path.join(app.getPath('appData'), 'Baseline', 'suite.json');

async function readSuite(): Promise<Record<string, Entry>> {
  try {
    const v = JSON.parse(await fs.readFile(suiteFile(), 'utf8')) as Record<string, Entry>;
    return v && typeof v === 'object' ? v : {};
  } catch {
    return {};
  }
}

/** Records where this app runs from (built copies only: a development run isn't something to launch). */
export async function registerSelf(id: SuiteAppId): Promise<void> {
  if (!app.isPackaged) return;
  // A portable build runs from a temporary copy; the file people start is in this variable
  const exe = process.env.PORTABLE_EXECUTABLE_FILE || process.execPath;
  try {
    const all = await readSuite();
    if (all[id]?.exe === exe && all[id]?.version === app.getVersion()) return;
    all[id] = { exe, version: app.getVersion(), seen: Date.now() };
    await fs.mkdir(path.dirname(suiteFile()), { recursive: true });
    await fs.writeFile(suiteFile(), JSON.stringify(all, null, 2));
  } catch {
    // not being found by the other apps isn't worth an error
  }
}

/** Where the installer puts an app (per-user first, then for everyone). */
function defaultExes(name: string): string[] {
  const local = process.env.LOCALAPPDATA;
  return [local ? path.join(local, 'Programs', name, `${name}.exe`) : '', path.join(process.env.ProgramFiles ?? 'C:\\Program Files', name, `${name}.exe`)].filter(Boolean);
}

export async function suiteApps(extra: Partial<Record<SuiteAppId, string>> = {}): Promise<SuiteAppInfo[]> {
  const known = await readSuite();
  return SUITE.map((a) => {
    const candidates = [extra[a.id], known[a.id]?.exe, ...defaultExes(a.name)].filter((p): p is string => !!p);
    const exe = candidates.find((p) => existsSync(p)) ?? null;
    const version = exe && known[a.id]?.exe === exe ? known[a.id].version : null;
    return { id: a.id, name: a.name, description: a.description, installed: !!exe, exe, version, download: `https://github.com/${a.repo}/releases` };
  });
}

/** Arguments one app may pass when starting another: files, and options like --new=16x16, --fps=24 or --bpm=90. */
export function safeLaunchArgs(args: unknown): string[] {
  if (!Array.isArray(args)) return [];
  return args.filter((a): a is string => typeof a === 'string' && a.length < 400 && (/^--(new|template|bg|name|fps|bpm)(=[^"\r\n]{0,120})?$/.test(a) || (path.isAbsolute(a) && existsSync(a))));
}

/** Starts another app of the suite. Returns false when it isn't installed. */
export async function launchSuiteApp(id: SuiteAppId, args: string[] = [], extra: Partial<Record<SuiteAppId, string>> = {}): Promise<boolean> {
  const info = (await suiteApps(extra)).find((a) => a.id === id);
  if (!info?.exe) return false;
  const child = spawn(info.exe, args, { detached: true, stdio: 'ignore' });
  child.on('error', () => undefined);
  child.unref();
  return true;
}
