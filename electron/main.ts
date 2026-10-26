// Electron main process: windows, native dialogs, file access and the clipboard.
import { app, BrowserWindow, clipboard, ClipboardItem, dialog, ipcMain, Menu, net, shell, type IpcMainInvokeEvent } from 'electron';
import { createHash } from 'node:crypto';
import { existsSync } from 'node:fs';
import fs from 'node:fs/promises';
import path from 'node:path';
import { launchSuiteApp, registerSelf, safeLaunchArgs, suiteApps, type SuiteAppId } from './suite';
import { setupUpdates } from './updater';

const DEV_URL = process.env.VITE_DEV_SERVER_URL;
const SMOKE_SCREENSHOT = process.env.GRIDLINE_SCREENSHOT;
const SUPPORTED = /\.(xlsx|xlsm|csv|tsv|txt)$/i;

/** Paths the UI may read or write: chosen in a dialog, passed on the command line, dropped, or recent. */
const allowedPaths = new Set<string>();
const dirty = new Map<number, boolean>();
const closing = new Set<number>();

function allow(p: string): string {
  const full = path.resolve(p);
  allowedPaths.add(full.toLowerCase());
  return full;
}

function isAllowed(p: string): boolean {
  return allowedPaths.has(path.resolve(p).toLowerCase());
}

// ---------- recent files ----------

const recentPath = () => path.join(app.getPath('userData'), 'recent.json');
interface Recent {
  path: string;
  name: string;
  openedAt: number;
  pinned?: boolean;
}

async function readRecent(): Promise<Recent[]> {
  try {
    const list = JSON.parse(await fs.readFile(recentPath(), 'utf8')) as Recent[];
    return list.filter((r) => typeof r.path === 'string' && existsSync(r.path));
  } catch {
    return [];
  }
}

async function writeRecent(list: Recent[]): Promise<void> {
  // Keep every pinned file plus the 30 most recent others
  let unpinned = 0;
  const kept = list.filter((r) => r.pinned || unpinned++ < 30);
  await fs.mkdir(app.getPath('userData'), { recursive: true });
  await fs.writeFile(recentPath(), JSON.stringify(kept, null, 2));
}

async function addRecent(p: string): Promise<void> {
  const all = await readRecent();
  const prev = all.find((r) => r.path.toLowerCase() === p.toLowerCase());
  const list = all.filter((r) => r !== prev);
  list.unshift({ path: p, name: path.basename(p), openedAt: Date.now(), pinned: prev?.pinned || undefined });
  await writeRecent(list);
  app.addRecentDocument(p);
}

// ---------- windows ----------

/** What Baseline (or a shortcut) asked for: a blank workbook (--new) or a preset (--template=budget). */
function startFromArgv(argv: string[]): { template?: string; blank?: boolean } {
  const t = argv.find((a) => a.startsWith('--template='))?.slice('--template='.length);
  if (t && /^[a-z-]{1,32}$/.test(t)) return { template: t };
  return argv.includes('--new') ? { blank: true } : {};
}

function filesFromArgv(argv: string[]): string[] {
  return argv.slice(1).filter((a) => !a.startsWith('-') && SUPPORTED.test(a) && existsSync(a));
}

function createWindow(filePath?: string, template?: string, launch = false): BrowserWindow {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 720,
    minHeight: 460,
    show: false,
    backgroundColor: '#121417',
    title: 'Gridline',
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    titleBarStyle: 'hidden',
    titleBarOverlay: { color: '#16181d', symbolColor: '#c9ced8', height: 40 },
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
      spellcheck: false,
    },
  });
  const id = win.webContents.id;
  if (filePath) allow(filePath);
  const query: Record<string, string> | undefined = filePath ? { file: path.resolve(filePath) } : template ? { template } : launch ? { launch: '1' } : undefined;
  if (DEV_URL) win.loadURL(DEV_URL + (query ? `?${new URLSearchParams(query)}` : ''));
  else win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'), { query });

  win.once('ready-to-show', () => {
    win.show();
    if (SMOKE_SCREENSHOT) {
      setTimeout(async () => {
        const img = await win.webContents.capturePage();
        await fs.writeFile(SMOKE_SCREENSHOT, img.toPNG());
        app.exit(0);
      }, Number(process.env.GRIDLINE_SCREENSHOT_DELAY ?? 2500));
    }
  });
  win.on('close', (e) => {
    if (dirty.get(id) && !closing.has(id)) {
      e.preventDefault();
      win.webContents.send('app:request-close');
    }
  });
  win.on('closed', () => {
    dirty.delete(id);
    closing.delete(id);
  });
  // Never navigate away from the app; open links in the default browser.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?:|mailto:)/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (url !== win.webContents.getURL()) e.preventDefault();
  });
  win.webContents.on('before-input-event', (_e, input) => {
    if (input.type === 'keyDown' && input.key === 'F12' && DEV_URL) win.webContents.toggleDevTools();
  });
  return win;
}

const senderWindow = (e: IpcMainInvokeEvent | Electron.IpcMainEvent) => BrowserWindow.fromWebContents(e.sender);

// ---------- IPC ----------

async function readFileFor(p: string) {
  const full = allow(p);
  const data = await fs.readFile(full);
  await addRecent(full);
  return { path: full, name: path.basename(full), data: new Uint8Array(data) };
}

ipcMain.handle('file:open-dialog', async (e) => {
  const win = senderWindow(e);
  const res = await dialog.showOpenDialog(win!, {
    title: 'Open spreadsheet',
    properties: ['openFile'],
    filters: [
      { name: 'Spreadsheets', extensions: ['xlsx', 'xlsm', 'csv', 'tsv', 'txt'] },
      { name: 'Excel workbook', extensions: ['xlsx', 'xlsm'] },
      { name: 'CSV / TSV', extensions: ['csv', 'tsv', 'txt'] },
    ],
  });
  if (res.canceled || !res.filePaths[0]) return null;
  return readFileFor(res.filePaths[0]);
});

ipcMain.handle('file:open-path', async (_e, p: string) => {
  if (typeof p !== 'string') return null;
  const recent = await readRecent();
  if (!isAllowed(p) && !recent.some((r) => r.path.toLowerCase() === path.resolve(p).toLowerCase())) {
    throw new Error('That file can only be opened from the Open dialog.');
  }
  return readFileFor(p);
});

ipcMain.handle('file:grant', async (_e, p: string) => {
  if (typeof p === 'string' && p && SUPPORTED.test(p)) {
    const full = allow(p);
    await addRecent(full);
    return full;
  }
  return null;
});

ipcMain.handle('file:save-dialog', async (e, defaultName: string, kind: string) => {
  const win = senderWindow(e);
  const filters =
    kind === 'csv'
      ? [{ name: 'CSV (comma-separated)', extensions: ['csv'] }]
      : kind === 'tsv'
        ? [{ name: 'Tab-separated text', extensions: ['tsv', 'txt'] }]
        : kind === 'json'
          ? [{ name: 'Gridline settings', extensions: ['json'] }]
          : kind === 'xlsm'
            ? [{ name: 'Excel macro-enabled workbook', extensions: ['xlsm'] }, { name: 'Excel workbook (no VBA code)', extensions: ['xlsx'] }]
            : [{ name: 'Excel workbook', extensions: ['xlsx'] }, { name: 'Excel macro-enabled workbook', extensions: ['xlsm'] }];
  const res = await dialog.showSaveDialog(win!, { title: kind === 'json' ? 'Export settings' : 'Save as', defaultPath: defaultName, filters });
  if (res.canceled || !res.filePath) return null;
  return allow(res.filePath);
});

ipcMain.handle('file:save', async (_e, p: string, data: Uint8Array) => {
  if (!isAllowed(p)) return { ok: false, error: 'Gridline isn\'t allowed to write to that location. Use Save As.' };
  const full = path.resolve(p);
  const tmp = `${full}.gridline-${process.pid}.tmp`;
  try {
    // Write to a temporary file first so a crash can't leave a half-written spreadsheet.
    await fs.writeFile(tmp, Buffer.from(data));
    await fs.rename(tmp, full);
    if (SUPPORTED.test(full)) await addRecent(full);
    return { ok: true };
  } catch (err) {
    await fs.rm(tmp, { force: true }).catch(() => undefined);
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'EBUSY' || code === 'EPERM' || code === 'EACCES') {
      return { ok: false, error: 'The file is locked — it may be open in another program (like Excel). Close it there and try again, or use Save As.' };
    }
    return { ok: false, error: (err as Error).message };
  }
});

ipcMain.handle('file:recent', async () => {
  const list = await readRecent();
  for (const r of list) allow(r.path);
  return list;
});

ipcMain.handle('file:clear-recent', async () => {
  // Pinned files stay
  await writeRecent((await readRecent()).filter((r) => r.pinned)).catch(() => undefined);
  app.clearRecentDocuments();
});

ipcMain.handle('file:remove-recent', async (_e, p: string) => {
  if (typeof p !== 'string') return;
  await writeRecent((await readRecent()).filter((r) => r.path.toLowerCase() !== p.toLowerCase()));
});

ipcMain.handle('file:pin-recent', async (_e, p: string, pinned: boolean) => {
  if (typeof p !== 'string') return;
  const list = await readRecent();
  for (const r of list) if (r.path.toLowerCase() === p.toLowerCase()) r.pinned = pinned ? true : undefined;
  await writeRecent(list);
});

ipcMain.handle('clipboard:read', async () => {
  const text = await clipboard.readText();
  let html = '';
  try {
    for (const item of await clipboard.read()) {
      if (!item.types.includes('text/html')) continue;
      const blob = await item.getType('text/html');
      html = typeof blob === 'string' ? blob : await (blob as Blob).text();
      break;
    }
  } catch {
    // HTML is optional
  }
  return { text, html };
});
ipcMain.handle('clipboard:write', async (_e, d: { text: string; html?: string }) => {
  const entries: Record<string, string> = { 'text/plain': String(d?.text ?? '') };
  if (d?.html) entries['text/html'] = String(d.html);
  await clipboard.write([new ClipboardItem(entries)]);
});

ipcMain.on('win:set-title', (e, title: string) => senderWindow(e)?.setTitle(String(title)));
ipcMain.on('win:set-dirty', (e, value: boolean) => {
  const win = senderWindow(e);
  dirty.set(e.sender.id, !!value);
  win?.setDocumentEdited?.(!!value);
});
ipcMain.on('win:set-theme', (e, t: { dark: boolean; bar: string; symbol: string }) => {
  const win = senderWindow(e);
  try {
    win?.setTitleBarOverlay?.({ color: t.bar, symbolColor: t.symbol, height: 40 });
    win?.setBackgroundColor(t.dark ? '#121417' : '#ffffff');
  } catch {
    // not supported on this platform
  }
});
ipcMain.on('win:new', (_e, p?: string, template?: unknown) => {
  if (p && !isAllowed(p)) return;
  createWindow(p, typeof template === 'string' && /^[a-z-]{1,32}$/.test(template) ? template : undefined);
});
ipcMain.on('win:confirm-close', (e) => {
  closing.add(e.sender.id);
  senderWindow(e)?.close();
});
ipcMain.on('win:devtools', (e) => e.sender.toggleDevTools());
ipcMain.on('shell:open-external', (_e, url: string) => {
  if (typeof url === 'string' && /^(https?:|mailto:)/i.test(url)) shell.openExternal(url);
});
ipcMain.on('shell:show-in-folder', (_e, p: string) => {
  if (typeof p === 'string' && isAllowed(p)) shell.showItemInFolder(p);
});
ipcMain.handle('app:version', () => app.getVersion());

// ---------- live data (NFL stats, prices, exchange rates) ----------

/** Only these sites can be contacted, and only over HTTPS. */
const DATA_HOSTS = new Set([
  'github.com',
  'raw.githubusercontent.com',
  'api.frankfurter.dev',
  'api.coinbase.com',
  'query1.finance.yahoo.com',
  'query2.finance.yahoo.com',
  'api.open-meteo.com',
  'geocoding-api.open-meteo.com',
  'date.nager.at',
  'api.worldbank.org',
]);
const cacheDir = () => path.join(app.getPath('userData'), 'data-cache');

ipcMain.handle('net:fetch', async (_e, url: string, opts: { maxAgeSec?: number; headers?: Record<string, string>; offline?: boolean }) => {
  let u: URL;
  try {
    u = new URL(String(url));
  } catch {
    return { ok: false, status: 0, error: 'Invalid address' };
  }
  if (u.protocol !== 'https:' || !DATA_HOSTS.has(u.hostname)) return { ok: false, status: 0, error: `Gridline doesn't download data from ${u.hostname}.` };
  const file = path.join(cacheDir(), createHash('sha1').update(u.href).digest('hex') + '.txt');
  const maxAge = Math.max(0, Number(opts?.maxAgeSec ?? 0)) * 1000;
  let cached: { text: string; age: number } | null = null;
  try {
    const st = await fs.stat(file);
    cached = { text: await fs.readFile(file, 'utf8'), age: Date.now() - st.mtimeMs };
  } catch {
    // not cached yet
  }
  if (cached && cached.age < maxAge) return { ok: true, status: 200, text: cached.text, cached: true };
  if (opts?.offline) {
    if (cached) return { ok: true, status: 200, text: cached.text, cached: true, stale: true };
    return { ok: false, status: 0, error: "Offline mode is on and this data hasn't been downloaded yet" };
  }
  try {
    const headers: Record<string, string> = {};
    for (const [k, v] of Object.entries(opts?.headers ?? {})) if (/^(user-agent|accept)$/i.test(k)) headers[k] = String(v);
    const res = await net.fetch(u.href, { headers, redirect: 'follow' });
    if (!res.ok) {
      // A stale copy beats nothing (e.g. offline or rate limited)
      if (cached && res.status !== 404) return { ok: true, status: 200, text: cached.text, cached: true, stale: true };
      return { ok: false, status: res.status, error: `HTTP ${res.status}` };
    }
    const text = await res.text();
    await fs.mkdir(cacheDir(), { recursive: true });
    await fs.writeFile(file, text, 'utf8');
    return { ok: true, status: 200, text };
  } catch (err) {
    if (cached) return { ok: true, status: 200, text: cached.text, cached: true, stale: true };
    return { ok: false, status: 0, error: (err as Error).message || 'No internet connection' };
  }
});

ipcMain.handle('net:clear-cache', async () => {
  await fs.rm(cacheDir(), { recursive: true, force: true }).catch(() => undefined);
});

ipcMain.handle('net:cache-size', async () => {
  let total = 0;
  try {
    for (const f of await fs.readdir(cacheDir())) total += (await fs.stat(path.join(cacheDir(), f))).size;
  } catch {
    // no cache yet
  }
  return total;
});

// The other apps of the suite (Baseline, Pixelline)
ipcMain.handle('suite:apps', () => suiteApps());
ipcMain.handle('suite:launch', async (_e, id: SuiteAppId, args: unknown) => {
  if (!['baseline', 'gridline', 'pixelline'].includes(id)) return false;
  return launchSuiteApp(id, safeLaunchArgs(args));
});

setupUpdates();

// ---------- lifecycle ----------

/** Lets a window open these files and asks it to (each becomes a tab). */
function sendPaths(win: BrowserWindow, files: string[]): void {
  const full = files.map((f) => allow(f));
  win.webContents.send('app:open-paths', full);
}

/** One window for several files: the first opens with it, the rest arrive as tabs. */
function openInNewWindow(files: string[]): void {
  const win = createWindow(files[0]);
  if (files.length > 1) win.webContents.once('did-finish-load', () => sendPaths(win, files.slice(1)));
}

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv) => {
    const files = filesFromArgv(argv);
    // Files opened while Gridline runs become tabs in the window used last
    const win = BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0];
    if (!win) {
      if (files.length) openInNewWindow(files);
      else createWindow();
      return;
    }
    if (win.isMinimized()) win.restore();
    win.focus();
    if (files.length) sendPaths(win, files);
    const start = startFromArgv(argv);
    if (start.template || start.blank) win.webContents.send('app:start', start);
  });

  app.whenReady().then(() => {
    Menu.setApplicationMenu(null);
    app.setAppUserModelId('com.gridline.app');
    const files = filesFromArgv(process.argv);
    if (process.env.GRIDLINE_OPEN && existsSync(process.env.GRIDLINE_OPEN)) files.push(process.env.GRIDLINE_OPEN);
    const start = startFromArgv(process.argv);
    const envTemplate = process.env.GRIDLINE_TEMPLATE;
    const startTemplate = start.template ?? (envTemplate && /^[a-z-]{1,32}$/.test(envTemplate) ? envTemplate : undefined);
    if (files.length) openInNewWindow(files);
    // A blank workbook asked for (--new) skips the start page
    else createWindow(undefined, startTemplate, !start.blank);
    void registerSelf('gridline');
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
  });
}
