// Checks GitHub Releases for a newer Gridline and installs it when the user agrees.
// The installed (NSIS) app updates in place with electron-updater. The portable exe can't replace
// itself, so it only offers to open the download page.
import { app, BrowserWindow, ipcMain, net, shell } from 'electron';
import { autoUpdater, type ProgressInfo, type UpdateInfo } from 'electron-updater';
import type { UpdateCheckOptions, UpdateState } from '../src/shared/api';

/** Where releases are published (keep in sync with "publish" in package.json). */
export const RELEASES = { owner: 'LucasHeh47', repo: 'gridline' };
const RELEASES_PAGE = `https://github.com/${RELEASES.owner}/${RELEASES.repo}/releases`;

let state: UpdateState = { status: 'idle', current: app.getVersion() };
let lastCheck = 0;
let checking: Promise<UpdateState> | null = null;

const isPortable = () => !!process.env.PORTABLE_EXECUTABLE_DIR;

/** Local test feed (http://localhost only), e.g. GRIDLINE_UPDATE_FEED=http://127.0.0.1:8123/ */
function testFeed(): string | null {
  const f = process.env.GRIDLINE_UPDATE_FEED;
  return f && /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?\//.test(f) ? f : null;
}

function setState(next: Partial<UpdateState>): void {
  state = { ...state, ...next, current: app.getVersion() };
  for (const w of BrowserWindow.getAllWindows()) w.webContents.send('updates:state', state);
}

function notesText(info: UpdateInfo): string {
  const n = info.releaseNotes;
  const raw = Array.isArray(n) ? n.map((x) => x.note ?? '').join('\n') : n ?? '';
  // GitHub sends release notes as HTML
  return raw
    .replace(/<\/(p|li|h\d)>/gi, '\n')
    .replace(/<li>/gi, '• ')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/** Compares "1.2.3" style versions (pre-release suffixes sort before the release). */
export function compareVersions(a: string, b: string): number {
  const parse = (v: string) => {
    const [main, pre] = v.replace(/^v/, '').split('-', 2);
    return { nums: main.split('.').map((x) => parseInt(x, 10) || 0), pre: pre ?? '' };
  };
  const x = parse(a);
  const y = parse(b);
  for (let i = 0; i < 3; i++) if ((x.nums[i] ?? 0) !== (y.nums[i] ?? 0)) return (x.nums[i] ?? 0) - (y.nums[i] ?? 0);
  if (x.pre === y.pre) return 0;
  if (!x.pre) return 1;
  if (!y.pre) return -1;
  return x.pre < y.pre ? -1 : 1;
}

/** Portable builds: ask the GitHub API about the latest release. */
async function checkPortable(opts: UpdateCheckOptions): Promise<UpdateState> {
  const url = `https://api.github.com/repos/${RELEASES.owner}/${RELEASES.repo}/releases${opts.allowPrerelease ? '?per_page=10' : '/latest'}`;
  const res = await net.fetch(url, { headers: { accept: 'application/vnd.github+json', 'user-agent': 'Gridline' } });
  if (res.status === 404) return { ...state, status: 'not-available' };
  if (!res.ok) throw new Error(`GitHub answered HTTP ${res.status}`);
  const data = (await res.json()) as { tag_name: string; body?: string; html_url: string; draft?: boolean } | { tag_name: string; body?: string; html_url: string; draft?: boolean }[];
  const rel = Array.isArray(data) ? data.find((r) => !r.draft) : data;
  if (!rel) return { ...state, status: 'not-available' };
  const version = rel.tag_name.replace(/^v/, '');
  if (compareVersions(version, app.getVersion()) <= 0) return { ...state, status: 'not-available', version };
  return { ...state, status: 'available', version, notes: (rel.body ?? '').trim(), portable: true, url: rel.html_url };
}

async function check(opts: UpdateCheckOptions): Promise<UpdateState> {
  if (checking) return checking;
  // Starting several windows shouldn't check several times
  if (opts.silent && Date.now() - lastCheck < 10 * 60_000 && state.status !== 'idle') return state;
  const feed = testFeed();
  if (!app.isPackaged && !feed) {
    setState({ status: 'unsupported', error: 'Updates are available in the installed app.' });
    return state;
  }
  checking = (async () => {
    lastCheck = Date.now();
    setState({ status: 'checking', error: undefined });
    try {
      if (isPortable() && !feed) {
        setState(await checkPortable(opts));
        return state;
      }
      autoUpdater.allowPrerelease = !!opts.allowPrerelease;
      autoUpdater.autoDownload = !!opts.autoDownload;
      const r = await autoUpdater.checkForUpdates();
      if (!r || !r.isUpdateAvailable) setState({ status: 'not-available', version: r?.updateInfo.version });
      return state;
    } catch (err) {
      // No release published yet means there's nothing newer
      if (isNoRelease(err)) setState({ status: 'not-available', error: undefined });
      else setState({ status: 'error', error: friendlyError(err) });
      return state;
    } finally {
      checking = null;
    }
  })();
  return checking;
}

function errorText(err: unknown): string {
  const e = err as { message?: string; code?: string } | null;
  return `${e?.code ?? ''} ${e?.message ?? String(err)}`;
}

/** GitHub has no regular release yet (or none for this channel). */
function isNoRelease(err: unknown): boolean {
  return /ERR_UPDATER_LATEST_VERSION_NOT_FOUND|ERR_UPDATER_NO_PUBLISHED_VERSIONS|No published versions|Unable to find latest version/i.test(errorText(err));
}

/** A short, plain-English reason an update failed (the raw error is too technical to show). */
function friendlyError(err: unknown): string {
  const msg = errorText(err);
  if (/ENOTFOUND|ERR_INTERNET_DISCONNECTED|ERR_NAME_NOT_RESOLVED|ETIMEDOUT|ECONNRESET|ECONNREFUSED|ERR_NETWORK/i.test(msg)) return "Couldn't reach GitHub. Check your internet connection and try again.";
  if (/ERR_UPDATER_CHANNEL_FILE_NOT_FOUND|HttpError: 404/i.test(msg)) return 'The newest version is still being published. Try again in a few minutes.';
  if (/sha512|checksum/i.test(msg)) return 'The download was damaged. Try again.';
  if (/(HttpError:|HTTP) (403|429)|rate limit/i.test(msg)) return 'GitHub is limiting requests right now. Try again later.';
  if (/ENOSPC/i.test(msg)) return "There isn't enough disk space to download the update.";
  return 'Something went wrong while checking for updates. Try again later.';
}

export function setupUpdates(): void {
  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.logger = null;
  const feed = testFeed();
  if (feed) {
    autoUpdater.forceDevUpdateConfig = true;
    autoUpdater.setFeedURL({ provider: 'generic', url: feed });
  }
  autoUpdater.on('update-available', (info: UpdateInfo) => setState({ status: autoUpdater.autoDownload ? 'downloading' : 'available', version: info.version, notes: notesText(info), percent: 0, portable: false }));
  autoUpdater.on('update-not-available', (info: UpdateInfo) => setState({ status: 'not-available', version: info.version }));
  autoUpdater.on('download-progress', (p: ProgressInfo) => setState({ status: 'downloading', percent: Math.round(p.percent) }));
  autoUpdater.on('update-downloaded', (info: UpdateInfo) => setState({ status: 'downloaded', version: info.version, percent: 100 }));
  autoUpdater.on('error', (err: Error) => {
    if (state.status === 'downloading' || state.status === 'checking') setState({ status: 'error', error: friendlyError(err) });
  });

  ipcMain.handle('updates:check', (_e, opts: UpdateCheckOptions) => check({ silent: !!opts?.silent, allowPrerelease: !!opts?.allowPrerelease, autoDownload: !!opts?.autoDownload }));
  ipcMain.handle('updates:state', () => state);
  ipcMain.handle('updates:download', async () => {
    if (state.portable) {
      await shell.openExternal(state.url ?? RELEASES_PAGE);
      return state;
    }
    if (state.status !== 'available' && state.status !== 'error') return state;
    setState({ status: 'downloading', percent: 0, error: undefined });
    try {
      await autoUpdater.downloadUpdate();
    } catch (err) {
      setState({ status: 'error', error: friendlyError(err) });
    }
    return state;
  });
  ipcMain.handle('updates:install', () => {
    if (state.status !== 'downloaded') return false;
    // Quietly run the installer, then reopen Gridline
    setImmediate(() => autoUpdater.quitAndInstall(true, true));
    return true;
  });
  ipcMain.handle('updates:open-page', () => shell.openExternal(RELEASES_PAGE));
}
