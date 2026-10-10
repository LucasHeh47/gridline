// Access to the desktop shell. Falls back to browser APIs when run outside Electron (development only).
import type { GridlineApi, OpenedFile, RecentFile, SaveKind, SaveResult } from '../shared/api';

declare global {
  interface Window {
    gridline?: GridlineApi;
  }
}

function pickFile(): Promise<OpenedFile | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xlsm,.csv,.tsv,.txt';
    input.onchange = async () => {
      const f = input.files?.[0];
      if (!f) return resolve(null);
      resolve({ path: null, name: f.name, data: new Uint8Array(await f.arrayBuffer()) });
    };
    input.click();
  });
}

const webFallback: Omit<GridlineApi, 'isDesktop'> & { isDesktop: boolean } = {
  platform: 'web',
  isDesktop: false,
  openFileDialog: pickFile,
  openPath: async () => null,
  readDroppedFile: async (file: File) => ({ path: null, name: file.name, data: new Uint8Array(await file.arrayBuffer()) }),
  saveFileDialog: async (name: string, _kind: SaveKind) => name,
  async saveFile(path: string, data: Uint8Array): Promise<SaveResult> {
    const blob = new Blob([data as BlobPart]);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = path.split(/[\\/]/).pop() || 'download';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    return { ok: true };
  },
  recentFiles: async (): Promise<RecentFile[]> => [],
  clearRecentFiles: async () => undefined,
  async readClipboard() {
    try {
      return { text: await navigator.clipboard.readText(), html: '' };
    } catch {
      return { text: '', html: '' };
    }
  },
  async writeClipboard(d) {
    try {
      if (d.html && 'ClipboardItem' in window) {
        await navigator.clipboard.write([
          new ClipboardItem({
            'text/plain': new Blob([d.text], { type: 'text/plain' }),
            'text/html': new Blob([d.html], { type: 'text/html' }),
          }),
        ]);
      } else await navigator.clipboard.writeText(d.text);
    } catch {
      // ignore
    }
  },
  setTitle: (t) => {
    document.title = t;
  },
  setDirty: () => undefined,
  setTheme: () => undefined,
  dataCacheSize: async () => 0,
  setZoomFactor: () => undefined,
  updates: {
    check: async () => ({ status: 'unsupported', current: __APP_VERSION__, error: 'Updates are available in the desktop app.' }),
    state: async () => ({ status: 'unsupported', current: __APP_VERSION__ }),
    download: async () => ({ status: 'unsupported', current: __APP_VERSION__ }),
    install: async () => false,
    openPage: async () => undefined,
    onState: () => undefined,
  },
  newWindow: (_p, t) => window.open(location.pathname + (t ? `?template=${encodeURIComponent(t)}` : ''), '_blank'),
  openExternal: (url) => {
    if (/^(https?:|mailto:)/i.test(url)) window.open(url, '_blank', 'noopener');
  },
  showInFolder: () => undefined,
  confirmClose: () => undefined,
  onRequestClose: () => undefined,
  toggleDevTools: () => undefined,
  appVersion: async () => __APP_VERSION__,
  async fetchData(url, opts) {
    try {
      const res = await fetch(url, { headers: opts.headers });
      if (!res.ok) return { ok: false, status: res.status, error: `HTTP ${res.status}` };
      return { ok: true, status: 200, text: await res.text() };
    } catch (err) {
      return { ok: false, status: 0, error: (err as Error).message };
    }
  },
  clearDataCache: async () => undefined,
};

export const platform: GridlineApi = (window.gridline ?? webFallback) as GridlineApi;
export const isDesktop = !!window.gridline;
