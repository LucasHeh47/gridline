// Exposes a minimal, typed API to the UI. The UI never gets direct Node or file-system access.
import { contextBridge, ipcRenderer, webFrame, webUtils } from 'electron';
import type { GridlineApi, OpenedFile, SaveKind } from '../src/shared/api';

const api: GridlineApi = {
  platform: process.platform,
  isDesktop: true,
  openFileDialog: () => ipcRenderer.invoke('file:open-dialog'),
  openPath: (p: string) => ipcRenderer.invoke('file:open-path', p),
  async readDroppedFile(file: File): Promise<OpenedFile> {
    const p = webUtils.getPathForFile(file);
    const granted: string | null = p ? await ipcRenderer.invoke('file:grant', p) : null;
    const data = new Uint8Array(await file.arrayBuffer());
    return { path: granted, name: file.name, data };
  },
  saveFileDialog: (defaultName: string, kind: SaveKind) => ipcRenderer.invoke('file:save-dialog', defaultName, kind),
  saveFile: (p: string, data: Uint8Array) => ipcRenderer.invoke('file:save', p, data),
  recentFiles: () => ipcRenderer.invoke('file:recent'),
  clearRecentFiles: () => ipcRenderer.invoke('file:clear-recent'),
  removeRecentFile: (p) => ipcRenderer.invoke('file:remove-recent', p),
  pinRecentFile: (p, pinned) => ipcRenderer.invoke('file:pin-recent', p, pinned),
  readClipboard: () => ipcRenderer.invoke('clipboard:read'),
  writeClipboard: (d) => ipcRenderer.invoke('clipboard:write', d),
  setTitle: (t) => ipcRenderer.send('win:set-title', t),
  setDirty: (d) => ipcRenderer.send('win:set-dirty', d),
  setTheme: (t) => ipcRenderer.send('win:set-theme', t),
  newWindow: (p, t) => ipcRenderer.send('win:new', p, t),
  openExternal: (url) => ipcRenderer.send('shell:open-external', url),
  showInFolder: (p) => ipcRenderer.send('shell:show-in-folder', p),
  confirmClose: () => ipcRenderer.send('win:confirm-close'),
  onRequestClose: (cb) => {
    ipcRenderer.on('app:request-close', () => cb());
  },
  onOpenPaths: (cb) => {
    ipcRenderer.on('app:open-paths', (_e, paths: string[]) => cb(Array.isArray(paths) ? paths.filter((p) => typeof p === 'string') : []));
  },
  onStart: (cb) => {
    ipcRenderer.on('app:start', (_e, start: { template?: string; blank?: boolean }) => cb(start && typeof start === 'object' ? start : {}));
  },
  suite: {
    apps: () => ipcRenderer.invoke('suite:apps'),
    launch: (id, args) => ipcRenderer.invoke('suite:launch', id, args ?? []),
  },
  toggleDevTools: () => ipcRenderer.send('win:devtools'),
  appVersion: () => ipcRenderer.invoke('app:version'),
  fetchData: (url, opts) => ipcRenderer.invoke('net:fetch', url, opts),
  clearDataCache: () => ipcRenderer.invoke('net:clear-cache'),
  dataCacheSize: () => ipcRenderer.invoke('net:cache-size'),
  setZoomFactor: (f) => webFrame.setZoomFactor(Math.min(2, Math.max(0.5, Number(f) || 1))),
  ai: {
    chat: (id, req) => ipcRenderer.invoke('ai:chat', id, req),
    cancel: (id) => ipcRenderer.send('ai:cancel', id),
    onEvent: (cb) => {
      ipcRenderer.on('ai:event', (_e, id, ev) => cb(id, ev));
    },
    setKey: (key) => ipcRenderer.invoke('ai:set-key', key),
    hasKey: () => ipcRenderer.invoke('ai:has-key'),
    localModels: (url) => ipcRenderer.invoke('ai:local-models', url),
  },
  updates: {
    check: (opts) => ipcRenderer.invoke('updates:check', opts),
    state: () => ipcRenderer.invoke('updates:state'),
    download: () => ipcRenderer.invoke('updates:download'),
    install: () => ipcRenderer.invoke('updates:install'),
    openPage: () => ipcRenderer.invoke('updates:open-page'),
    onState: (cb) => {
      ipcRenderer.on('updates:state', (_e, s) => cb(s));
    },
  },
};

contextBridge.exposeInMainWorld('gridline', api);
