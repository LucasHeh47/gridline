// The bridge between the Electron main process and the UI (exposed as window.gridline).

export interface OpenedFile {
  path: string | null;
  name: string;
  data: Uint8Array;
}

export interface RecentFile {
  path: string;
  name: string;
  openedAt: number;
}

export type SaveKind = 'xlsx' | 'csv' | 'tsv';

export interface SaveResult {
  ok: boolean;
  error?: string;
}

export interface GridlineApi {
  platform: string;
  isDesktop: true;
  openFileDialog(): Promise<OpenedFile | null>;
  openPath(path: string): Promise<OpenedFile | null>;
  readDroppedFile(file: File): Promise<OpenedFile>;
  saveFileDialog(defaultName: string, kind: SaveKind): Promise<string | null>;
  saveFile(path: string, data: Uint8Array): Promise<SaveResult>;
  recentFiles(): Promise<RecentFile[]>;
  clearRecentFiles(): Promise<void>;
  readClipboard(): Promise<{ text: string; html: string }>;
  writeClipboard(data: { text: string; html?: string }): Promise<void>;
  setTitle(title: string): void;
  setDirty(dirty: boolean): void;
  setTheme(theme: { dark: boolean; bar: string; symbol: string }): void;
  /** Opens another window, optionally with a file or a starter template. */
  newWindow(path?: string, template?: string): void;
  openExternal(url: string): void;
  showInFolder(path: string): void;
  confirmClose(): void;
  onRequestClose(cb: () => void): void;
  toggleDevTools(): void;
  appVersion(): Promise<string>;
  /** Downloads live data through the desktop shell (allow-listed sites, cached on disk). */
  fetchData(url: string, opts: { maxAgeSec: number; headers?: Record<string, string> }): Promise<FetchResult>;
  clearDataCache(): Promise<void>;
}

export interface FetchResult {
  ok: boolean;
  status: number;
  text?: string;
  error?: string;
  cached?: boolean;
  stale?: boolean;
}
