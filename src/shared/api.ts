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
  pinned?: boolean;
}

export type SaveKind = 'xlsx' | 'xlsm' | 'csv' | 'tsv' | 'json';

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
  removeRecentFile(path: string): Promise<void>;
  pinRecentFile(path: string, pinned: boolean): Promise<void>;
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
  fetchData(url: string, opts: { maxAgeSec: number; headers?: Record<string, string>; offline?: boolean }): Promise<FetchResult>;
  clearDataCache(): Promise<void>;
  /** Size of the downloaded-data cache in bytes. */
  dataCacheSize(): Promise<number>;
  /** Scales the whole window (1 = 100%). */
  setZoomFactor(factor: number): void;
  updates: UpdatesApi;
}

export interface UpdateCheckOptions {
  /** A background check at startup (no message when there's nothing new). */
  silent?: boolean;
  allowPrerelease?: boolean;
  autoDownload?: boolean;
}

export interface UpdateState {
  status: 'idle' | 'checking' | 'available' | 'not-available' | 'downloading' | 'downloaded' | 'error' | 'unsupported';
  /** Version currently running. */
  current: string;
  /** Newest version found. */
  version?: string;
  notes?: string;
  percent?: number;
  error?: string;
  /** Portable builds can't update themselves; `url` is the download page. */
  portable?: boolean;
  url?: string;
}

export interface UpdatesApi {
  check(opts: UpdateCheckOptions): Promise<UpdateState>;
  state(): Promise<UpdateState>;
  download(): Promise<UpdateState>;
  /** Closes Gridline, installs the downloaded update and reopens. */
  install(): Promise<boolean>;
  openPage(): Promise<void>;
  onState(cb: (s: UpdateState) => void): void;
}

export interface FetchResult {
  ok: boolean;
  status: number;
  text?: string;
  error?: string;
  cached?: boolean;
  stale?: boolean;
}
