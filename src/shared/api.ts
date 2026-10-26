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
  /** Files opened from outside (Explorer, a second launch) for this window to open in tabs. */
  onOpenPaths(cb: (paths: string[]) => void): void;
  /** A blank workbook or a preset asked for from outside (Baseline). */
  onStart(cb: (start: { template?: string; blank?: boolean }) => void): void;
  /** The other apps of the suite (Baseline, Pixelline). */
  suite: SuiteApi;
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
  ai: AiApi;
}

// ---------- AI assistant ----------

/** 'claude': Anthropic's API with the user's own key. 'local': an OpenAI-compatible server on the PC (Ollama, LM Studio). */
export type AiProvider = 'claude' | 'local';

export interface AiToolDef {
  name: string;
  description: string;
  /** JSON Schema of the tool's input. */
  input_schema: Record<string, unknown>;
}

export type AiBlock =
  | { type: 'text'; text: string }
  | { type: 'tool_use'; id: string; name: string; input: unknown }
  | { type: 'tool_result'; tool_use_id: string; content: string; is_error?: boolean };

export interface AiMessage {
  role: 'user' | 'assistant';
  content: AiBlock[];
}

export interface AiRequest {
  provider: AiProvider;
  model: string;
  system: string;
  messages: AiMessage[];
  tools: AiToolDef[];
  maxTokens?: number;
  /** Local provider: the server address, e.g. http://localhost:11434/v1 */
  baseUrl?: string;
  /** Local provider: how many tokens the model may read at once (Ollama's num_ctx). */
  contextSize?: number;
}

export type AiEvent =
  | { type: 'text'; text: string }
  | { type: 'tool_use'; id: string; name: string; input: unknown }
  | { type: 'done'; stopReason: string; usage?: { input: number; output: number }; warning?: string }
  | { type: 'error'; message: string };

export interface AiApi {
  /** Starts a request; its events arrive through onEvent with the same id. Resolves when it ends. */
  chat(id: string, req: AiRequest): Promise<void>;
  cancel(id: string): void;
  onEvent(cb: (id: string, e: AiEvent) => void): void;
  /** Stores the Claude API key encrypted on this computer (null removes it). The UI never reads it back. */
  setKey(key: string | null): Promise<{ ok: boolean; error?: string }>;
  hasKey(): Promise<boolean>;
  /** Models offered by a local server. */
  localModels(baseUrl: string): Promise<{ ok: boolean; models: string[]; error?: string }>;
}

export type SuiteAppId = 'baseline' | 'gridline' | 'pixelline' | 'frameline';

export interface SuiteApp {
  id: SuiteAppId;
  name: string;
  description: string;
  installed: boolean;
  exe: string | null;
  version: string | null;
  download: string;
}

export interface SuiteApi {
  apps(): Promise<SuiteApp[]>;
  /** Starts an app, optionally with files to open or options like "--new=16x16". */
  launch(id: SuiteAppId, args?: string[]): Promise<boolean>;
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
