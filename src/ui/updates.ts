// Update prompts and "What's new". The desktop shell does the checking and installing.
import changelog from '../../CHANGELOG.md?raw';
import type { UpdateState } from '../shared/api';
import type { App } from './app';
import { h } from './dom';
import { isDesktop, platform } from './platform';
import { dialog, toast } from './widgets';

const SKIP_KEY = 'gridline.skipVersion';
const SEEN_KEY = 'gridline.lastVersion';

/** Release notes for one version from CHANGELOG.md, as plain text lines. */
export function notesFor(version: string): string[] {
  const parts = changelog.split(/^## /m);
  const sec = parts.find((p) => p.startsWith(version + '\n') || p.startsWith(version + '\r'));
  if (!sec) return [];
  return sec
    .split(/\r?\n/)
    .slice(1)
    .map((l) => l.trim())
    .filter(Boolean);
}

export class Updates {
  private state: UpdateState = { status: 'idle', current: '' };
  /** The user asked (menu or settings), so report every outcome. */
  private manual = false;
  private prompted = new Set<string>();

  constructor(private app: App) {
    platform.updates.onState((s) => this.onState(s));
  }

  /** Called once the window is ready. */
  async startup(): Promise<void> {
    this.started = true;
    const current = await platform.appVersion();
    let last: string | null = null;
    try {
      last = localStorage.getItem(SEEN_KEY);
      localStorage.setItem(SEEN_KEY, current);
    } catch {
      // storage unavailable
    }
    // People updating from 1.0.0 have no record of the last version, but do have saved preferences
    const usedBefore = (() => {
      try {
        return !!localStorage.getItem('gridline.prefs');
      } catch {
        return false;
      }
    })();
    const updated = last ? last !== current : usedBefore;
    if (updated && this.app.settings.bool('updates.showNotes') && notesFor(current).length) setTimeout(() => this.showWhatsNew(), 600);
    if (isDesktop && this.app.settings.bool('updates.onStartup')) setTimeout(() => void this.check(true), 4000);
  }

  /** Turning on the Developer branch checks right away. */
  devBranchChanged(on: boolean): void {
    if (!this.started) return;
    if (on) {
      toast("Developer branch is on. Checking for a development build…");
      void this.check(false);
    } else toast("Developer branch is off. You'll stay on this version until the next regular release.");
  }

  private started = false;

  async check(silent: boolean): Promise<void> {
    this.manual = !silent;
    if (!silent) this.app.statusbar.flash('Checking for updates…');
    const s = await platform.updates.check({
      silent,
      allowPrerelease: this.app.settings.bool('updates.devBranch'),
      autoDownload: this.app.settings.bool('updates.autoDownload'),
    });
    this.onState(s);
  }

  private onState(s: UpdateState): void {
    const prev = this.state.status;
    this.state = s;
    const manual = this.manual;
    switch (s.status) {
      case 'not-available':
        if (manual && prev === 'checking') toast(`You're up to date — Gridline ${s.current} is the newest version.`, { kind: 'success' });
        if (manual) this.manual = false;
        break;
      case 'unsupported':
        if (manual) toast(s.error ?? 'Updates are available in the installed app.');
        this.manual = false;
        break;
      case 'error':
        if (manual || prev === 'downloading') toast(`Couldn't update: ${s.error ?? 'unknown error'}`, { kind: 'error' });
        this.manual = false;
        break;
      case 'available':
        if (!s.version) break;
        if (!manual && this.skipped() === s.version) break;
        if (this.prompted.has(s.version) && !manual) break;
        this.prompted.add(s.version);
        this.manual = false;
        void this.promptAvailable(s);
        break;
      case 'downloading':
        this.app.statusbar.setProgress(`Downloading Gridline ${s.version ?? ''}…`, s.percent ?? 0);
        break;
      case 'downloaded':
        this.app.statusbar.setProgress(null, 100);
        void this.promptInstall(s);
        break;
    }
  }

  private skipped(): string | null {
    try {
      return localStorage.getItem(SKIP_KEY);
    } catch {
      return null;
    }
  }

  private notesEl(text: string | undefined, version: string): HTMLElement {
    const lines = text?.trim() ? text.trim().split(/\r?\n/) : notesFor(version);
    const box = h('div', { class: 'update-notes' });
    if (!lines.length) box.append(h('p', null, 'Bug fixes and improvements.'));
    for (const l of lines) box.append(h('p', null, l.replace(/^[-*]\s*/, '• ').replace(/\*\*/g, '')));
    return box;
  }

  private async promptAvailable(s: UpdateState): Promise<void> {
    const body = h(
      'div',
      null,
      h('p', null, `Gridline ${s.version} is available. You have ${s.current}.`),
      h('div', { class: 'update-notes-title' }, "What's new"),
      this.notesEl(s.notes, s.version!),
    );
    const choice = await dialog<'now' | 'later' | 'skip'>(
      'Update available',
      body,
      [
        { label: s.portable ? 'Download' : 'Update now', value: 'now', primary: true },
        { label: 'Skip this version', value: 'skip' },
        { label: 'Later', value: 'later' },
      ],
      { cancelValue: 'later' },
    );
    if (choice === 'skip') {
      try {
        localStorage.setItem(SKIP_KEY, s.version!);
      } catch {
        // ignore
      }
      toast(`You won't be reminded about ${s.version}. Check again any time from the Gridline menu.`);
    } else if (choice === 'now') {
      if (s.portable) toast('Opening the download page. This portable copy can be replaced with the new one.');
      await platform.updates.download();
    }
  }

  private async promptInstall(s: UpdateState): Promise<void> {
    const choice = await dialog<'now' | 'later'>(
      'Ready to update',
      `Gridline ${s.version} has been downloaded. Gridline will close, update and reopen — it only takes a moment.`,
      [
        { label: 'Restart and update', value: 'now', primary: true },
        { label: 'When I close Gridline', value: 'later' },
      ],
      { cancelValue: 'later' },
    );
    if (choice !== 'now') {
      toast('The update will be installed the next time you close Gridline.');
      return;
    }
    if (!(await this.app.saveBeforeQuit())) return;
    await platform.updates.install();
  }

  async showWhatsNew(): Promise<void> {
    const v = await platform.appVersion();
    const lines = notesFor(v);
    const body = h('div', null, h('p', null, `Here's what changed in Gridline ${v}.`), this.notesEl(lines.join('\n'), v));
    await dialog(`What's new in ${v}`, body, [{ label: 'Got it', value: true, primary: true }], { wide: true });
  }
}
