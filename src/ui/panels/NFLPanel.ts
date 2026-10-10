// "NFL data" panel: build NFL formulas by picking a player/team, seasons and stats.
import * as L from 'lucide';
import { cellKey, colToName, formatCell } from '../../core/address';
import { currentSeason, FIRST_SEASON } from '../../core/data/nfl';
import { STATS } from '../../core/data/nflStats';
import { TEAMS, teamName } from '../../core/data/nflTeams';
import { applyStyle, setCellInput } from '../../core/ops/cells';
import { setColumnWidth } from '../../core/ops/structure';
import type { App } from '../app';
import { argSuggestions } from '../argSuggest';
import { h, icon } from '../dom';
import { toast } from '../widgets';
import type { PanelContent } from './SidePanel';

type Mode = 'player' | 'team' | 'gamelog' | 'schedule' | 'standings' | 'leaders';
const MODES: Mode[] = ['player', 'team', 'gamelog', 'schedule', 'standings', 'leaders'];

const PLAYER_PRESETS: Record<string, string[]> = {
  Quarterback: ['Passing Yards', 'Passing TDs', 'Interceptions Thrown', 'Completion %', 'Passer Rating', 'Rushing Yards', 'Rushing TDs'],
  'Running back': ['Carries', 'Rushing Yards', 'Rushing TDs', 'Yards per Carry', 'Receptions', 'Receiving Yards', 'Receiving TDs'],
  Receiver: ['Targets', 'Receptions', 'Receiving Yards', 'Receiving TDs', 'Yards per Reception', 'Catch Rate'],
  Fantasy: ['Fantasy Points', 'Fantasy Points Half PPR', 'Fantasy Points PPR', 'PPR Points per Game', 'Games'],
  Defense: ['Tackles', 'Def Sacks', 'Tackles for Loss', 'QB Hits', 'Def Interceptions', 'Passes Defended', 'Forced Fumbles'],
  Kicker: ['FG Made', 'FG Attempts', 'FG %', 'Longest FG', 'XP Made', 'XP %'],
};

const TEAM_PRESETS: Record<string, string[]> = {
  Record: ['Record', 'Wins', 'Losses', 'Win %', 'Points For', 'Points Against', 'Point Differential'],
  Offense: ['Offensive Yards', 'Passing Yards', 'Rushing Yards', 'Passing TDs', 'Rushing TDs', 'Turnovers'],
  Defense: ['Def Sacks', 'Takeaways', 'Passing Yards Allowed', 'Rushing Yards Allowed', 'Turnover Differential'],
};

const quote = (s: string) => `"${s.replace(/"/g, '""')}"`;

export class NFLPanel implements PanelContent {
  title = 'NFL data';
  readonly el: HTMLElement;
  private mode: Mode;
  private player = 'Patrick Mahomes';
  private team: string;
  private first: number;
  private last = currentSeason();
  private stats = new Set<string>(['Passing Yards', 'Passing TDs', 'Rushing Yards', 'Rushing TDs']);
  private layout: 'table' | 'cells';
  private seasonType: string;
  private group = '';
  private leaderStat = 'Passing Yards';
  private count = 10;
  private position = '';
  private insertLabel: HTMLElement | null = null;

  constructor(
    private app: App,
    mode?: string,
  ) {
    const s = app.settings;
    this.mode = MODES.includes(mode as Mode) ? (mode as Mode) : 'player';
    this.team = s.str('nfl.team') || 'Chiefs';
    this.first = Math.max(FIRST_SEASON, this.last - Math.max(1, s.num('nfl.seasons')) + 1);
    this.layout = s.str('nfl.layout') === 'table' ? 'table' : 'cells';
    this.seasonType = s.str('nfl.seasonType') || 'REG';
    const fantasy = { ppr: 'Fantasy Points PPR', half: 'Fantasy Points Half PPR', standard: 'Fantasy Points' }[s.str('nfl.fantasy')] ?? 'Fantasy Points PPR';
    PLAYER_PRESETS.Fantasy = [fantasy, 'Fantasy Points', 'Fantasy Points Half PPR', 'Fantasy Points PPR', 'PPR Points per Game', 'Games'].filter((x, i, a) => a.indexOf(x) === i);
    if (this.mode === 'team') this.stats = new Set(TEAM_PRESETS.Record.slice(0, 4));
    this.el = h('div', { class: 'sp-body' });
    this.render();
  }

  /** Switches the panel to another kind of data (Player stats, Schedule…). */
  setTab(mode: string): void {
    if (!MODES.includes(mode as Mode)) return;
    this.mode = mode as Mode;
    if (mode === 'team' && ![...this.stats].some((x) => TEAM_PRESETS.Record.includes(x))) this.stats = new Set(TEAM_PRESETS.Record.slice(0, 4));
    this.render();
  }

  private render(): void {
    const b = this.el;
    b.innerHTML = '';
    const modes: [Mode, string][] = [
      ['player', 'Player stats'],
      ['team', 'Team stats'],
      ['gamelog', 'Game log'],
      ['schedule', 'Schedule'],
      ['standings', 'Standings'],
      ['leaders', 'Leaders'],
    ];
    const chips = h('div', { class: 'chips', style: { marginBottom: '12px' } });
    for (const [m, label] of modes) chips.append(h('button', { class: `chip${m === this.mode ? ' on' : ''}`, onclick: () => {
      this.mode = m;
      if (m === 'team' && ![...this.stats].some((x) => TEAM_PRESETS.Record.includes(x))) this.stats = new Set(TEAM_PRESETS.Record.slice(0, 4));
      if (m === 'player' && ![...this.stats].some((x) => STATS.some((d) => d.label === x && d.scope !== 'team'))) this.stats = new Set(PLAYER_PRESETS.Quarterback.slice(0, 4));
      this.render();
    } }, label));
    b.append(chips);

    if (this.mode === 'player' || this.mode === 'gamelog') b.append(this.playerField());
    if (this.mode === 'team' || this.mode === 'schedule') b.append(this.teamField());
    if (this.mode === 'player' || this.mode === 'team') {
      b.append(this.seasonRange());
      b.append(this.statPicker(this.mode === 'team' ? TEAM_PRESETS : PLAYER_PRESETS, this.mode === 'team' ? 'nflTeamStat' : 'nflPlayerStat'));
    }
    if (this.mode === 'player') {
      const lay = h('div', { class: 'segmented' },
        h('button', { class: `btn${this.layout === 'cells' ? ' on' : ''}`, onclick: () => {
          this.layout = 'cells';
          this.render();
        } }, 'Cells I can drag'),
        h('button', { class: `btn${this.layout === 'table' ? ' on' : ''}`, onclick: () => {
          this.layout = 'table';
          this.render();
        } }, 'One table formula'),
      );
      b.append(h('div', { class: 'field' }, h('label', null, 'Layout'), lay, h('div', { class: 'hint' }, this.layout === 'cells' ? 'A grid of NFL.PLAYER formulas: seasons down the side, stats across. Drag the last row down to add seasons.' : 'A single formula that returns the whole table.')));
    }
    if (this.mode === 'gamelog' || this.mode === 'schedule' || this.mode === 'standings' || this.mode === 'leaders') {
      const season = h('input', { class: 'input', type: 'number', min: String(FIRST_SEASON), max: String(currentSeason()), value: String(this.last), style: { width: '100px' } }) as HTMLInputElement;
      season.addEventListener('change', () => (this.last = Math.max(FIRST_SEASON, Math.min(currentSeason(), +season.value || currentSeason()))));
      b.append(h('div', { class: 'field' }, h('label', null, 'Season'), season));
    }
    if (this.mode === 'standings') {
      const sel = h('select', { class: 'input' }) as HTMLSelectElement;
      sel.append(h('option', { value: '' }, 'Whole league'));
      for (const g of argSuggestions('nflGroup', '')) sel.append(h('option', { value: g.value }, g.value));
      sel.value = this.group;
      sel.addEventListener('change', () => (this.group = sel.value));
      b.append(h('div', { class: 'field' }, h('label', null, 'Conference or division'), sel));
    }
    if (this.mode === 'leaders') {
      const stat = h('select', { class: 'input' }) as HTMLSelectElement;
      for (const d of STATS) if (d.scope !== 'team' && !d.text) stat.append(h('option', { value: d.label }, `${d.label} · ${d.group}`));
      stat.value = this.leaderStat;
      stat.addEventListener('change', () => (this.leaderStat = stat.value));
      const count = h('input', { class: 'input', type: 'number', min: '1', max: '100', value: String(this.count), style: { width: '80px' } }) as HTMLInputElement;
      count.addEventListener('change', () => (this.count = Math.max(1, Math.min(100, +count.value || 10))));
      const pos = h('select', { class: 'input' }) as HTMLSelectElement;
      for (const [v, label] of [['', 'Any position'], ['QB', 'QB'], ['RB', 'RB'], ['WR', 'WR'], ['TE', 'TE'], ['K', 'K'], ['DL', 'DL'], ['LB', 'LB'], ['DB', 'DB']]) pos.append(h('option', { value: v }, label));
      pos.value = this.position;
      pos.addEventListener('change', () => (this.position = pos.value));
      b.append(h('div', { class: 'field' }, h('label', null, 'Stat'), stat), h('div', { class: 'row' }, h('div', { class: 'field' }, h('label', null, 'How many'), count), h('div', { class: 'field' }, h('label', null, 'Position'), pos)));
    }
    if (this.mode !== 'standings' && this.mode !== 'schedule') {
      const st = h('select', { class: 'input' }) as HTMLSelectElement;
      for (const [v, label] of [['REG', 'Regular season'], ['POST', 'Playoffs'], ['ALL', 'Regular season + playoffs']]) st.append(h('option', { value: v }, label));
      st.value = this.seasonType;
      st.addEventListener('change', () => (this.seasonType = st.value));
      b.append(h('div', { class: 'field' }, h('label', null, 'Games'), st));
    }

    const a = this.app.sel.active;
    this.insertLabel = h('span', null, `Insert at ${formatCell(a.row, a.col)}`);
    const preview = h('div', { class: 'preview-box', style: { fontSize: '12px', wordBreak: 'break-all' } }, this.formulaPreview());
    b.append(
      h('div', { class: 'field', style: { marginTop: '8px' } }, h('label', null, 'Formula'), preview),
      h('div', { class: 'row', style: { justifyContent: 'flex-end' } }, h('button', { class: 'btn primary', onclick: () => this.insert() }, icon(L.Plus, 14), this.insertLabel)),
      h('p', { class: 'hint', style: { marginTop: '14px' } }, 'Data comes from nflverse (1999–today) and is saved on your computer after the first download, so it works offline. Use "Refresh live data" in the command palette to update.'),
    );
  }

  selectionChanged(): void {
    const a = this.app.sel.active;
    if (this.insertLabel) this.insertLabel.textContent = `Insert at ${formatCell(a.row, a.col)}`;
  }

  private playerField(): HTMLElement {
    const input = h('input', { class: 'input', value: this.player, placeholder: 'e.g. Josh Allen', list: 'nfl-player-list', spellcheck: false }) as HTMLInputElement;
    const list = h('datalist', { id: 'nfl-player-list' });
    for (const s of argSuggestions('nflPlayer', '')) list.append(h('option', { value: s.value }));
    input.addEventListener('input', () => {
      this.player = input.value;
      list.innerHTML = '';
      for (const s of argSuggestions('nflPlayer', input.value).slice(0, 30)) list.append(h('option', { value: s.value }));
    });
    input.addEventListener('change', () => this.render());
    return h('div', { class: 'field' }, h('label', null, 'Player'), input, list, h('div', { class: 'hint' }, 'Two players with the same name? Add the team: "Josh Allen (JAX)".'));
  }

  private teamField(): HTMLElement {
    const sel = h('select', { class: 'input' }) as HTMLSelectElement;
    for (const t of [...TEAMS].sort((x, y) => teamName(x.abbr).localeCompare(teamName(y.abbr)))) sel.append(h('option', { value: t.nickname }, `${teamName(t.abbr)} (${t.abbr})`));
    sel.value = this.team;
    sel.addEventListener('change', () => {
      this.team = sel.value;
      this.render();
    });
    return h('div', { class: 'field' }, h('label', null, 'Team'), sel);
  }

  private seasonRange(): HTMLElement {
    const mk = (v: number, set: (n: number) => void) => {
      const i = h('input', { class: 'input', type: 'number', min: String(FIRST_SEASON), max: String(currentSeason()), value: String(v), style: { width: '100px' } }) as HTMLInputElement;
      i.addEventListener('change', () => {
        set(Math.max(FIRST_SEASON, Math.min(currentSeason(), Math.round(+i.value) || currentSeason())));
        this.render();
      });
      return i;
    };
    return h('div', { class: 'field' }, h('label', null, 'Seasons'), h('div', { class: 'row' }, mk(this.first, (n) => (this.first = n)), h('span', { class: 'hint' }, 'to'), mk(this.last, (n) => (this.last = n))));
  }

  private statPicker(presets: Record<string, string[]>, kind: 'nflPlayerStat' | 'nflTeamStat'): HTMLElement {
    const wrap = h('div', { class: 'field' }, h('label', null, `Stats (${this.stats.size} selected)`));
    const presetRow = h('div', { class: 'chips', style: { marginBottom: '6px' } });
    for (const [name, list] of Object.entries(presets)) {
      presetRow.append(h('button', { class: 'chip', title: list.join(', '), onclick: () => {
        this.stats = new Set(list);
        this.render();
      } }, name));
    }
    const chosen = h('div', { class: 'chips' });
    for (const s of this.stats) {
      chosen.append(h('button', { class: 'chip on', title: 'Remove', onclick: () => {
        this.stats.delete(s);
        this.render();
      } }, s, ' ✕'));
    }
    const add = h('select', { class: 'input', style: { marginTop: '6px' } }) as HTMLSelectElement;
    add.append(h('option', { value: '' }, '+ Add a stat…'));
    let lastGroup = '';
    let og: HTMLOptGroupElement | null = null;
    for (const s of argSuggestions(kind, '').concat(kind === 'nflTeamStat' ? argSuggestions(kind, 'allowed') : [])) {
      if (this.stats.has(s.value)) continue;
      if (s.kind !== lastGroup) {
        og = h('optgroup', { label: s.kind }) as HTMLOptGroupElement;
        add.append(og);
        lastGroup = s.kind;
      }
      og!.append(h('option', { value: s.value }, s.value));
    }
    add.addEventListener('change', () => {
      if (add.value) this.stats.add(add.value);
      this.render();
    });
    wrap.append(h('div', { class: 'hint' }, 'Quick sets:'), presetRow, chosen, add);
    return wrap;
  }

  private typeArg(): string {
    return this.seasonType === 'REG' ? '' : `, ${quote(this.seasonType)}`;
  }

  private formulaPreview(): string {
    const stats = [...this.stats];
    switch (this.mode) {
      case 'player':
        if (this.layout === 'table') return `=NFL.PLAYERSTATS(${quote(this.player)}, ${quote(stats.join(', '))}, ${this.first}, ${this.last}${this.typeArg()})`;
        return `A grid of =NFL.PLAYER(player, stat, season) formulas — ${this.last - this.first + 1} seasons × ${stats.length} stats`;
      case 'team':
        return `A grid of =NFL.TEAM(${quote(this.team)}, stat, season) formulas — ${this.last - this.first + 1} seasons × ${stats.length} stats`;
      case 'gamelog':
        return `=NFL.GAMELOG(${quote(this.player)}, ${this.last}${this.seasonType === 'REG' ? '' : `, , ${quote(this.seasonType)}`})`;
      case 'schedule':
        return `=NFL.SCHEDULE(${quote(this.team)}, ${this.last})`;
      case 'standings':
        return `=NFL.STANDINGS(${this.last}${this.group ? `, ${quote(this.group)}` : ''})`;
      case 'leaders':
        return `=NFL.LEADERS(${quote(this.leaderStat)}, ${this.last}, ${this.count}${this.position ? `, ${quote(this.position)}` : this.seasonType !== 'REG' ? ', ""' : ''}${this.typeArg()})`;
    }
  }

  private insert(): void {
    const app = this.app;
    const s = app.sheet;
    const { row, col } = app.sel.active;
    const stats = [...this.stats];
    if ((this.mode === 'player' || this.mode === 'team') && !stats.length) {
      toast('Pick at least one stat.', { kind: 'error' });
      return;
    }
    if (this.mode === 'team' || (this.mode === 'player' && this.layout === 'cells')) {
      const isTeam = this.mode === 'team';
      const seasons = this.last - this.first + 1;
      const h1 = row + 1;
      // Make sure we don't overwrite data silently
      for (let r = row; r <= h1 + seasons; r++)
        for (let c = col; c <= col + stats.length; c++) {
          const cell = s.cells.get(cellKey(r, c));
          if (cell && (cell.f !== undefined || (cell.v !== null && cell.v !== undefined && cell.v !== ''))) {
            toast(`There's already data in ${formatCell(r, c)}. Pick an empty area first.`, { kind: 'error' });
            return;
          }
        }
      const nameRef = `$${colToName(col)}$${row + 1}`;
      const fn = isTeam ? 'NFL.TEAM' : 'NFL.PLAYER';
      app.transact(isTeam ? 'Insert NFL team stats' : 'Insert NFL player stats', (tx) => {
        setCellInput(tx, app.doc, s, row, col, isTeam ? this.team : this.player);
        setCellInput(tx, app.doc, s, h1, col, 'Season');
        stats.forEach((st, j) => setCellInput(tx, app.doc, s, h1, col + 1 + j, st));
        for (let i = 0; i < seasons; i++) {
          const r = h1 + 1 + i;
          // Later seasons count up from the row above, so dragging the last row down adds seasons
          setCellInput(tx, app.doc, s, r, col, i === 0 ? String(this.first) : `=${colToName(col)}${r}+1`);
          stats.forEach((_st, j) => {
            const c = col + 1 + j;
            const statRef = `${colToName(c)}$${h1 + 1}`;
            const seasonRef = `$${colToName(col)}${r + 1}`;
            setCellInput(tx, app.doc, s, r, c, `=${fn}(${nameRef}, ${statRef}, ${seasonRef}${this.typeArg() ? `, , ${quote(this.seasonType)}` : ''})`);
          });
        }
        applyStyle(tx, app.doc.wb, s, [{ r1: row, c1: col, r2: row, c2: col }], { bold: true, fontSize: 13 });
        applyStyle(tx, app.doc.wb, s, [{ r1: h1, c1: col, r2: h1, c2: col + stats.length }], { bold: true, borderBottom: { style: 'thin' } });
        // Widen stat columns so their names fit
        stats.forEach((st, j) => {
          const want = Math.round(st.length * 7.5 + 22);
          if (s.colWidth(col + 1 + j) < want) setColumnWidth(tx, s, [col + 1 + j], want);
        });
      });
      app.select({ r1: h1, c1: col, r2: h1 + seasons, c2: col + stats.length }, { row: h1, col });
      app.statusbar.flash('Inserted — change the name in the first cell to switch players, or drag the last row down for more seasons');
      return;
    }
    app.enterInput(s, row, col, this.formulaPreview(), false);
    app.focusGrid();
  }
}
