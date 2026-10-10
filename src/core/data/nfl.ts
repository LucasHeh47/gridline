// NFL data from nflverse (https://github.com/nflverse): player & team stats, games, players.
import { config } from '../config';
import { Errors, ErrorValue, type Scalar } from '../values';
import { external, type Got, type Resource } from './external';
import { codeInSeason, franchiseOf, teamInfo, teamName } from './nflTeams';
import { AVG_COLS, MAX_COLS, type ResolvedStat, type StatContext } from './nflStats';
import { Table } from './table';

const RELEASES = 'https://github.com/nflverse/nflverse-data/releases/download';
const GAMES_URL = 'https://github.com/nflverse/nfldata/raw/master/data/games.csv';
export const FIRST_SEASON = 1999;
const HOUR = 3600;
const DAY = 86400;

/** The season in progress (or just finished): September-December -> this year, January-August -> last year. */
export function currentSeason(now = new Date()): number {
  return now.getMonth() >= 8 ? now.getFullYear() : now.getFullYear() - 1;
}

function maxAgeFor(season: number): number {
  const cur = currentSeason();
  if (season >= cur) return config.nflCurrentSeasonHours * HOUR;
  if (season === cur - 1) return 2 * DAY;
  return 30 * DAY;
}

export type SeasonType = 'REG' | 'POST' | 'ALL';

export function parseSeasonType(v: Scalar | undefined): SeasonType | ErrorValue {
  if (v === undefined || v === null || v === '') return 'REG';
  if (v instanceof ErrorValue) return v;
  const s = String(v).trim().toLowerCase();
  if (['reg', 'regular', 'regular season', 'season'].includes(s)) return 'REG';
  if (['post', 'playoffs', 'playoff', 'postseason'].includes(s)) return 'POST';
  if (['all', 'total', 'regpost', 'both', 'reg+post'].includes(s)) return 'ALL';
  return Errors.value(`season_type should be "REG" (regular season), "POST" (playoffs) or "ALL".`);
}

export function parseSeason(v: Scalar | undefined): number | ErrorValue {
  if (v === undefined || v === null || v === '') return currentSeason();
  if (v instanceof ErrorValue) return v;
  const n = typeof v === 'number' ? v : Number(String(v).trim());
  if (!Number.isInteger(n)) return Errors.value(`"${v}" isn't a season. Use a year like 2023.`);
  if (n < FIRST_SEASON) return Errors.na(`NFL stats are available from ${FIRST_SEASON} onward.`);
  if (n > currentSeason()) return Errors.na(`The ${n} season hasn't started yet.`);
  return n;
}

/** Week number, or a playoff round name ("WC", "Divisional", "Conference", "Super Bowl"). */
export function parseWeek(v: Scalar | undefined, season: number): number | null | ErrorValue {
  if (v === undefined || v === null || v === '') return null;
  if (v instanceof ErrorValue) return v;
  const regWeeks = season >= 2021 ? 18 : 17;
  if (typeof v === 'string') {
    const s = v.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    const rounds: Record<string, number> = { wc: 1, wildcard: 1, div: 2, divisional: 2, con: 3, conf: 3, conference: 3, championship: 3, sb: 4, superbowl: 4 };
    if (rounds[s]) return regWeeks + rounds[s];
    const n = Number(s.replace(/^week/, ''));
    if (Number.isInteger(n)) v = n;
    else return Errors.value(`"${v}" isn't a week. Use 1-${regWeeks}, or "WC", "DIV", "CON", "SB" for the playoffs.`);
  }
  if (typeof v !== 'number' || !Number.isInteger(v) || v < 1 || v > regWeeks + 4) {
    return Errors.value(`Week must be between 1 and ${regWeeks + 4} in ${season}.`);
  }
  return v;
}

// ---------- names ----------

export function normName(s: string): string {
  return s
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[.'’`]/g, ' ')
    .replace(/-/g, ' ')
    .replace(/\b(jr|sr|ii|iii|iv|v)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** "Josh Allen (BUF)" -> name + qualifier (team or position). */
export function parsePlayerQuery(input: string): { name: string; qualifier?: string } {
  const m = /^(.*?)\s*[([]([^)\]]+)[)\]]\s*$/.exec(input.trim());
  if (m) return { name: m[1], qualifier: m[2].trim() };
  return { name: input.trim() };
}

function lev(a: string, b: string): number {
  const d: number[] = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    let prev = d[0];
    d[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = d[j];
      d[j] = Math.min(d[j] + 1, d[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = tmp;
    }
  }
  return d[b.length];
}

// ---------- tables ----------

/** A player stats file indexed by name and id. */
export class PlayerTable extends Table {
  private byName = new Map<string, number[]>();
  private byLast = new Map<string, number[]>();
  private byId = new Map<string, number[]>();

  constructor(text: string) {
    super(text);
    const nameCol = this.col('player_display_name');
    const shortCol = this.col('player_name');
    const idCol = this.col('player_id');
    const add = (map: Map<string, number[]>, k: string, i: number) => {
      if (!k) return;
      const list = map.get(k);
      if (list) {
        if (list[list.length - 1] !== i) list.push(i);
      } else map.set(k, [i]);
    };
    this.rows.forEach((r, i) => {
      const full = normName(r[nameCol] ?? '');
      add(this.byName, full, i);
      const short = normName(r[shortCol] ?? '');
      add(this.byName, short, i);
      add(this.byName, short.replace(/ /g, ''), i);
      const parts = full.split(' ');
      if (parts.length > 1) add(this.byLast, parts[parts.length - 1], i);
      add(this.byId, r[idCol] ?? '', i);
    });
  }

  rowsForId(id: string): string[][] {
    return (this.byId.get(id) ?? []).map((i) => this.rows[i]);
  }

  /** Rows matching a name (all players with that name). */
  candidates(name: string): number[] {
    const raw = name.trim();
    if (/^\d\d-\d{7}$/.test(raw)) return this.byId.get(raw) ?? [];
    const n = normName(raw);
    return this.byName.get(n) ?? this.byName.get(n.replace(/ /g, '')) ?? (n.includes(' ') ? [] : this.byLast.get(n) ?? []);
  }

  suggest(name: string): string | null {
    const n = normName(name);
    let best: string | null = null;
    let bestD = Infinity;
    const nameCol = this.col('player_display_name');
    const seen = new Set<string>();
    for (const r of this.rows) {
      const disp = r[nameCol];
      if (!disp || seen.has(disp)) continue;
      seen.add(disp);
      const d = lev(n, normName(disp));
      if (d < bestD) {
        bestD = d;
        best = disp;
      }
    }
    return bestD <= Math.max(2, Math.floor(n.length / 4)) ? best : null;
  }

  /** Every distinct display name (for autocomplete). */
  names(): string[] {
    const nameCol = this.col('player_display_name');
    return [...new Set(this.rows.map((r) => r[nameCol]).filter(Boolean))];
  }
}

export class Games {
  readonly list: Game[] = [];
  private bySeason = new Map<number, Game[]>();

  constructor(text: string) {
    const t = new Table(text);
    const n = (r: string[], c: string) => t.num(r, c);
    const s = (r: string[], c: string) => t.str(r, c);
    for (const r of t.rows) {
      const g: Game = {
        id: s(r, 'game_id'),
        season: n(r, 'season') ?? 0,
        type: s(r, 'game_type'),
        week: n(r, 'week') ?? 0,
        date: s(r, 'gameday'),
        weekday: s(r, 'weekday'),
        time: s(r, 'gametime'),
        awayCode: s(r, 'away_team'),
        homeCode: s(r, 'home_team'),
        away: franchiseOf(s(r, 'away_team')),
        home: franchiseOf(s(r, 'home_team')),
        awayScore: n(r, 'away_score'),
        homeScore: n(r, 'home_score'),
        location: s(r, 'location'),
        overtime: n(r, 'overtime') === 1,
        spread: n(r, 'spread_line'),
        totalLine: n(r, 'total_line'),
        awayML: n(r, 'away_moneyline'),
        homeML: n(r, 'home_moneyline'),
        roof: s(r, 'roof'),
        surface: s(r, 'surface'),
        temp: n(r, 'temp'),
        wind: n(r, 'wind'),
        awayQB: s(r, 'away_qb_name'),
        homeQB: s(r, 'home_qb_name'),
        awayCoach: s(r, 'away_coach'),
        homeCoach: s(r, 'home_coach'),
        stadium: s(r, 'stadium'),
        divGame: n(r, 'div_game') === 1,
      };
      this.list.push(g);
      const list = this.bySeason.get(g.season) ?? [];
      list.push(g);
      this.bySeason.set(g.season, list);
    }
  }

  season(season: number): Game[] {
    return this.bySeason.get(season) ?? [];
  }

  latestSeason(): number {
    return Math.max(...this.bySeason.keys());
  }
}

export interface Game {
  id: string;
  season: number;
  type: string;
  week: number;
  date: string;
  weekday: string;
  time: string;
  away: string;
  home: string;
  awayCode: string;
  homeCode: string;
  awayScore: number | null;
  homeScore: number | null;
  location: string;
  overtime: boolean;
  spread: number | null;
  totalLine: number | null;
  awayML: number | null;
  homeML: number | null;
  roof: string;
  surface: string;
  temp: number | null;
  wind: number | null;
  awayQB: string;
  homeQB: string;
  awayCoach: string;
  homeCoach: string;
  stadium: string;
  divGame: boolean;
}

export class PlayersTable extends Table {
  private byName = new Map<string, number[]>();
  private byId = new Map<string, number>();

  constructor(text: string) {
    super(text);
    const nameCol = this.col('display_name');
    const idCol = this.col('gsis_id');
    this.rows.forEach((r, i) => {
      const k = normName(r[nameCol] ?? '');
      if (k) {
        const list = this.byName.get(k) ?? [];
        list.push(i);
        this.byName.set(k, list);
      }
      if (r[idCol]) this.byId.set(r[idCol], i);
    });
  }

  /** Best match for a name: active players first, then the most recent. */
  find(query: string): string[] | null {
    const q = parsePlayerQuery(query);
    if (/^\d\d-\d{7}$/.test(q.name.trim())) {
      const i = this.byId.get(q.name.trim());
      return i === undefined ? null : this.rows[i];
    }
    let idx = this.byName.get(normName(q.name)) ?? [];
    if (q.qualifier) {
      const qu = q.qualifier.toUpperCase();
      const filtered = idx.filter((i) => {
        const r = this.rows[i];
        return franchiseOf(this.str(r, 'latest_team')) === franchiseOf(qu) || this.str(r, 'position').toUpperCase() === qu || this.str(r, 'position_group').toUpperCase() === qu;
      });
      if (filtered.length) idx = filtered;
    }
    if (!idx.length) return null;
    const score = (i: number) => {
      const r = this.rows[i];
      return (this.str(r, 'status') === 'ACT' ? 10000 : 0) + (this.num(r, 'last_season') ?? 0);
    };
    idx = [...idx].sort((a, b) => score(b) - score(a));
    return this.rows[idx[0]];
  }

  byGsis(id: string): string[] | null {
    const i = this.byId.get(id);
    return i === undefined ? null : this.rows[i];
  }
}

// ---------- resources ----------

function resource<T>(key: string, url: string, maxAgeSec: number, label: string, parse: (t: string) => T): Resource<T> {
  return { key, url, maxAgeSec, label, parse };
}

const FILE_TYPE: Record<SeasonType, string> = { REG: 'reg', POST: 'post', ALL: 'regpost' };
const TYPE_LABEL: Record<SeasonType, string> = { REG: '', POST: ' playoff', ALL: ' (regular + playoffs)' };

export function playerSeasonRes(season: number, type: SeasonType): Resource<PlayerTable> {
  return resource(
    `nfl:player:${type}:${season}`,
    `${RELEASES}/stats_player/stats_player_${FILE_TYPE[type]}_${season}.csv`,
    maxAgeFor(season),
    `NFL ${season}${TYPE_LABEL[type]} player stats`,
    (t) => new PlayerTable(t),
  );
}

export function playerWeeksRes(season: number): Resource<PlayerTable> {
  return resource(`nfl:playerweek:${season}`, `${RELEASES}/stats_player/stats_player_week_${season}.csv`, maxAgeFor(season), `NFL ${season} weekly player stats`, (t) => new PlayerTable(t));
}

export function teamSeasonRes(season: number, type: SeasonType): Resource<Table> {
  return resource(
    `nfl:team:${type}:${season}`,
    `${RELEASES}/stats_team/stats_team_${FILE_TYPE[type]}_${season}.csv`,
    maxAgeFor(season),
    `NFL ${season}${TYPE_LABEL[type]} team stats`,
    (t) => new Table(t),
  );
}

export function teamWeeksRes(season: number): Resource<Table> {
  return resource(`nfl:teamweek:${season}`, `${RELEASES}/stats_team/stats_team_week_${season}.csv`, maxAgeFor(season), `NFL ${season} weekly team stats`, (t) => new Table(t));
}

export const gamesRes: Resource<Games> = {
  ...resource('nfl:games', GAMES_URL, 0, 'NFL game results', (t) => new Games(t)),
  get maxAgeSec() {
    return config.nflCurrentSeasonHours * HOUR;
  },
};
export const playersRes: Resource<PlayersTable> = resource('nfl:players', `${RELEASES}/players/players.csv`, DAY, 'NFL player profiles', (t) => new PlayersTable(t));

export const playerSeason = (season: number, type: SeasonType): Got<PlayerTable> => external.get(playerSeasonRes(season, type));
export const playerWeeks = (season: number): Got<PlayerTable> => external.get(playerWeeksRes(season));
export const teamSeason = (season: number, type: SeasonType): Got<Table> => external.get(teamSeasonRes(season, type));
export const teamWeeks = (season: number): Got<Table> => external.get(teamWeeksRes(season));
export const games = (): Got<Games> => external.get(gamesRes);
export const players = (): Got<PlayersTable> => external.get(playersRes);

// ---------- players ----------

export interface FoundPlayer {
  id: string;
  rows: string[][];
}

/** Finds a player in a stats table. Returns their rows (one per season or per game). */
export function findPlayer(pt: PlayerTable, query: string, context: string): FoundPlayer | ErrorValue {
  const q = parsePlayerQuery(query);
  if (!q.name) return Errors.value('Enter a player name, like "Patrick Mahomes".');
  let idx = pt.candidates(q.name);
  const idCol = pt.col('player_id');
  if (q.qualifier && idx.length) {
    const qu = q.qualifier.toUpperCase();
    const filtered = idx.filter((i) => {
      const r = pt.rows[i];
      const team = pt.str(r, 'recent_team') || pt.str(r, 'team');
      return franchiseOf(team) === franchiseOf(qu) || pt.str(r, 'position').toUpperCase() === qu || pt.str(r, 'position_group').toUpperCase() === qu;
    });
    if (filtered.length) idx = filtered;
  }
  if (!idx.length) {
    const sug = pt.suggest(q.name);
    return Errors.na(`No ${context} for "${q.name}".${sug ? ` Did you mean ${sug}?` : ' Check the spelling, or they may not have played then.'}`);
  }
  // Several players can share a name: pick the most productive (use "Name (TEAM)" to choose).
  const totals = new Map<string, number>();
  for (const i of idx) {
    const r = pt.rows[i];
    const id = r[idCol];
    totals.set(id, (totals.get(id) ?? 0) + (pt.num(r, 'fantasy_points_ppr') ?? 0) + (pt.num(r, 'games') ?? 1) * 0.01);
  }
  let bestId = '';
  let best = -Infinity;
  for (const [id, v] of totals) {
    if (v > best) {
      best = v;
      bestId = id;
    }
  }
  return { id: bestId, rows: pt.rowsForId(bestId) };
}

/** Reads a stat from one or more rows (summing counting stats). */
export function statFromRows(t: Table, rows: string[][], stat: ResolvedStat, gamesOverride?: number): Scalar {
  const def = stat.def;
  if (!rows.length) return '';
  if (def.text) {
    const r = rows[rows.length - 1];
    switch (def.text) {
      case 'team':
        return t.str(r, 'recent_team') || t.str(r, 'team');
      case 'position':
        return t.str(r, 'position');
      case 'name':
        return t.str(r, 'player_display_name') || teamName(t.str(r, 'team'));
      case 'opponent':
        return rows.map((x) => t.str(x, 'opponent_team')).filter(Boolean).join(', ');
      case 'id':
        return t.str(r, 'player_id');
    }
  }
  const g = (col: string): number => {
    if (rows.length === 1) return t.num(rows[0], col) ?? 0;
    if (MAX_COLS.has(col)) return Math.max(0, ...rows.map((r) => t.num(r, col) ?? 0));
    if (AVG_COLS.has(col)) {
      const vals = rows.map((r) => t.num(r, col)).filter((v): v is number => v !== null);
      return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    }
    let s = 0;
    for (const r of rows) s += t.num(r, col) ?? 0;
    return s;
  };
  const ctx: StatContext = {
    position: t.str(rows[0], 'position'),
    games: gamesOverride ?? (t.has('games') ? g('games') : rows.length),
  };
  if (def.col) {
    if (!t.has(def.col)) return '';
    const v = g(def.col);
    return Math.round(v * 1e6) / 1e6;
  }
  if (def.calc) {
    const v = def.calc(g, ctx);
    return v === null ? '' : Math.round(v * 1e6) / 1e6;
  }
  return '';
}

// ---------- teams ----------

export interface TeamRecord {
  w: number;
  l: number;
  t: number;
  pf: number;
  pa: number;
  homeW: number;
  homeL: number;
  homeT: number;
  awayW: number;
  awayL: number;
  awayT: number;
  games: number;
  coach: string;
}

export function isTypeMatch(g: Game, type: SeasonType): boolean {
  if (type === 'ALL') return true;
  return type === 'REG' ? g.type === 'REG' : g.type !== 'REG';
}

export function teamGames(all: Games, franchise: string, season: number, type: SeasonType): Game[] {
  return all.season(season).filter((g) => isTypeMatch(g, type) && (g.home === franchise || g.away === franchise));
}

export function perspective(g: Game, franchise: string) {
  const home = g.home === franchise;
  const pf = home ? g.homeScore : g.awayScore;
  const pa = home ? g.awayScore : g.homeScore;
  const played = pf !== null && pa !== null;
  const result = !played ? null : pf! > pa! ? 'W' : pf! < pa! ? 'L' : 'T';
  return {
    home,
    neutral: g.location === 'Neutral',
    opp: home ? g.away : g.home,
    oppCode: home ? g.awayCode : g.homeCode,
    pf,
    pa,
    played,
    result,
    // Team's line: negative = favored (e.g. -3.5)
    line: g.spread === null ? null : home ? -g.spread : g.spread,
    moneyline: home ? g.homeML : g.awayML,
    qb: home ? g.homeQB : g.awayQB,
    oppQb: home ? g.awayQB : g.homeQB,
    coach: home ? g.homeCoach : g.awayCoach,
    oppCoach: home ? g.awayCoach : g.homeCoach,
  };
}

export function teamRecord(all: Games, franchise: string, season: number, type: SeasonType): TeamRecord {
  const rec: TeamRecord = { w: 0, l: 0, t: 0, pf: 0, pa: 0, homeW: 0, homeL: 0, homeT: 0, awayW: 0, awayL: 0, awayT: 0, games: 0, coach: '' };
  const coaches = new Map<string, number>();
  for (const g of teamGames(all, franchise, season, type)) {
    const p = perspective(g, franchise);
    if (p.coach) coaches.set(p.coach, (coaches.get(p.coach) ?? 0) + 1);
    if (!p.played) continue;
    rec.games++;
    rec.pf += p.pf!;
    rec.pa += p.pa!;
    const key = p.result === 'W' ? 'W' : p.result === 'L' ? 'L' : 'T';
    if (key === 'W') rec.w++;
    else if (key === 'L') rec.l++;
    else rec.t++;
    if (p.home) rec[`home${key}` as 'homeW']++;
    else rec[`away${key}` as 'awayW']++;
  }
  // Most frequent coach (handles mid-season changes)
  let max = 0;
  for (const [c, n] of coaches) if (n > max) {
    max = n;
    rec.coach = c;
  }
  return rec;
}

const fmtRecord = (w: number, l: number, t: number) => (t ? `${w}-${l}-${t}` : `${w}-${l}`);

export function recordStat(label: string, rec: TeamRecord, franchise: string, season: number): Scalar {
  const winPct = rec.games ? Math.round(((rec.w + rec.t / 2) / rec.games) * 1000) / 1000 : 0;
  const info = teamInfo(franchise);
  switch (label) {
    case 'Wins':
      return rec.w;
    case 'Losses':
      return rec.l;
    case 'Ties':
      return rec.t;
    case 'Record':
      return fmtRecord(rec.w, rec.l, rec.t);
    case 'Win %':
      return winPct;
    case 'Points For':
      return rec.pf;
    case 'Points Against':
      return rec.pa;
    case 'Point Differential':
      return rec.pf - rec.pa;
    case 'Points per Game':
      return rec.games ? Math.round((rec.pf / rec.games) * 10) / 10 : 0;
    case 'Points Allowed per Game':
      return rec.games ? Math.round((rec.pa / rec.games) * 10) / 10 : 0;
    case 'Home Record':
      return fmtRecord(rec.homeW, rec.homeL, rec.homeT);
    case 'Away Record':
      return fmtRecord(rec.awayW, rec.awayL, rec.awayT);
    case 'Head Coach':
      return rec.coach;
    case 'Division':
      return info && season >= 2002 ? `${info.conf} ${info.div}` : '';
    case 'Conference':
      return info?.conf ?? '';
    case 'Team Name':
      return teamName(franchise, season);
  }
  return '';
}

/** Rows of a team stats table for a franchise (season or one week). */
export function teamRows(t: Table, franchise: string, week?: number | null): string[][] {
  return t.rows.filter((r) => franchiseOf(t.str(r, 'team')) === franchise && (week === undefined || week === null || t.num(r, 'week') === week));
}

/** Rows of a weekly team table for this team's opponents ("allowed" stats). */
export function opponentRows(t: Table, franchise: string, week?: number | null): string[][] {
  return t.rows.filter((r) => franchiseOf(t.str(r, 'opponent_team')) === franchise && (week === undefined || week === null || t.num(r, 'week') === week));
}

export { codeInSeason };
