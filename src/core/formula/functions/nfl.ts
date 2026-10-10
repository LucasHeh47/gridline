// NFL.* functions: player, team and game data from nflverse, downloaded on demand.
import {
  currentSeason,
  findPlayer,
  FIRST_SEASON,
  games,
  parseSeason,
  parseSeasonType,
  parseWeek,
  perspective,
  playerSeason,
  players,
  playerWeeks,
  recordStat,
  statFromRows,
  teamGames,
  teamRecord,
  teamRows,
  opponentRows,
  teamSeason,
  teamWeeks,
  type Game,
  type PlayerTable,
  type SeasonType,
} from '../../data/nfl';
import { resolveStat, STATS, statLabels, type ResolvedStat } from '../../data/nflStats';
import { resolveTeam, teamInfo, teamName, TEAMS } from '../../data/nflTeams';
import { dateToSerial } from '../../dates';
import { Errors, ErrorValue, isError, type Scalar } from '../../values';
import { argText, eachValue, scalarOf, toText } from '../coerce';
import { Matrix, type Arg, type Ctx } from '../types';
import { define } from './registry';

function textArg(ctx: Ctx, a: Arg, what: string): string | ErrorValue {
  const v = scalarOf(ctx, a);
  if (v instanceof ErrorValue) return v;
  if (v === null || v === '') return Errors.value(`Enter ${what}.`);
  const t = toText(v);
  return isError(t) ? t : t;
}

function opt(ctx: Ctx, a: Arg): Scalar | undefined {
  return a === undefined ? undefined : scalarOf(ctx, a);
}

/** A list of names from a range, an array, or one text with commas ("Passing TDs, Rushing TDs"). */
function nameList(ctx: Ctx, a: Arg): string[] {
  const out: string[] = [];
  eachValue(ctx, a, (v) => {
    if (v === null || v === '') return;
    const t = toText(v);
    if (isError(t)) return;
    for (const part of t.split(',')) if (part.trim()) out.push(part.trim());
  });
  return out;
}

function teamArg(ctx: Ctx, a: Arg): string | ErrorValue {
  const t = textArg(ctx, a, 'a team, like "KC" or "Chiefs"');
  if (isError(t)) return t;
  const r = resolveTeam(t);
  return 'error' in r ? Errors.value(r.error) : r.abbr;
}

const ROUND_NAMES: Record<string, string> = { WC: 'WC', DIV: 'DIV', CON: 'CON', SB: 'SB' };

function weekLabel(g: Game): Scalar {
  return g.type === 'REG' ? g.week : ROUND_NAMES[g.type] ?? g.type;
}

/** "2024-09-08" as a date serial (so it sorts and formats like a date), or the text if it isn't one. */
function gameDate(g: Game): Scalar {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(g.date);
  return m ? dateToSerial(+m[1], +m[2], +m[3]) : g.date;
}

// ---------- NFL.PLAYER ----------

define(
  'NFL.PLAYER',
  'NFL',
  'player, stat, [season], [week], [season_type]',
  'A player\'s stat for a season (or one week). Example: =NFL.PLAYER("Patrick Mahomes", "Passing TDs", 2023). Use ranges for player, stat or season to fill a whole table, and drag down to step through seasons.',
  (ctx, [p, s, se, wk, st]) => {
    const player = textArg(ctx, p, 'a player name, like "Patrick Mahomes"');
    if (isError(player)) return player;
    const statName = textArg(ctx, s, 'a stat, like "Passing Yards" (NFL.STATS() lists them all)');
    if (isError(statName)) return statName;
    const season = parseSeason(opt(ctx, se));
    if (isError(season)) return season;
    const type = parseSeasonType(opt(ctx, st));
    if (isError(type)) return type;
    const week = parseWeek(opt(ctx, wk), season);
    if (isError(week)) return week;
    if (week === null) {
      const got = playerSeason(season, type);
      if (!got.ok) return got.error;
      const rs = resolveStat(statName, 'player', got.value.headers);
      if ('error' in rs) return Errors.value(rs.error);
      const found = findPlayer(got.value, player, `${season} stats`);
      if (isError(found)) return found;
      return statFromRows(got.value, found.rows, rs);
    }
    const got = playerWeeks(season);
    if (!got.ok) return got.error;
    const rs = resolveStat(statName, 'player', got.value.headers);
    if ('error' in rs) return Errors.value(rs.error);
    const found = findPlayer(got.value, player, `${season} game stats`);
    if (isError(found)) return found;
    const rows = found.rows.filter((r) => got.value.num(r, 'week') === week);
    if (!rows.length) {
      const name = got.value.str(found.rows[0], 'player_display_name');
      return Errors.na(`${name} has no stats for week ${week} of ${season} (bye week, injury, or not played yet).`);
    }
    return statFromRows(got.value, rows, rs, rows.length);
  },
  {
    lift: [0, 1, 2, 3, 4],
    external: true,
    series: { season: 2, week: 3 },
    suggest: { 0: 'nflPlayer', 1: 'nflPlayerStat', 4: 'nflSeasonType' },
    argHelp: {
      player: 'Name like "Josh Allen". Add a team or position in parentheses if two players share a name: "Josh Allen (BUF)".',
      stat: 'Like "Passing Yards", "Rushing TDs", "Receptions", "Fantasy Points PPR". NFL.STATS() lists all.',
      season: `Year from ${FIRST_SEASON}. Leave out for the current season.`,
      week: 'Week number, or "WC", "DIV", "CON", "SB" for the playoffs. Leave out for season totals.',
      season_type: '"REG" (default), "POST" for playoffs, or "ALL".',
    },
  },
);

// ---------- NFL.PLAYERSTATS ----------

define(
  'NFL.PLAYERSTATS',
  'NFL',
  'player, stats, [first_season], [last_season], [season_type]',
  'A season-by-season table for a player. Example: =NFL.PLAYERSTATS("Tom Brady", "Passing TDs, Rushing TDs, Passing Yards", 2015, 2020). Leave out the seasons for the whole career.',
  (ctx, [p, s, fs, ls, st]) => {
    const player = textArg(ctx, p, 'a player name');
    if (isError(player)) return player;
    const stats = nameList(ctx, s);
    if (!stats.length) return Errors.value('List the stats you want, like "Passing Yards, Passing TDs".');
    const type = parseSeasonType(opt(ctx, st));
    if (isError(type)) return type;
    const cur = currentSeason();
    let first: number;
    let last: number;
    const fsv = opt(ctx, fs);
    if (fsv === undefined || fsv === null || fsv === '') {
      // Whole career, from the player profile
      const pl = players();
      if (!pl.ok) return pl.error;
      const prof = pl.value.find(player);
      if (!prof) return Errors.na(`No NFL player named "${player}".`);
      first = Math.max(FIRST_SEASON, pl.value.num(prof, 'rookie_season') ?? FIRST_SEASON);
      last = Math.min(cur, pl.value.num(prof, 'last_season') ?? cur);
    } else {
      const f = parseSeason(fsv);
      if (isError(f)) return f;
      first = f;
      const lsv = opt(ctx, ls);
      const l = lsv === undefined || lsv === null || lsv === '' ? cur : parseSeason(lsv);
      if (isError(l)) return l;
      last = l;
    }
    if (last < first) [first, last] = [last, first];
    if (last - first > 30) return Errors.value('That\'s more than 30 seasons. Pick a smaller range.');
    // Start every download at once so they run in parallel
    const tables = new Map<number, PlayerTable | null>();
    let busy: ErrorValue | null = null;
    for (let y = first; y <= last; y++) {
      const got = playerSeason(y, type);
      if (got.ok) tables.set(y, got.value);
      else if (got.error.code === '#BUSY!') busy = got.error;
      else if (got.error.code === '#N/A') tables.set(y, null);
      else return got.error;
    }
    if (busy) return busy;
    // Identify the player once (latest season they appear), then follow their id
    let id = '';
    let lastErr: ErrorValue | null = null;
    for (let y = last; y >= first && !id; y--) {
      const t = tables.get(y);
      if (!t) continue;
      const found = findPlayer(t, player, `stats between ${first} and ${last}`);
      if (isError(found)) lastErr = found;
      else id = found.id;
    }
    if (!id) return lastErr ?? Errors.na(`No stats for "${player}" between ${first} and ${last}.`);
    const resolved: ResolvedStat[] = [];
    const anyTable = [...tables.values()].find((t) => t) as PlayerTable;
    for (const name of stats) {
      const rs = resolveStat(name, 'player', anyTable.headers);
      if ('error' in rs) return Errors.value(rs.error);
      resolved.push(rs);
    }
    const header: Scalar[] = ['Season', 'Team', ...resolved.map((r) => r.def.label)];
    const out: Scalar[][] = [header];
    for (let y = first; y <= last; y++) {
      const t = tables.get(y);
      const rows = t ? t.rowsForId(id) : [];
      if (!rows.length) {
        out.push([y, '', ...resolved.map(() => '')]);
        continue;
      }
      out.push([y, t!.str(rows[rows.length - 1], 'recent_team'), ...resolved.map((rs) => statFromRows(t!, rows, rs))]);
    }
    return new Matrix(out);
  },
  {
    external: true,
    suggest: { 0: 'nflPlayer', 1: 'nflPlayerStat', 4: 'nflSeasonType' },
    argHelp: {
      stats: 'A range of stat names, an array like {"Passing TDs","Passing Yards"}, or one text with commas.',
      first_season: 'First year. Leave out for the whole career.',
      last_season: 'Last year. Leave out for "through the current season".',
    },
  },
);

// ---------- NFL.GAMELOG ----------

const DEFAULT_LOG: Record<string, string[]> = {
  QB: ['Completions', 'Pass Attempts', 'Passing Yards', 'Passing TDs', 'Interceptions Thrown', 'Carries', 'Rushing Yards', 'Rushing TDs', 'Fantasy Points PPR'],
  RB: ['Carries', 'Rushing Yards', 'Rushing TDs', 'Targets', 'Receptions', 'Receiving Yards', 'Receiving TDs', 'Fantasy Points PPR'],
  WR: ['Targets', 'Receptions', 'Receiving Yards', 'Receiving TDs', 'Carries', 'Rushing Yards', 'Fantasy Points PPR'],
  TE: ['Targets', 'Receptions', 'Receiving Yards', 'Receiving TDs', 'Fantasy Points PPR'],
  K: ['FG Made', 'FG Attempts', 'Longest FG', 'XP Made', 'XP Attempts'],
  P: ['Punts', 'Punt Yards', 'Net Punt Yards', 'Punts Inside 20'],
  DEF: ['Tackles', 'Solo Tackles', 'Tackles for Loss', 'Def Sacks', 'QB Hits', 'Def Interceptions', 'Passes Defended', 'Forced Fumbles'],
};

define(
  'NFL.GAMELOG',
  'NFL',
  'player, [season], [stats], [season_type]',
  'A player\'s game-by-game stats for a season, with opponents and results. Example: =NFL.GAMELOG("Justin Jefferson", 2022).',
  (ctx, [p, se, s, st]) => {
    const player = textArg(ctx, p, 'a player name');
    if (isError(player)) return player;
    const season = parseSeason(opt(ctx, se));
    if (isError(season)) return season;
    const typeArg = opt(ctx, st);
    const type: SeasonType | ErrorValue = typeArg === undefined || typeArg === null || typeArg === '' ? 'ALL' : parseSeasonType(typeArg);
    if (isError(type)) return type;
    const got = playerWeeks(season);
    const gg = games();
    if (!got.ok) return got.error;
    if (!gg.ok) return gg.error;
    const t = got.value;
    const found = findPlayer(t, player, `${season} game stats`);
    if (isError(found)) return found;
    const pos = t.str(found.rows[0], 'position');
    const group = DEFAULT_LOG[pos] ? pos : ['FB'].includes(pos) ? 'RB' : /^(DE|DT|NT|DL|LB|ILB|OLB|MLB|CB|S|SS|FS|DB)$/.test(pos) ? 'DEF' : 'WR';
    const statNames = s === undefined ? DEFAULT_LOG[group] : nameList(ctx, s);
    const resolved: ResolvedStat[] = [];
    for (const name of statNames) {
      const rs = resolveStat(name, 'player', t.headers);
      if ('error' in rs) return Errors.value(rs.error);
      resolved.push(rs);
    }
    const gameById = new Map(gg.value.season(season).map((g) => [g.id, g]));
    const rows = found.rows
      .filter((r) => type === 'ALL' || (type === 'REG' ? t.str(r, 'season_type') === 'REG' : t.str(r, 'season_type') !== 'REG'))
      .sort((a, b) => (t.num(a, 'week') ?? 0) - (t.num(b, 'week') ?? 0));
    const out: Scalar[][] = [['Week', 'Opponent', 'Result', ...resolved.map((r) => r.def.label)]];
    for (const r of rows) {
      const g = gameById.get(t.str(r, 'game_id'));
      const team = t.str(r, 'team');
      let matchup = t.str(r, 'opponent_team');
      let result: Scalar = '';
      let week: Scalar = t.num(r, 'week') ?? '';
      if (g) {
        const pv = perspective(g, teamInfo(team)?.abbr ?? team);
        matchup = `${pv.home ? 'vs' : '@'} ${pv.oppCode}`;
        if (pv.played) result = `${pv.result} ${pv.pf}-${pv.pa}${g.overtime ? ' OT' : ''}`;
        week = weekLabel(g);
      }
      out.push([week, matchup, result, ...resolved.map((rs) => statFromRows(t, [r], rs, 1))]);
    }
    if (out.length === 1) return Errors.na(`No games for ${player} in ${season}.`);
    return new Matrix(out);
  },
  { external: true, suggest: { 0: 'nflPlayer', 2: 'nflPlayerStat', 3: 'nflSeasonType' } },
);

// ---------- NFL.TEAM ----------

function gameStat(label: string, g: Game, franchise: string): Scalar | undefined {
  const p = perspective(g, franchise);
  switch (label) {
    case 'Wins':
      return p.result === 'W' ? 1 : 0;
    case 'Losses':
      return p.result === 'L' ? 1 : 0;
    case 'Ties':
      return p.result === 'T' ? 1 : 0;
    case 'Record':
      return p.result ?? '';
    case 'Points For':
      return p.pf ?? '';
    case 'Points Against':
      return p.pa ?? '';
    case 'Point Differential':
      return p.played ? p.pf! - p.pa! : '';
    case 'Head Coach':
      return p.coach;
  }
  return undefined;
}

define(
  'NFL.TEAM',
  'NFL',
  'team, stat, [season], [week], [season_type]',
  'A team stat for a season or one week: wins, points, yards, turnovers… Add "Allowed" for what opponents did: "Rushing Yards Allowed". Example: =NFL.TEAM("Chiefs", "Wins", 2023).',
  (ctx, [tm, s, se, wk, st]) => {
    const franchise = teamArg(ctx, tm);
    if (isError(franchise)) return franchise;
    const statName = textArg(ctx, s, 'a stat, like "Wins" or "Passing Yards"');
    if (isError(statName)) return statName;
    const season = parseSeason(opt(ctx, se));
    if (isError(season)) return season;
    const type = parseSeasonType(opt(ctx, st));
    if (isError(type)) return type;
    const week = parseWeek(opt(ctx, wk), season);
    if (isError(week)) return week;
    const rs = resolveStat(statName, 'team');
    if ('error' in rs && !/Unknown stat/.test(rs.error)) return Errors.value(rs.error);

    if (!('error' in rs) && rs.def.fromGames) {
      const gg = games();
      if (!gg.ok) return gg.error;
      if (week !== null) {
        const g = teamGames(gg.value, franchise, season, 'ALL').find((x) => x.week === week);
        if (!g) return Errors.na(`${franchise} didn't play in week ${week} of ${season} (bye week?).`);
        const v = gameStat(rs.def.label, g, franchise);
        if (v !== undefined) return v;
      }
      return recordStat(rs.def.label, teamRecord(gg.value, franchise, season, type), franchise, season);
    }
    // Stats from the team stats files
    if (week !== null || (!('error' in rs) && rs.allowed)) {
      const tw = teamWeeks(season);
      if (!tw.ok) return tw.error;
      const t = tw.value;
      const rs2 = 'error' in rs ? resolveStat(statName, 'team', t.headers) : rs;
      if ('error' in rs2) return Errors.value(rs2.error);
      const pick = rs2.allowed ? opponentRows(t, franchise, week) : teamRows(t, franchise, week);
      const rows = week !== null ? pick : pick.filter((r) => type === 'ALL' || (type === 'REG' ? t.str(r, 'season_type') === 'REG' : t.str(r, 'season_type') !== 'REG'));
      if (!rows.length) return Errors.na(week !== null ? `${franchise} didn't play in week ${week} of ${season} (bye week?).` : `No ${season} games found for ${franchise}.`);
      return statFromRows(t, rows, rs2, rows.length);
    }
    const ts = teamSeason(season, type);
    if (!ts.ok) return ts.error;
    const rs2 = 'error' in rs ? resolveStat(statName, 'team', ts.value.headers) : rs;
    if ('error' in rs2) return Errors.value(rs2.error);
    const rows = teamRows(ts.value, franchise);
    if (!rows.length) return Errors.na(`No ${season} stats for ${teamName(franchise, season)}.`);
    return statFromRows(ts.value, rows, rs2);
  },
  {
    lift: [0, 1, 2, 3, 4],
    external: true,
    series: { season: 2, week: 3 },
    suggest: { 0: 'nflTeam', 1: 'nflTeamStat', 4: 'nflSeasonType' },
    argHelp: {
      team: 'Code like "KC", or a name like "Chiefs" / "Kansas City Chiefs". Old names work too ("Oakland Raiders").',
      stat: 'Like "Wins", "Record", "Points For", "Passing Yards", "Turnovers", "Sacks", or add "Allowed": "Passing Yards Allowed".',
    },
  },
);

// ---------- NFL.GAME ----------

const GAME_FIELDS = [
  'Result',
  'Score',
  'Opponent',
  'Opponent Name',
  'Matchup',
  'Home/Away',
  'Points For',
  'Points Against',
  'Margin',
  'Outcome',
  'Date',
  'Day',
  'Kickoff',
  'Spread',
  'Covered',
  'Over/Under',
  'Total Points',
  'Over/Under Result',
  'Moneyline',
  'QB',
  'Opponent QB',
  'Coach',
  'Opponent Coach',
  'Stadium',
  'Roof',
  'Surface',
  'Temperature',
  'Wind',
  'Overtime',
  'Division Game',
  'Game ID',
];

define(
  'NFL.GAME',
  'NFL',
  'team, season, week, [field]',
  'Details of one game from a team\'s point of view: result, score, opponent, spread, over/under… Example: =NFL.GAME("Eagles", 2024, 5, "Score").',
  (ctx, [tm, se, wk, f]) => {
    const franchise = teamArg(ctx, tm);
    if (isError(franchise)) return franchise;
    const season = parseSeason(opt(ctx, se));
    if (isError(season)) return season;
    const week = parseWeek(opt(ctx, wk), season);
    if (isError(week)) return week;
    if (week === null) return Errors.value('Enter a week number (or "WC", "DIV", "CON", "SB").');
    const fieldRaw = opt(ctx, f);
    const field = fieldRaw === undefined || fieldRaw === null || fieldRaw === '' ? 'Result' : String(fieldRaw);
    const key = field.toLowerCase().replace(/[^a-z]/g, '');
    const canonical = GAME_FIELDS.find((x) => x.toLowerCase().replace(/[^a-z]/g, '') === key) ?? ({ opp: 'Opponent', vs: 'Opponent', homeaway: 'Home/Away', line: 'Spread', ou: 'Over/Under', total: 'Over/Under', ats: 'Covered', time: 'Kickoff', temp: 'Temperature', ot: 'Overtime' } as Record<string, string>)[key];
    if (!canonical) return Errors.value(`Unknown game field "${field}". Try: ${GAME_FIELDS.slice(0, 12).join(', ')}…`);
    const gg = games();
    if (!gg.ok) return gg.error;
    const g = teamGames(gg.value, franchise, season, 'ALL').find((x) => x.week === week);
    if (!g) return Errors.na(`${teamName(franchise, season)} didn't play in week ${week} of ${season} (bye week?).`);
    const p = perspective(g, franchise);
    switch (canonical) {
      case 'Result':
        return p.played ? `${p.result} ${p.pf}-${p.pa}${g.overtime ? ' (OT)' : ''}` : 'Scheduled';
      case 'Score':
        return p.played ? `${p.pf}-${p.pa}` : '';
      case 'Opponent':
        return p.oppCode;
      case 'Opponent Name':
        return teamName(p.opp, season);
      case 'Matchup':
        return `${p.neutral ? 'vs' : p.home ? 'vs' : '@'} ${p.oppCode}`;
      case 'Home/Away':
        return p.neutral ? 'Neutral' : p.home ? 'Home' : 'Away';
      case 'Points For':
        return p.pf ?? '';
      case 'Points Against':
        return p.pa ?? '';
      case 'Margin':
        return p.played ? p.pf! - p.pa! : '';
      case 'Outcome':
        return p.result ?? '';
      case 'Date':
        return g.date;
      case 'Day':
        return g.weekday;
      case 'Kickoff':
        return g.time;
      case 'Spread':
        return p.line ?? '';
      case 'Covered': {
        if (!p.played || p.line === null) return '';
        const ats = p.pf! - p.pa! + p.line;
        return ats > 0 ? 'Yes' : ats < 0 ? 'No' : 'Push';
      }
      case 'Over/Under':
        return g.totalLine ?? '';
      case 'Total Points':
        return p.played ? p.pf! + p.pa! : '';
      case 'Over/Under Result': {
        if (!p.played || g.totalLine === null) return '';
        const tot = p.pf! + p.pa!;
        return tot > g.totalLine ? 'Over' : tot < g.totalLine ? 'Under' : 'Push';
      }
      case 'Moneyline':
        return p.moneyline ?? '';
      case 'QB':
        return p.qb;
      case 'Opponent QB':
        return p.oppQb;
      case 'Coach':
        return p.coach;
      case 'Opponent Coach':
        return p.oppCoach;
      case 'Stadium':
        return g.stadium;
      case 'Roof':
        return g.roof;
      case 'Surface':
        return g.surface;
      case 'Temperature':
        return g.temp ?? '';
      case 'Wind':
        return g.wind ?? '';
      case 'Overtime':
        return g.overtime;
      case 'Division Game':
        return g.divGame;
      case 'Game ID':
        return g.id;
    }
    return '';
  },
  { lift: [0, 1, 2, 3], external: true, series: { season: 1, week: 2 }, suggest: { 0: 'nflTeam', 3: 'nflGameField' }, min: 3 },
);

// ---------- NFL.SCHEDULE ----------

define(
  'NFL.SCHEDULE',
  'NFL',
  'team, [season], [season_type]',
  'A team\'s full schedule with results, scores and betting lines. Example: =NFL.SCHEDULE("Bills", 2024).',
  (ctx, [tm, se, st]) => {
    const franchise = teamArg(ctx, tm);
    if (isError(franchise)) return franchise;
    const season = parseSeason(opt(ctx, se));
    if (isError(season)) return season;
    const typeArg = opt(ctx, st);
    const type: SeasonType | ErrorValue = typeArg === undefined || typeArg === null || typeArg === '' ? 'ALL' : parseSeasonType(typeArg);
    if (isError(type)) return type;
    const gg = games();
    if (!gg.ok) return gg.error;
    const list = teamGames(gg.value, franchise, season, type).sort((a, b) => a.week - b.week);
    if (!list.length) return Errors.na(`No ${season} games for ${teamName(franchise, season)}.`);
    const out: Scalar[][] = [['Week', 'Date', 'Matchup', 'Result', 'Score', 'Spread', 'Over/Under']];
    // Regular-season weeks the team didn't play are bye weeks
    const regWeeks = new Set(gg.value.season(season).filter((g) => g.type === 'REG').map((g) => g.week));
    const played = new Set(list.filter((g) => g.type === 'REG').map((g) => g.week));
    const byes = type === 'POST' ? [] : [...regWeeks].filter((w) => !played.has(w) && w > 0);
    const rows: (Game | number)[] = [...list];
    for (const w of byes) {
      const at = rows.findIndex((x) => typeof x !== 'number' && (x.type !== 'REG' || x.week > w));
      rows.splice(at < 0 ? rows.length : at, 0, w);
    }
    for (const g of rows) {
      if (typeof g === 'number') {
        out.push([g, '', 'BYE', '', '', '', '']);
        continue;
      }
      const p = perspective(g, franchise);
      out.push([
        weekLabel(g),
        gameDate(g),
        `${p.home && !p.neutral ? 'vs' : p.neutral ? 'vs' : '@'} ${p.oppCode}`,
        p.result ?? '',
        p.played ? `${p.pf}-${p.pa}${g.overtime ? ' OT' : ''}` : '',
        p.line ?? '',
        g.totalLine ?? '',
      ]);
    }
    return new Matrix(out, [undefined, 'ddd mmm d']);
  },
  { external: true, suggest: { 0: 'nflTeam', 2: 'nflSeasonType' } },
);

// ---------- NFL.STANDINGS ----------

define(
  'NFL.STANDINGS',
  'NFL',
  '[season], [group]',
  'Standings: wins, losses, points for/against. group can be "AFC", "NFC" or a division like "NFC East". Example: =NFL.STANDINGS(2023, "AFC West").',
  (ctx, [se, gr]) => {
    const season = parseSeason(opt(ctx, se));
    if (isError(season)) return season;
    const groupRaw = opt(ctx, gr);
    const group = groupRaw === undefined || groupRaw === null ? '' : String(groupRaw).trim().toUpperCase();
    const gg = games();
    if (!gg.ok) return gg.error;
    const teams = TEAMS.filter((t) => {
      if (!group) return true;
      if (group === t.conf) return true;
      return group === `${t.conf} ${t.div.toUpperCase()}`;
    });
    if (!teams.length) return Errors.value(`Unknown group "${groupRaw}". Use "AFC", "NFC", or a division like "AFC North".`);
    const rows = teams.map((t) => {
      const rec = teamRecord(gg.value, t.abbr, season, 'REG');
      const pct = rec.games ? Math.round(((rec.w + rec.t / 2) / rec.games) * 1000) / 1000 : 0;
      return { t, rec, pct };
    });
    rows.sort((a, b) => b.pct - a.pct || b.rec.pf - b.rec.pa - (a.rec.pf - a.rec.pa) || b.rec.w - a.rec.w);
    const out: Scalar[][] = [['Rank', 'Team', 'W', 'L', 'T', 'Win %', 'PF', 'PA', 'Diff', 'Division']];
    rows.forEach((r, i) => {
      out.push([i + 1, teamName(r.t.abbr, season), r.rec.w, r.rec.l, r.rec.t, r.pct, r.rec.pf, r.rec.pa, r.rec.pf - r.rec.pa, season >= 2002 ? `${r.t.conf} ${r.t.div}` : r.t.conf]);
    });
    return new Matrix(out, [undefined, undefined, undefined, undefined, undefined, '0.000']);
  },
  { external: true, suggest: { 1: 'nflGroup' }, min: 0 },
);

// ---------- NFL.LEADERS ----------

define(
  'NFL.LEADERS',
  'NFL',
  'stat, [season], [count], [position], [season_type]',
  'The league leaders for a stat. Example: =NFL.LEADERS("Rushing Yards", 2024, 10) or only tight ends: =NFL.LEADERS("Receptions", 2024, 5, "TE").',
  (ctx, [s, se, c, pos, st]) => {
    const statName = textArg(ctx, s, 'a stat, like "Passing Yards"');
    if (isError(statName)) return statName;
    const season = parseSeason(opt(ctx, se));
    if (isError(season)) return season;
    const cv = opt(ctx, c);
    const count = cv === undefined || cv === null || cv === '' ? 10 : Math.trunc(Number(cv));
    if (!(count >= 1 && count <= 500)) return Errors.value('count must be between 1 and 500.');
    const posRaw = opt(ctx, pos);
    const position = posRaw === undefined || posRaw === null ? '' : String(posRaw).trim().toUpperCase();
    const type = parseSeasonType(opt(ctx, st));
    if (isError(type)) return type;
    const got = playerSeason(season, type);
    if (!got.ok) return got.error;
    const t = got.value;
    const rs = resolveStat(statName, 'player', t.headers);
    if ('error' in rs) return Errors.value(rs.error);
    if (rs.def.text) return Errors.value(`"${rs.def.label}" isn't a number, so it can't be ranked.`);
    const scored: { r: string[]; v: number }[] = [];
    for (const r of t.rows) {
      if (position && t.str(r, 'position').toUpperCase() !== position && t.str(r, 'position_group').toUpperCase() !== position) continue;
      const v = statFromRows(t, [r], rs);
      if (typeof v === 'number') scored.push({ r, v });
    }
    scored.sort((a, b) => b.v - a.v);
    const out: Scalar[][] = [['Rank', 'Player', 'Team', 'Position', rs.def.label]];
    scored.slice(0, count).forEach((x, i) => out.push([i + 1, t.str(x.r, 'player_display_name'), t.str(x.r, 'recent_team'), t.str(x.r, 'position'), x.v]));
    if (out.length === 1) return Errors.na('No players matched.');
    return new Matrix(out);
  },
  { external: true, suggest: { 0: 'nflPlayerStat', 4: 'nflSeasonType' } },
);

// ---------- NFL.ROSTER ----------

define(
  'NFL.ROSTER',
  'NFL',
  'team, [season], [position]',
  'Players who recorded stats for a team in a season, most productive first. Example: =NFL.ROSTER("49ers", 2023, "WR").',
  (ctx, [tm, se, pos]) => {
    const franchise = teamArg(ctx, tm);
    if (isError(franchise)) return franchise;
    const season = parseSeason(opt(ctx, se));
    if (isError(season)) return season;
    const posRaw = opt(ctx, pos);
    const position = posRaw === undefined || posRaw === null ? '' : String(posRaw).trim().toUpperCase();
    const got = playerSeason(season, 'REG');
    if (!got.ok) return got.error;
    const t = got.value;
    const rows = t.rows.filter((r) => teamInfo(t.str(r, 'recent_team'))?.abbr === franchise && (!position || t.str(r, 'position').toUpperCase() === position || t.str(r, 'position_group').toUpperCase() === position));
    rows.sort((a, b) => (t.num(b, 'fantasy_points_ppr') ?? 0) - (t.num(a, 'fantasy_points_ppr') ?? 0));
    if (!rows.length) return Errors.na(`No ${season} players found for ${teamName(franchise, season)}${position ? ` at ${position}` : ''}.`);
    const out: Scalar[][] = [['Player', 'Position', 'Games', 'PPR Points']];
    for (const r of rows) out.push([t.str(r, 'player_display_name'), t.str(r, 'position'), t.num(r, 'games') ?? '', t.num(r, 'fantasy_points_ppr') ?? '']);
    return new Matrix(out);
  },
  { external: true, suggest: { 0: 'nflTeam' } },
);

// ---------- NFL.PLAYERINFO ----------

const INFO_FIELDS = ['Name', 'Position', 'Team', 'Jersey', 'Height', 'Height Inches', 'Weight', 'Birth Date', 'Age', 'College', 'Rookie Season', 'Last Season', 'Experience', 'Draft Year', 'Draft Round', 'Draft Pick', 'Draft Team', 'Status', 'Headshot', 'Player ID'];

define(
  'NFL.PLAYERINFO',
  'NFL',
  'player, field',
  'Profile details: position, team, college, height, weight, age, draft info… Example: =NFL.PLAYERINFO("Travis Kelce", "College").',
  (ctx, [p, f]) => {
    const player = textArg(ctx, p, 'a player name');
    if (isError(player)) return player;
    const field = textArg(ctx, f, 'a field, like "College" or "Age"');
    if (isError(field)) return field;
    const key = field.toLowerCase().replace(/[^a-z]/g, '');
    const canonical = INFO_FIELDS.find((x) => x.toLowerCase().replace(/[^a-z]/g, '') === key) ?? ({ pos: 'Position', number: 'Jersey', school: 'College', born: 'Birth Date', birthday: 'Birth Date', dob: 'Birth Date', drafted: 'Draft Year', id: 'Player ID', ht: 'Height', wt: 'Weight', exp: 'Experience' } as Record<string, string>)[key];
    if (!canonical) return Errors.value(`Unknown field "${field}". Try: ${INFO_FIELDS.join(', ')}.`);
    const pl = players();
    if (!pl.ok) return pl.error;
    const t = pl.value;
    const r = t.find(player);
    if (!r) return Errors.na(`No NFL player named "${player}".`);
    const n = (c: string) => t.num(r, c) ?? '';
    switch (canonical) {
      case 'Name':
        return t.str(r, 'display_name');
      case 'Position':
        return t.str(r, 'position');
      case 'Team':
        return t.str(r, 'latest_team');
      case 'Jersey':
        return n('jersey_number');
      case 'Height': {
        const h = t.num(r, 'height');
        return h === null ? '' : `${Math.floor(h / 12)}'${h % 12}"`;
      }
      case 'Height Inches':
        return n('height');
      case 'Weight':
        return n('weight');
      case 'Birth Date':
        return t.str(r, 'birth_date');
      case 'Age': {
        const b = t.str(r, 'birth_date');
        if (!b) return '';
        const [y, m, d] = b.split('-').map(Number);
        const now = new Date();
        let age = now.getFullYear() - y;
        if (now.getMonth() + 1 < m || (now.getMonth() + 1 === m && now.getDate() < d)) age--;
        return age;
      }
      case 'College':
        return t.str(r, 'college_name');
      case 'Rookie Season':
        return n('rookie_season');
      case 'Last Season':
        return n('last_season');
      case 'Experience':
        return n('years_of_experience');
      case 'Draft Year':
        return n('draft_year');
      case 'Draft Round':
        return n('draft_round');
      case 'Draft Pick':
        return n('draft_pick');
      case 'Draft Team':
        return t.str(r, 'draft_team');
      case 'Status':
        return t.str(r, 'ngs_status_short_description') || t.str(r, 'status');
      case 'Headshot':
        return t.str(r, 'headshot');
      case 'Player ID':
        return t.str(r, 'gsis_id');
    }
    return '';
  },
  { lift: [0, 1], external: true, suggest: { 0: 'nflPlayer', 1: 'nflInfoField' } },
);

// ---------- reference lists ----------

define('NFL.TEAMS', 'NFL', '[season]', 'All 32 teams with codes, names, conference and division (names as they were in the given season).', (ctx, [se]) => {
  const season = parseSeason(opt(ctx, se));
  if (isError(season)) return season;
  const out: Scalar[][] = [['Code', 'Team', 'Conference', 'Division']];
  for (const t of TEAMS) out.push([t.abbr, teamName(t.abbr, season), t.conf, `${t.conf} ${t.div}`]);
  return new Matrix(out);
}, { custom: true, min: 0 });

define('NFL.STATS', 'NFL', '[kind]', 'Every stat name you can use with NFL.PLAYER ("player", the default) or NFL.TEAM ("team").', (ctx, [k]) => {
  const kind = argText(ctx, k, 'player');
  if (isError(kind)) return kind;
  const scope = kind.toLowerCase().startsWith('t') ? 'team' : 'player';
  const out: Scalar[][] = [['Stat', 'Category']];
  for (const d of STATS) if (d.scope === 'both' || d.scope === scope) out.push([d.label, d.group]);
  return new Matrix(out);
}, { custom: true, min: 0 });

export const NFL_GAME_FIELDS = GAME_FIELDS;
export const NFL_INFO_FIELDS = INFO_FIELDS;
export { statLabels };

