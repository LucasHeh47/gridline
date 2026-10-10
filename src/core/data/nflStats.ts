// Friendly NFL stat names ("Passing TDs", "Rush Yds", "PPR") mapped to nflverse columns.

/** Reads a numeric column (summed across rows); returns 0 for blanks. */
export type Getter = (col: string) => number;

export interface StatContext {
  position: string;
  games: number;
}

export interface StatDef {
  label: string;
  group: string;
  aliases?: string[];
  /** Direct column. */
  col?: string;
  /** Computed stat. Return null when it can't be computed (e.g. no attempts). */
  calc?: (g: Getter, ctx: StatContext) => number | null;
  /** Text stats. */
  text?: 'team' | 'position' | 'name' | 'opponent' | 'id';
  /** Where it's available. */
  scope: 'player' | 'team' | 'both';
  /** Team stats that come from game results instead of the stats files. */
  fromGames?: boolean;
}

const r1 = (n: number) => Math.round(n * 10) / 10;
const pct = (num: number, den: number) => (den > 0 ? r1((num / den) * 100) : null);

function passerRating(g: Getter): number | null {
  const att = g('attempts');
  if (att <= 0) return null;
  const clamp = (x: number) => Math.max(0, Math.min(2.375, x));
  const a = clamp((g('completions') / att - 0.3) * 5);
  const b = clamp((g('passing_yards') / att - 3) * 0.25);
  const c = clamp((g('passing_tds') / att) * 20);
  const d = clamp(2.375 - (g('passing_interceptions') / att) * 25);
  return r1(((a + b + c + d) / 6) * 100);
}

const fumbles = (g: Getter) => g('rushing_fumbles') + g('receiving_fumbles') + g('sack_fumbles');
const fumblesLost = (g: Getter) => g('rushing_fumbles_lost') + g('receiving_fumbles_lost') + g('sack_fumbles_lost');
const isDefense = (pos: string) => /^(DE|DT|NT|DL|LB|ILB|OLB|MLB|CB|S|SS|FS|DB|EDGE)$/i.test(pos);

export const STATS: StatDef[] = [
  // Passing
  { label: 'Completions', group: 'Passing', col: 'completions', aliases: ['cmp', 'comp', 'pass completions'], scope: 'both' },
  { label: 'Pass Attempts', group: 'Passing', col: 'attempts', aliases: ['attempts', 'att', 'passing attempts', 'pass att'], scope: 'both' },
  { label: 'Completion %', group: 'Passing', calc: (g) => pct(g('completions'), g('attempts')), aliases: ['comp pct', 'completion percentage', 'cmp%', 'comp%', 'completion pct'], scope: 'both' },
  { label: 'Passing Yards', group: 'Passing', col: 'passing_yards', aliases: ['pass yds', 'pass yards', 'passing yds', 'py'], scope: 'both' },
  { label: 'Passing TDs', group: 'Passing', col: 'passing_tds', aliases: ['pass td', 'pass tds', 'td passes', 'passing touchdowns', 'passing td', 'ptd'], scope: 'both' },
  {
    label: 'Interceptions',
    group: 'Passing',
    calc: (g, ctx) => (isDefense(ctx.position) ? g('def_interceptions') : g('passing_interceptions')),
    aliases: ['int', 'ints', 'picks'],
    scope: 'both',
  },
  { label: 'Interceptions Thrown', group: 'Passing', col: 'passing_interceptions', aliases: ['ints thrown', 'passing interceptions'], scope: 'both' },
  { label: 'Passer Rating', group: 'Passing', calc: passerRating, aliases: ['rating', 'qb rating', 'qbr'], scope: 'both' },
  { label: 'Yards per Attempt', group: 'Passing', calc: (g) => (g('attempts') > 0 ? r1(g('passing_yards') / g('attempts')) : null), aliases: ['ypa', 'y/a'], scope: 'both' },
  { label: 'Sacks Taken', group: 'Passing', col: 'sacks_suffered', aliases: ['times sacked', 'sacked', 'sacks suffered'], scope: 'both' },
  { label: 'Sack Yards Lost', group: 'Passing', col: 'sack_yards_lost', scope: 'both' },
  { label: 'Passing Air Yards', group: 'Passing', col: 'passing_air_yards', scope: 'both' },
  { label: 'Passing First Downs', group: 'Passing', col: 'passing_first_downs', scope: 'both' },
  { label: 'Passing EPA', group: 'Passing', col: 'passing_epa', scope: 'both' },
  { label: 'CPOE', group: 'Passing', col: 'passing_cpoe', aliases: ['completion percentage over expected'], scope: 'player' },
  // Rushing
  { label: 'Carries', group: 'Rushing', col: 'carries', aliases: ['rushes', 'rush attempts', 'rushing attempts', 'rush att'], scope: 'both' },
  { label: 'Rushing Yards', group: 'Rushing', col: 'rushing_yards', aliases: ['rush yds', 'rush yards', 'rushing yds', 'ry'], scope: 'both' },
  { label: 'Rushing TDs', group: 'Rushing', col: 'rushing_tds', aliases: ['rush td', 'rush tds', 'rushing touchdowns', 'rushing td'], scope: 'both' },
  { label: 'Yards per Carry', group: 'Rushing', calc: (g) => (g('carries') > 0 ? r1(g('rushing_yards') / g('carries')) : null), aliases: ['ypc', 'yards per rush'], scope: 'both' },
  { label: 'Rushing First Downs', group: 'Rushing', col: 'rushing_first_downs', scope: 'both' },
  { label: 'Rushing EPA', group: 'Rushing', col: 'rushing_epa', scope: 'both' },
  // Receiving
  { label: 'Receptions', group: 'Receiving', col: 'receptions', aliases: ['rec', 'catches'], scope: 'both' },
  { label: 'Targets', group: 'Receiving', col: 'targets', aliases: ['tgt', 'tgts'], scope: 'both' },
  { label: 'Receiving Yards', group: 'Receiving', col: 'receiving_yards', aliases: ['rec yds', 'rec yards', 'receiving yds'], scope: 'both' },
  { label: 'Receiving TDs', group: 'Receiving', col: 'receiving_tds', aliases: ['rec td', 'rec tds', 'receiving touchdowns', 'receiving td'], scope: 'both' },
  { label: 'Yards per Reception', group: 'Receiving', calc: (g) => (g('receptions') > 0 ? r1(g('receiving_yards') / g('receptions')) : null), aliases: ['ypr', 'y/r', 'yards per catch'], scope: 'both' },
  { label: 'Catch Rate', group: 'Receiving', calc: (g) => pct(g('receptions'), g('targets')), aliases: ['catch %', 'catch pct'], scope: 'player' },
  { label: 'Target Share', group: 'Receiving', col: 'target_share', scope: 'player' },
  { label: 'Receiving Air Yards', group: 'Receiving', col: 'receiving_air_yards', aliases: ['air yards'], scope: 'both' },
  { label: 'Yards After Catch', group: 'Receiving', col: 'receiving_yards_after_catch', aliases: ['yac'], scope: 'both' },
  { label: 'Receiving First Downs', group: 'Receiving', col: 'receiving_first_downs', scope: 'both' },
  { label: 'Receiving EPA', group: 'Receiving', col: 'receiving_epa', scope: 'both' },
  { label: 'WOPR', group: 'Receiving', col: 'wopr', aliases: ['weighted opportunity rating'], scope: 'player' },
  // Totals
  { label: 'Total TDs', group: 'Totals', calc: (g) => g('rushing_tds') + g('receiving_tds') + g('special_teams_tds') + g('def_tds') + g('fumble_recovery_tds'), aliases: ['touchdowns', 'tds', 'total touchdowns'], scope: 'player' },
  { label: 'Scrimmage Yards', group: 'Totals', calc: (g) => g('rushing_yards') + g('receiving_yards'), aliases: ['yards from scrimmage'], scope: 'player' },
  { label: 'Total Yards', group: 'Totals', calc: (g) => g('passing_yards') + g('rushing_yards') + g('receiving_yards'), aliases: ['total offense', 'all purpose yards'], scope: 'player' },
  { label: 'Touches', group: 'Totals', calc: (g) => g('carries') + g('receptions'), scope: 'player' },
  { label: 'Fumbles', group: 'Totals', calc: fumbles, scope: 'both' },
  { label: 'Fumbles Lost', group: 'Totals', calc: fumblesLost, scope: 'both' },
  { label: 'First Downs', group: 'Totals', calc: (g) => g('passing_first_downs') + g('rushing_first_downs') + g('receiving_first_downs'), scope: 'both' },
  { label: 'Games', group: 'Totals', calc: (_g, ctx) => ctx.games, aliases: ['games played', 'gp', 'g'], scope: 'both' },
  // Fantasy
  { label: 'Fantasy Points', group: 'Fantasy', col: 'fantasy_points', aliases: ['standard', 'standard points', 'fantasy', 'fpts'], scope: 'player' },
  { label: 'Fantasy Points PPR', group: 'Fantasy', col: 'fantasy_points_ppr', aliases: ['ppr', 'ppr points'], scope: 'player' },
  { label: 'Fantasy Points Half PPR', group: 'Fantasy', calc: (g) => r1(g('fantasy_points') + 0.5 * g('receptions')), aliases: ['half ppr', 'half-ppr', 'half'], scope: 'player' },
  { label: 'PPR Points per Game', group: 'Fantasy', calc: (g, ctx) => (ctx.games > 0 ? r1(g('fantasy_points_ppr') / ctx.games) : null), aliases: ['ppr per game', 'ppg'], scope: 'player' },
  // Defense
  { label: 'Tackles', group: 'Defense', calc: (g) => g('def_tackles_solo') + g('def_tackle_assists'), aliases: ['total tackles', 'tkl'], scope: 'both' },
  { label: 'Solo Tackles', group: 'Defense', col: 'def_tackles_solo', scope: 'both' },
  { label: 'Assisted Tackles', group: 'Defense', col: 'def_tackle_assists', scope: 'both' },
  { label: 'Tackles for Loss', group: 'Defense', col: 'def_tackles_for_loss', aliases: ['tfl'], scope: 'both' },
  {
    label: 'Sacks',
    group: 'Defense',
    calc: (g, ctx) => (ctx.position === 'QB' ? g('sacks_suffered') : g('def_sacks')),
    scope: 'both',
  },
  { label: 'Def Sacks', group: 'Defense', col: 'def_sacks', aliases: ['defensive sacks', 'sacks made'], scope: 'both' },
  { label: 'QB Hits', group: 'Defense', col: 'def_qb_hits', scope: 'both' },
  { label: 'Def Interceptions', group: 'Defense', col: 'def_interceptions', aliases: ['defensive interceptions', 'interceptions made'], scope: 'both' },
  { label: 'Passes Defended', group: 'Defense', col: 'def_pass_defended', aliases: ['pd', 'pass deflections', 'passes defensed'], scope: 'both' },
  { label: 'Forced Fumbles', group: 'Defense', col: 'def_fumbles_forced', aliases: ['ff'], scope: 'both' },
  { label: 'Fumble Recoveries', group: 'Defense', col: 'fumble_recovery_opp', scope: 'both' },
  { label: 'Def TDs', group: 'Defense', col: 'def_tds', aliases: ['defensive tds', 'defensive touchdowns'], scope: 'both' },
  { label: 'Safeties', group: 'Defense', col: 'def_safeties', scope: 'both' },
  // Kicking & punting
  { label: 'FG Made', group: 'Kicking', col: 'fg_made', aliases: ['field goals', 'field goals made', 'fgm'], scope: 'both' },
  { label: 'FG Attempts', group: 'Kicking', col: 'fg_att', aliases: ['field goal attempts', 'fga'], scope: 'both' },
  { label: 'FG %', group: 'Kicking', calc: (g) => pct(g('fg_made'), g('fg_att')), aliases: ['field goal %', 'fg pct'], scope: 'both' },
  { label: 'Longest FG', group: 'Kicking', col: 'fg_long', aliases: ['long fg', 'fg long'], scope: 'both' },
  { label: 'FG Made 50+', group: 'Kicking', calc: (g) => g('fg_made_50_59') + g('fg_made_60_'), aliases: ['50+ fg', 'fg 50+'], scope: 'both' },
  { label: 'XP Made', group: 'Kicking', col: 'pat_made', aliases: ['extra points', 'extra points made', 'pat made'], scope: 'both' },
  { label: 'XP Attempts', group: 'Kicking', col: 'pat_att', aliases: ['extra point attempts', 'pat att'], scope: 'both' },
  { label: 'XP %', group: 'Kicking', calc: (g) => pct(g('pat_made'), g('pat_att')), aliases: ['extra point %', 'pat pct'], scope: 'both' },
  { label: 'Punts', group: 'Punting', col: 'pt_att', scope: 'both' },
  { label: 'Punt Yards', group: 'Punting', col: 'pt_yards', scope: 'both' },
  { label: 'Net Punt Yards', group: 'Punting', col: 'pt_net_yards', scope: 'both' },
  { label: 'Punts Inside 20', group: 'Punting', col: 'pt_inside_20', scope: 'both' },
  // Returns
  { label: 'Kick Returns', group: 'Returns', col: 'kickoff_returns', scope: 'both' },
  { label: 'Kick Return Yards', group: 'Returns', col: 'kickoff_return_yards', scope: 'both' },
  { label: 'Punt Returns', group: 'Returns', col: 'punt_returns', scope: 'both' },
  { label: 'Punt Return Yards', group: 'Returns', col: 'punt_return_yards', scope: 'both' },
  { label: 'Return TDs', group: 'Returns', col: 'special_teams_tds', aliases: ['special teams tds'], scope: 'both' },
  { label: 'Penalties', group: 'Misc', col: 'penalties', scope: 'both' },
  { label: 'Penalty Yards', group: 'Misc', col: 'penalty_yards', scope: 'both' },
  // Info
  { label: 'Team', group: 'Info', text: 'team', aliases: ['tm'], scope: 'player' },
  { label: 'Position', group: 'Info', text: 'position', aliases: ['pos'], scope: 'player' },
  { label: 'Name', group: 'Info', text: 'name', aliases: ['player'], scope: 'both' },
  { label: 'Opponent', group: 'Info', text: 'opponent', aliases: ['opp', 'vs'], scope: 'player' },
  { label: 'Player ID', group: 'Info', text: 'id', aliases: ['gsis id', 'id'], scope: 'player' },
  // Team results (from game scores)
  { label: 'Wins', group: 'Record', fromGames: true, aliases: ['w'], scope: 'team' },
  { label: 'Losses', group: 'Record', fromGames: true, aliases: ['l'], scope: 'team' },
  { label: 'Ties', group: 'Record', fromGames: true, aliases: ['t'], scope: 'team' },
  { label: 'Record', group: 'Record', fromGames: true, aliases: ['w-l', 'win-loss'], scope: 'team' },
  { label: 'Win %', group: 'Record', fromGames: true, aliases: ['win pct', 'win percentage', 'pct'], scope: 'team' },
  { label: 'Points For', group: 'Record', fromGames: true, aliases: ['points', 'pf', 'points scored'], scope: 'team' },
  { label: 'Points Against', group: 'Record', fromGames: true, aliases: ['pa', 'points allowed'], scope: 'team' },
  { label: 'Point Differential', group: 'Record', fromGames: true, aliases: ['diff', 'point diff', 'net points'], scope: 'team' },
  { label: 'Points per Game', group: 'Record', fromGames: true, aliases: ['ppg scored'], scope: 'team' },
  { label: 'Points Allowed per Game', group: 'Record', fromGames: true, aliases: ['papg'], scope: 'team' },
  { label: 'Home Record', group: 'Record', fromGames: true, scope: 'team' },
  { label: 'Away Record', group: 'Record', fromGames: true, aliases: ['road record'], scope: 'team' },
  { label: 'Head Coach', group: 'Record', fromGames: true, aliases: ['coach'], scope: 'team' },
  { label: 'Division', group: 'Record', fromGames: true, scope: 'team' },
  { label: 'Conference', group: 'Record', fromGames: true, scope: 'team' },
  { label: 'Team Name', group: 'Record', fromGames: true, aliases: ['full name'], scope: 'team' },
  // Team-only computed
  { label: 'Turnovers', group: 'Totals', calc: (g) => g('passing_interceptions') + fumblesLost(g), aliases: ['giveaways'], scope: 'team' },
  { label: 'Takeaways', group: 'Defense', calc: (g) => g('def_interceptions') + g('fumble_recovery_opp'), scope: 'team' },
  {
    label: 'Turnover Differential',
    group: 'Defense',
    calc: (g) => g('def_interceptions') + g('fumble_recovery_opp') - (g('passing_interceptions') + fumblesLost(g)),
    aliases: ['turnover margin', 'to diff'],
    scope: 'team',
  },
  { label: 'Offensive Yards', group: 'Totals', calc: (g) => g('passing_yards') + g('rushing_yards'), aliases: ['total offense yards', 'yards'], scope: 'team' },
];

export function normKey(s: string): string {
  return s.toLowerCase().replace(/%/g, 'pct').replace(/\+/g, 'plus').replace(/[^a-z0-9]/g, '');
}

const INDEX = new Map<string, StatDef[]>();
(() => {
  const add = (k: string, d: StatDef) => {
    const key = normKey(k);
    const list = INDEX.get(key) ?? [];
    if (!list.includes(d)) list.push(d);
    INDEX.set(key, list);
  };
  for (const d of STATS) {
    add(d.label, d);
    for (const a of d.aliases ?? []) add(a, d);
    if (d.col) add(d.col, d);
  }
})();

export interface ResolvedStat {
  def: StatDef;
  /** For team stats: count what opponents did ("Rushing Yards Allowed"). */
  allowed: boolean;
}

/** Finds a stat by friendly name. Raw nflverse column names (e.g. "passing_air_yards") also work. */
export function resolveStat(name: string, scope: 'player' | 'team', columns?: string[]): ResolvedStat | { error: string } {
  let raw = name.trim();
  let allowed = false;
  const m = /^(.*?)\s+(allowed|against|given up)$/i.exec(raw);
  if (scope === 'team' && m && !/^points?$/i.test(m[1])) {
    raw = m[1];
    allowed = true;
  }
  const key = normKey(raw);
  let list = INDEX.get(key) ?? INDEX.get(key.replace(/s$/, '')) ?? INDEX.get(key + 's');
  if (list) {
    const ok = list.filter((d) => d.scope === 'both' || d.scope === scope);
    if (ok.length) {
      const def = ok[0];
      if (allowed && (def.fromGames || def.text)) return { error: `"${name}" can't be counted for opponents.` };
      return { def, allowed };
    }
    return { error: `"${name}" is a ${scope === 'team' ? 'player' : 'team'} stat. Use ${scope === 'team' ? 'NFL.PLAYER' : 'NFL.TEAM'} for it.` };
  }
  // Any raw column from the data file
  if (columns) {
    const col = columns.find((c) => normKey(c) === key);
    if (col) return { def: { label: col, group: 'Raw', col, scope }, allowed };
  }
  list = undefined;
  const sug = suggestStat(raw, scope);
  return { error: `Unknown stat "${name}".${sug ? ` Did you mean "${sug}"?` : ''} NFL.STATS() lists every stat.` };
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

export function suggestStat(name: string, scope: 'player' | 'team'): string | null {
  const key = normKey(name);
  let best: string | null = null;
  let bestD = Infinity;
  for (const d of STATS) {
    if (d.scope !== 'both' && d.scope !== scope) continue;
    for (const k of [d.label, ...(d.aliases ?? [])]) {
      const dist = lev(key, normKey(k));
      if (dist < bestD) {
        bestD = dist;
        best = d.label;
      }
    }
  }
  return bestD <= Math.max(2, Math.floor(key.length / 3)) ? best : null;
}

export function statLabels(scope: 'player' | 'team'): string[] {
  return STATS.filter((d) => d.scope === 'both' || d.scope === scope).map((d) => d.label);
}

/** Columns that shouldn't be summed when combining several rows. */
export const MAX_COLS = new Set(['fg_long', 'pt_long']);
export const AVG_COLS = new Set(['passing_cpoe', 'pacr', 'racr', 'target_share', 'air_yards_share', 'wopr', 'fg_pct', 'pat_pct']);
