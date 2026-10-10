// NFL teams: names, divisions and relocations (so franchise history is continuous).

export interface TeamInfo {
  abbr: string;
  city: string;
  nickname: string;
  conf: 'AFC' | 'NFC';
  div: 'East' | 'North' | 'South' | 'West';
}

export const TEAMS: TeamInfo[] = [
  { abbr: 'BUF', city: 'Buffalo', nickname: 'Bills', conf: 'AFC', div: 'East' },
  { abbr: 'MIA', city: 'Miami', nickname: 'Dolphins', conf: 'AFC', div: 'East' },
  { abbr: 'NE', city: 'New England', nickname: 'Patriots', conf: 'AFC', div: 'East' },
  { abbr: 'NYJ', city: 'New York', nickname: 'Jets', conf: 'AFC', div: 'East' },
  { abbr: 'BAL', city: 'Baltimore', nickname: 'Ravens', conf: 'AFC', div: 'North' },
  { abbr: 'CIN', city: 'Cincinnati', nickname: 'Bengals', conf: 'AFC', div: 'North' },
  { abbr: 'CLE', city: 'Cleveland', nickname: 'Browns', conf: 'AFC', div: 'North' },
  { abbr: 'PIT', city: 'Pittsburgh', nickname: 'Steelers', conf: 'AFC', div: 'North' },
  { abbr: 'HOU', city: 'Houston', nickname: 'Texans', conf: 'AFC', div: 'South' },
  { abbr: 'IND', city: 'Indianapolis', nickname: 'Colts', conf: 'AFC', div: 'South' },
  { abbr: 'JAX', city: 'Jacksonville', nickname: 'Jaguars', conf: 'AFC', div: 'South' },
  { abbr: 'TEN', city: 'Tennessee', nickname: 'Titans', conf: 'AFC', div: 'South' },
  { abbr: 'DEN', city: 'Denver', nickname: 'Broncos', conf: 'AFC', div: 'West' },
  { abbr: 'KC', city: 'Kansas City', nickname: 'Chiefs', conf: 'AFC', div: 'West' },
  { abbr: 'LV', city: 'Las Vegas', nickname: 'Raiders', conf: 'AFC', div: 'West' },
  { abbr: 'LAC', city: 'Los Angeles', nickname: 'Chargers', conf: 'AFC', div: 'West' },
  { abbr: 'DAL', city: 'Dallas', nickname: 'Cowboys', conf: 'NFC', div: 'East' },
  { abbr: 'NYG', city: 'New York', nickname: 'Giants', conf: 'NFC', div: 'East' },
  { abbr: 'PHI', city: 'Philadelphia', nickname: 'Eagles', conf: 'NFC', div: 'East' },
  { abbr: 'WAS', city: 'Washington', nickname: 'Commanders', conf: 'NFC', div: 'East' },
  { abbr: 'CHI', city: 'Chicago', nickname: 'Bears', conf: 'NFC', div: 'North' },
  { abbr: 'DET', city: 'Detroit', nickname: 'Lions', conf: 'NFC', div: 'North' },
  { abbr: 'GB', city: 'Green Bay', nickname: 'Packers', conf: 'NFC', div: 'North' },
  { abbr: 'MIN', city: 'Minnesota', nickname: 'Vikings', conf: 'NFC', div: 'North' },
  { abbr: 'ATL', city: 'Atlanta', nickname: 'Falcons', conf: 'NFC', div: 'South' },
  { abbr: 'CAR', city: 'Carolina', nickname: 'Panthers', conf: 'NFC', div: 'South' },
  { abbr: 'NO', city: 'New Orleans', nickname: 'Saints', conf: 'NFC', div: 'South' },
  { abbr: 'TB', city: 'Tampa Bay', nickname: 'Buccaneers', conf: 'NFC', div: 'South' },
  { abbr: 'ARI', city: 'Arizona', nickname: 'Cardinals', conf: 'NFC', div: 'West' },
  { abbr: 'LA', city: 'Los Angeles', nickname: 'Rams', conf: 'NFC', div: 'West' },
  { abbr: 'SF', city: 'San Francisco', nickname: '49ers', conf: 'NFC', div: 'West' },
  { abbr: 'SEA', city: 'Seattle', nickname: 'Seahawks', conf: 'NFC', div: 'West' },
];

const BY_ABBR = new Map(TEAMS.map((t) => [t.abbr, t]));

/** Codes used in older data (relocations and alternate abbreviations) -> current franchise code. */
const FRANCHISE: Record<string, string> = {
  OAK: 'LV',
  SD: 'LAC',
  STL: 'LA',
  LAR: 'LA',
  JAC: 'JAX',
  WSH: 'WAS',
  ARZ: 'ARI',
  BLT: 'BAL',
  CLV: 'CLE',
  HST: 'HOU',
};

export function franchiseOf(code: string): string {
  const c = code.toUpperCase();
  return FRANCHISE[c] ?? c;
}

/** The code a franchise used in a given season (LV in 2015 -> OAK). */
export function codeInSeason(franchise: string, season: number): string {
  if (franchise === 'LV' && season <= 2019) return 'OAK';
  if (franchise === 'LAC' && season <= 2016) return 'SD';
  if (franchise === 'LA' && season <= 2015) return 'STL';
  return franchise;
}

export function teamInfo(abbr: string): TeamInfo | undefined {
  return BY_ABBR.get(franchiseOf(abbr));
}

/** Full team name as it was in a season ("Oakland Raiders" in 2015). */
export function teamName(abbr: string, season?: number): string {
  const f = franchiseOf(abbr);
  const t = BY_ABBR.get(f);
  if (!t) return abbr;
  const s = season ?? 9999;
  if (f === 'LV' && s <= 2019) return 'Oakland Raiders';
  if (f === 'LAC' && s <= 2016) return 'San Diego Chargers';
  if (f === 'LA' && s <= 2015) return 'St. Louis Rams';
  if (f === 'WAS') return s <= 2019 ? 'Washington Redskins' : s <= 2021 ? 'Washington Football Team' : 'Washington Commanders';
  return `${t.city} ${t.nickname}`;
}

function norm(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9]/g, '');
}

const ALIASES = new Map<string, string>();
(() => {
  const add = (alias: string, abbr: string) => ALIASES.set(norm(alias), abbr);
  for (const t of TEAMS) {
    add(t.abbr, t.abbr);
    add(t.nickname, t.abbr);
    add(`${t.city} ${t.nickname}`, t.abbr);
    if (!['New York', 'Los Angeles'].includes(t.city)) add(t.city, t.abbr);
  }
  for (const [old, cur] of Object.entries(FRANCHISE)) add(old, cur);
  add('Oakland Raiders', 'LV');
  add('Oakland', 'LV');
  add('San Diego Chargers', 'LAC');
  add('San Diego', 'LAC');
  add('St. Louis Rams', 'LA');
  add('St Louis Rams', 'LA');
  add('St. Louis', 'LA');
  add('Washington Redskins', 'WAS');
  add('Redskins', 'WAS');
  add('Washington Football Team', 'WAS');
  add('Football Team', 'WAS');
  add('Niners', 'SF');
  add('Bucs', 'TB');
  add('Pats', 'NE');
  add('Jags', 'JAX');
  add('Philly', 'PHI');
  add('Big Blue', 'NYG');
  add('NY Giants', 'NYG');
  add('NY Jets', 'NYJ');
  add('LA Rams', 'LA');
  add('LA Chargers', 'LAC');
})();

/** Accepts "KC", "Chiefs", "Kansas City Chiefs", "Kansas City", old names and codes. */
export function resolveTeam(input: string): { abbr: string } | { error: string } {
  const key = norm(input);
  if (!key) return { error: 'Enter a team, like "KC", "Chiefs" or "Kansas City Chiefs".' };
  const hit = ALIASES.get(key);
  if (hit) return { abbr: hit };
  if (key === 'newyork' || key === 'ny') return { error: '"New York" could be the Giants (NYG) or the Jets (NYJ).' };
  if (key === 'losangeles' || key === 'la') return { abbr: 'LA' };
  // Partial match on nickname or full name
  const partial = TEAMS.filter((t) => norm(`${t.city} ${t.nickname}`).includes(key) || norm(t.nickname).startsWith(key));
  if (partial.length === 1) return { abbr: partial[0].abbr };
  return { error: `"${input}" isn't an NFL team Gridline recognizes. Use a name like "Chiefs" or a code like "KC". NFL.TEAMS() lists them all.` };
}
