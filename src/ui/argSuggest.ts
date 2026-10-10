// Suggestions for text arguments while typing a formula (stat names, teams, players…).
import { external } from '../core/data/external';
import { currentSeason, playerSeason, type PlayerTable } from '../core/data/nfl';
import { STATS } from '../core/data/nflStats';
import { TEAMS, teamName } from '../core/data/nflTeams';
import type { SuggestKind } from '../core/formula/functions';
import { NFL_GAME_FIELDS, NFL_INFO_FIELDS } from '../core/formula/functions/nfl';
import { STOCK_FIELDS } from '../core/formula/functions/money';
import { COUNTRY_FIELDS, WEATHER_FIELDS } from '../core/formula/functions/world';

export interface ArgSuggestion {
  value: string;
  desc: string;
  kind: string;
}

const CURRENCIES: [string, string][] = [
  ['USD', 'US Dollar'],
  ['EUR', 'Euro'],
  ['GBP', 'British Pound'],
  ['JPY', 'Japanese Yen'],
  ['CAD', 'Canadian Dollar'],
  ['AUD', 'Australian Dollar'],
  ['CHF', 'Swiss Franc'],
  ['CNY', 'Chinese Yuan'],
  ['MXN', 'Mexican Peso'],
  ['INR', 'Indian Rupee'],
  ['BRL', 'Brazilian Real'],
  ['KRW', 'South Korean Won'],
  ['SEK', 'Swedish Krona'],
  ['NOK', 'Norwegian Krone'],
  ['DKK', 'Danish Krone'],
  ['PLN', 'Polish Zloty'],
  ['NZD', 'New Zealand Dollar'],
  ['SGD', 'Singapore Dollar'],
  ['HKD', 'Hong Kong Dollar'],
  ['ZAR', 'South African Rand'],
  ['TRY', 'Turkish Lira'],
];

function playerNames(): string[] {
  const out = new Set<string>();
  const cur = currentSeason();
  for (const y of [cur, cur - 1, cur - 2]) {
    const t = external.peek<PlayerTable>(`nfl:player:REG:${y}`);
    if (t) for (const n of t.names()) out.add(n);
  }
  if (!out.size) {
    // Start downloading recent player lists so names are suggested next time
    playerSeason(cur, 'REG');
    playerSeason(cur - 1, 'REG');
  }
  return [...out];
}

export function argSuggestions(kind: SuggestKind, partial: string): ArgSuggestion[] {
  const q = partial.toLowerCase();
  let all: ArgSuggestion[] = [];
  switch (kind) {
    case 'nflPlayerStat':
    case 'nflTeamStat': {
      const scope = kind === 'nflPlayerStat' ? 'player' : 'team';
      all = STATS.filter((d) => d.scope === 'both' || d.scope === scope).map((d) => ({ value: d.label, desc: d.aliases?.length ? `also: ${d.aliases.slice(0, 3).join(', ')}` : '', kind: d.group }));
      if (scope === 'team' && q.endsWith('allowed')) all = all.filter((s) => !['Record', 'Info'].includes(s.kind)).map((s) => ({ ...s, value: `${s.value} Allowed` }));
      break;
    }
    case 'nflTeam':
      all = TEAMS.map((t) => ({ value: t.nickname, desc: `${t.abbr} · ${teamName(t.abbr)}`, kind: `${t.conf} ${t.div}` }));
      break;
    case 'nflPlayer':
      all = playerNames().map((n) => ({ value: n, desc: '', kind: 'Player' }));
      break;
    case 'nflGameField':
      all = NFL_GAME_FIELDS.map((f) => ({ value: f, desc: '', kind: 'Game' }));
      break;
    case 'nflInfoField':
      all = NFL_INFO_FIELDS.map((f) => ({ value: f, desc: '', kind: 'Profile' }));
      break;
    case 'nflSeasonType':
      all = [
        { value: 'REG', desc: 'Regular season', kind: 'Season' },
        { value: 'POST', desc: 'Playoffs', kind: 'Season' },
        { value: 'ALL', desc: 'Regular season + playoffs', kind: 'Season' },
      ];
      break;
    case 'nflGroup':
      all = [
        { value: 'AFC', desc: 'American Football Conference', kind: 'Conference' },
        { value: 'NFC', desc: 'National Football Conference', kind: 'Conference' },
        ...['AFC', 'NFC'].flatMap((c) => ['East', 'North', 'South', 'West'].map((d) => ({ value: `${c} ${d}`, desc: '', kind: 'Division' }))),
      ];
      break;
    case 'stockField':
      all = STOCK_FIELDS.map((f) => ({ value: f, desc: '', kind: 'Quote' }));
      break;
    case 'currency':
      all = CURRENCIES.map(([c, n]) => ({ value: c, desc: n, kind: 'Currency' }));
      break;
    case 'interval':
      all = ['daily', 'weekly', 'monthly'].map((v) => ({ value: v, desc: '', kind: 'Interval' }));
      break;
    case 'weatherField':
      all = WEATHER_FIELDS.map((f) => ({ value: f, desc: '', kind: 'Weather' }));
      break;
    case 'countryField':
      all = COUNTRY_FIELDS.map((f) => ({ value: f, desc: '', kind: 'Country' }));
      break;
  }
  if (!q) return all.slice(0, 60);
  const starts = all.filter((s) => s.value.toLowerCase().startsWith(q) || s.desc.toLowerCase().startsWith(q));
  const contains = all.filter((s) => !starts.includes(s) && (s.value.toLowerCase().includes(q) || s.desc.toLowerCase().includes(q)));
  return [...starts, ...contains].slice(0, 40);
}
