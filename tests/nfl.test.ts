import { beforeEach, describe, expect, it } from 'vitest';
import { parseRange } from '../src/core/address';
import { external, FetchError } from '../src/core/data/external';
import { autofill, bumpSeriesArgs } from '../src/core/ops/fill';
import { ErrorValue } from '../src/core/values';
import { makeDoc, val } from './helpers';

const R = 'https://github.com/nflverse/nflverse-data/releases/download';
const PLAYER_COLS =
  'player_id,player_name,player_display_name,position,position_group,season,season_type,recent_team,games,completions,attempts,passing_yards,passing_tds,passing_interceptions,sacks_suffered,carries,rushing_yards,rushing_tds,receptions,targets,receiving_yards,receiving_tds,def_sacks,def_interceptions,def_tackles_solo,def_tackle_assists,fantasy_points,fantasy_points_ppr';

function seasonFile(season: number, rows: string[]): string {
  return [PLAYER_COLS, ...rows.map((r) => r.replace('{S}', String(season)))].join('\n');
}

const brady = (team: string, tds: number, rtd: number, yds: number) =>
  `00-0019596,T.Brady,Tom Brady,QB,QB,{S},REG,${team},16,400,600,${yds},${tds},10,20,30,10,${rtd},0,0,0,0,0,0,0,0,300,300`;

const FILES: Record<string, string> = {
  [`${R}/stats_player/stats_player_reg_2015.csv`]: seasonFile(2015, [brady('NE', 36, 3, 4770)]),
  [`${R}/stats_player/stats_player_reg_2016.csv`]: seasonFile(2016, [brady('NE', 28, 0, 3554)]),
  [`${R}/stats_player/stats_player_reg_2017.csv`]: seasonFile(2017, [brady('NE', 32, 0, 4577)]),
  [`${R}/stats_player/stats_player_reg_2018.csv`]: seasonFile(2018, [brady('NE', 29, 2, 4355)]),
  [`${R}/stats_player/stats_player_reg_2019.csv`]: seasonFile(2019, [brady('NE', 24, 3, 4057)]),
  [`${R}/stats_player/stats_player_reg_2020.csv`]: seasonFile(2020, [
    brady('TB', 40, 3, 4633),
    '00-0034857,J.Allen,Josh Allen,QB,QB,2020,REG,BUF,16,396,572,4544,37,10,26,102,421,8,1,1,12,1,0,0,0,0,395.6,396.6',
    '00-0035000,J.Allen,Josh Allen,LB,LB,2020,REG,JAX,16,0,0,0,0,0,0,0,0,0,0,0,0,0,2.5,1,60,30,0,0',
    '00-0033873,P.Mahomes,Patrick Mahomes,QB,QB,2020,REG,KC,15,390,588,4740,38,6,22,62,308,2,0,0,0,0,0,0,0,0,374.0,374.0',
  ]),
  [`${R}/stats_player/stats_player_week_2020.csv`]: [
    'player_id,player_name,player_display_name,position,position_group,season,week,season_type,game_id,team,opponent_team,completions,attempts,passing_yards,passing_tds,passing_interceptions,carries,rushing_yards,rushing_tds,fantasy_points_ppr',
    '00-0033873,P.Mahomes,Patrick Mahomes,QB,QB,2020,1,REG,2020_01_HOU_KC,KC,HOU,24,32,211,3,0,7,25,1,27.9',
    '00-0033873,P.Mahomes,Patrick Mahomes,QB,QB,2020,2,REG,2020_02_KC_LAC,KC,LAC,26,40,302,2,0,6,54,0,25.5',
  ].join('\n'),
  'https://github.com/nflverse/nfldata/raw/master/data/games.csv': [
    'game_id,season,game_type,week,gameday,weekday,gametime,away_team,away_score,home_team,home_score,location,result,total,overtime,spread_line,total_line,away_moneyline,home_moneyline,div_game,roof,surface,temp,wind,away_qb_name,home_qb_name,away_coach,home_coach,stadium',
    '2020_01_HOU_KC,2020,REG,1,2020-09-10,Thursday,20:20,HOU,20,KC,34,Home,14,54,0,9.5,53.5,350,-450,0,outdoors,grass,,,Deshaun Watson,Patrick Mahomes,Bill O\'Brien,Andy Reid,Arrowhead Stadium',
    '2020_02_KC_LAC,2020,REG,2,2020-09-20,Sunday,16:25,KC,23,LAC,20,Home,-3,43,1,-8.5,46,-400,320,1,outdoors,grass,,,Patrick Mahomes,Tyrod Taylor,Andy Reid,Anthony Lynn,SoFi Stadium',
    '2020_03_KC_BAL,2020,REG,3,2020-09-28,Monday,20:15,KC,34,BAL,20,Home,-14,54,0,3.5,54.5,140,-160,0,outdoors,grass,,,Patrick Mahomes,Lamar Jackson,Andy Reid,John Harbaugh,M&T Bank Stadium',
    '2015_01_PIT_NE,2015,REG,1,2015-09-10,Thursday,20:30,PIT,21,NE,28,Home,7,49,0,7,51,250,-300,0,outdoors,turf,,,Ben Roethlisberger,Tom Brady,Mike Tomlin,Bill Belichick,Gillette Stadium',
    '2015_02_NE_BUF,2015,REG,2,2015-09-20,Sunday,13:00,NE,40,BUF,32,Home,-8,72,0,-1,45,-120,100,1,outdoors,turf,,,Tom Brady,Tyrod Taylor,Bill Belichick,Rex Ryan,Ralph Wilson Stadium',
  ].join('\n'),
  [`${R}/stats_team/stats_team_reg_2020.csv`]: ['season,team,season_type,games,passing_yards,rushing_yards,passing_interceptions,def_sacks,def_interceptions,fumble_recovery_opp,rushing_fumbles_lost', '2020,KC,REG,16,4854,1765,6,32,10,6,3', '2020,BUF,REG,16,4620,1723,11,38,15,11,4'].join('\n'),
  [`${R}/stats_team/stats_team_week_2020.csv`]: [
    'season,week,team,season_type,game_id,opponent_team,passing_yards,rushing_yards',
    '2020,1,KC,REG,2020_01_HOU_KC,HOU,211,166',
    '2020,1,HOU,REG,2020_01_HOU_KC,KC,253,54',
    '2020,2,KC,REG,2020_02_KC_LAC,LAC,302,132',
    '2020,2,LAC,REG,2020_02_KC_LAC,KC,244,153',
  ].join('\n'),
  [`${R}/players/players.csv`]: ['gsis_id,display_name,position,position_group,latest_team,status,rookie_season,last_season,college_name,height,weight,birth_date,jersey_number', '00-0019596,Tom Brady,QB,QB,TB,RET,2000,2022,Michigan,76,225,1977-08-03,12'].join('\n'),
  'https://query1.finance.yahoo.com/v8/finance/chart/VOO?range=1d&interval=1d': JSON.stringify({ chart: { result: [{ meta: { regularMarketPrice: 500, chartPreviousClose: 490, longName: 'Vanguard S&P 500 ETF', currency: 'USD' } }] } }),
  'https://api.frankfurter.dev/v1/latest?base=USD&symbols=EUR': JSON.stringify({ rates: { EUR: 0.9 } }),
};

let fetches: string[] = [];

beforeEach(() => {
  external.reset();
  fetches = [];
  external.fetcher = async (url) => {
    fetches.push(url);
    const t = FILES[url];
    if (t === undefined) throw new FetchError('Not Found', 404);
    return t;
  };
});

const settle = () => new Promise((r) => setTimeout(r, 120));

describe('NFL functions', () => {
  it('shows #BUSY! while loading, then the value', async () => {
    const d = makeDoc({ A1: '=NFL.PLAYER("Tom Brady","Passing TDs",2020)' });
    expect(val(d, 'A1')).toBe('#BUSY!');
    await settle();
    expect(val(d, 'A1')).toBe(40);
  });

  it('accepts a skipped week argument', async () => {
    const d = makeDoc({ A1: 'Tom Brady', B1: 'Passing TDs', A2: '2020', B2: '=NFL.PLAYER($A$1, B$1, $A2, , "REG")', A3: '=A2+1' });
    await settle();
    expect(val(d, 'B2')).toBe(40);
  });

  it('fills a seasons × stats table from ranges', async () => {
    const d = makeDoc({
      B1: 'Passing TDs',
      C1: 'Rushing TDs',
      D1: 'Passing Yards',
      A2: '2015',
      A3: '2016',
      A4: '2017',
      A5: '2018',
      A6: '2019',
      A7: '2020',
      B2: '=NFL.PLAYER("Tom Brady",B1:D1,A2:A7)',
    });
    await settle();
    expect([val(d, 'B2'), val(d, 'C2'), val(d, 'D2')]).toEqual([36, 3, 4770]);
    expect([val(d, 'B7'), val(d, 'C7'), val(d, 'D7')]).toEqual([40, 3, 4633]);
  });

  it('builds a career table', async () => {
    const d = makeDoc({ A1: '=NFL.PLAYERSTATS("tom brady", "Passing TDs, Rushing TDs, Passing Yards", 2015, 2020)' });
    await settle();
    expect([val(d, 'A1'), val(d, 'B1'), val(d, 'C1'), val(d, 'E1')]).toEqual(['Season', 'Team', 'Passing TDs', 'Passing Yards']);
    expect([val(d, 'A2'), val(d, 'B2'), val(d, 'C2')]).toEqual([2015, 'NE', 36]);
    expect([val(d, 'A7'), val(d, 'B7'), val(d, 'E7')]).toEqual([2020, 'TB', 4633]);
  });

  it('steps the season when dragging down', async () => {
    expect(bumpSeriesArgs('NFL.PLAYER("Tom Brady","Passing TDs",2015)', 2)).toBe('NFL.PLAYER("Tom Brady","Passing TDs",2017)');
    expect(bumpSeriesArgs('NFL.PLAYER("X","Passing Yards",2020,3)', 1)).toBe('NFL.PLAYER("X","Passing Yards",2020,4)');
    expect(bumpSeriesArgs('SUM(2015)', 1)).toBe('SUM(2015)');
    const d = makeDoc({ A1: '=NFL.PLAYER("Tom Brady","Passing TDs",2015)' });
    const s = d.wb.activeSheet;
    d.transact('fill', (tx) => autofill(tx, d, s, parseRange('A1')!, parseRange('A1:A6')!, 'auto'));
    expect(s.getCell(5, 0)?.f).toBe('NFL.PLAYER("Tom Brady","Passing TDs",2020)');
    await settle();
    expect([val(d, 'A1'), val(d, 'A3'), val(d, 'A6')]).toEqual([36, 32, 40]);
  });

  it('tells players with the same name apart', async () => {
    const d = makeDoc({ A1: '=NFL.PLAYER("Josh Allen","Position",2020)', A2: '=NFL.PLAYER("Josh Allen (JAX)","Sacks",2020)', A3: '=NFL.PLAYER("Josh Allen","Sacks",2020)' });
    await settle();
    expect(val(d, 'A1')).toBe('QB');
    expect(val(d, 'A2')).toBe(2.5);
    expect(val(d, 'A3')).toBe(26); // a QB's "sacks" are times sacked
  });

  it('reads weekly stats and game logs', async () => {
    const d = makeDoc({ A1: '=NFL.PLAYER("Patrick Mahomes","Passing Yards",2020,2)', A2: '=NFL.PLAYER("Patrick Mahomes","Passing Yards",2020,5)', C1: '=NFL.GAMELOG("Patrick Mahomes",2020)' });
    await settle();
    expect(val(d, 'A1')).toBe(302);
    expect(val(d, 'A2')).toBe('#N/A');
    expect([val(d, 'C1'), val(d, 'D2'), val(d, 'E2'), val(d, 'E3')]).toEqual(['Week', 'vs HOU', 'W 34-20', 'W 23-20 OT']);
  });

  it('computes team records, stats and allowed stats', async () => {
    const d = makeDoc({
      A1: '=NFL.TEAM("Chiefs","Record",2020)',
      A2: '=NFL.TEAM("KC","Points For",2020)',
      A3: '=NFL.TEAM("Kansas City Chiefs","Passing Yards",2020)',
      A4: '=NFL.TEAM("KC","Rushing Yards Allowed",2020)',
      A5: '=NFL.TEAM("KC","Turnovers",2020)',
      A6: '=NFL.TEAM("Patriots","Team Name",2015)',
      A7: '=NFL.TEAM("KC","Wins",2020,2)',
    });
    await settle();
    expect(val(d, 'A1')).toBe('3-0');
    expect(val(d, 'A2')).toBe(91);
    expect(val(d, 'A3')).toBe(4854);
    expect(val(d, 'A4')).toBe(54 + 153);
    expect(val(d, 'A5')).toBe(9);
    expect(val(d, 'A6')).toBe('New England Patriots');
    expect(val(d, 'A7')).toBe(1);
  });

  it('describes games, schedules and standings', async () => {
    const d = makeDoc({
      A1: '=NFL.GAME("KC",2020,1)',
      A2: '=NFL.GAME("KC",2020,1,"Spread")',
      A3: '=NFL.GAME("KC",2020,1,"Covered")',
      A4: '=NFL.GAME("KC",2020,2,"Matchup")',
      A5: '=NFL.GAME("KC",2020,4)',
      C1: '=NFL.SCHEDULE("KC",2020)',
      L1: '=NFL.STANDINGS(2020,"AFC West")',
    });
    await settle();
    expect(val(d, 'A1')).toBe('W 34-20');
    expect(val(d, 'A2')).toBe(-9.5);
    expect(val(d, 'A3')).toBe('Yes');
    expect(val(d, 'A4')).toBe('@ LAC');
    expect(val(d, 'A5')).toBe('#N/A');
    expect([val(d, 'C2'), val(d, 'E2'), val(d, 'F2')]).toEqual([1, 'vs HOU', 'W']);
    expect([val(d, 'M2'), val(d, 'N2')]).toEqual(['Kansas City Chiefs', 3]);
  });

  it('explains mistakes', async () => {
    const d = makeDoc({ A1: '=NFL.PLAYER("Tom Bradyy","Passing TDs",2020)', A2: '=NFL.PLAYER("Tom Brady","Pasing TDz",2020)', A3: '=NFL.TEAM("Bananas","Wins",2020)', A4: '=NFL.PLAYER("Tom Brady","Passing TDs",2031)' });
    await settle();
    const e1 = d.engine.getValue(d.wb.activeSheet, 0, 0) as ErrorValue;
    expect(e1.code).toBe('#N/A');
    expect(e1.message).toContain('Did you mean Tom Brady');
    const e2 = d.engine.getValue(d.wb.activeSheet, 1, 0) as ErrorValue;
    expect(e2.code).toBe('#VALUE!');
    expect(e2.message).toContain('Passing TDs');
    expect(val(d, 'A3')).toBe('#VALUE!');
    expect(val(d, 'A4')).toBe('#N/A');
  });

  it('downloads each file only once', async () => {
    makeDoc({ A1: '=NFL.PLAYER("Tom Brady","Passing TDs",2020)', A2: '=NFL.PLAYER("Tom Brady","Passing Yards",2020)', A3: '=NFL.PLAYER("Josh Allen","Rushing Yards",2020)' });
    await settle();
    expect(fetches.filter((u) => u.includes('stats_player_reg_2020'))).toHaveLength(1);
  });

  it('reads player profiles and careers', async () => {
    const d = makeDoc({ A1: '=NFL.PLAYERINFO("Tom Brady","College")', A2: '=NFL.PLAYERINFO("Tom Brady","Height")', C1: '=NFL.PLAYERSTATS("Tom Brady","Passing TDs")' });
    await settle();
    await settle();
    expect(val(d, 'A1')).toBe('Michigan');
    expect(val(d, 'A2')).toBe(`6'4"`);
    // Career 2000-2022: rows for seasons without data are blank
    expect(val(d, 'C2')).toBe(2000);
  });
});

describe('money functions', () => {
  it('builds an amortization schedule', () => {
    const d = makeDoc({ A1: '=AMORTIZATION(6%, 30, 200000, DATE(2026,1,1))', J1: '=PMT(6%/12,360,-200000)' });
    expect(val(d, 'C2')).toBeCloseTo(val(d, 'J1') as number, 1);
    expect(val(d, 'G361')).toBe(0);
    expect(d.engine.spillFormatAt(d.wb.activeSheet, 1, 1)).toBe('mmm d, yyyy');
  });

  it('extra payments shorten the loan', () => {
    const d = makeDoc({ A1: '=ROWS(AMORTIZATION(6%, 30, 200000, , 200))-1' });
    expect(val(d, 'A1') as number).toBeLessThan(300);
  });

  it('computes balances, payoff and savings', () => {
    const d = makeDoc({
      A1: '=LOANBALANCE(6%,30,200000,360)',
      A2: '=PAYOFFMONTHS(24%,5000,200)',
      A3: '=PAYOFFMONTHS(24%,5000,50)',
      A4: '=SAVINGSGOAL(12000,1,0%)',
      A5: '=ROUND(CAGR(10000,20000,10),4)',
      A6: '=PAYOFFINTEREST(0%,1000,100)',
    });
    expect(val(d, 'A1')).toBe(0);
    expect(val(d, 'A2')).toBe(36);
    expect(val(d, 'A3')).toBe('#NUM!');
    expect(val(d, 'A4')).toBe(1000);
    expect(val(d, 'A5')).toBe(0.0718);
    expect(val(d, 'A6')).toBe(0);
  });

  it('plans debt payoff', () => {
    const d = makeDoc({
      A1: 'Card',
      A2: 'Car',
      B1: '3000',
      B2: '9000',
      C1: '22%',
      C2: '6%',
      D1: '90',
      D2: '250',
      F1: '=DEBTPAYOFF(A1:A2,B1:B2,C1:C2,D1:D2,600,"avalanche")',
      J1: '=DEBTPAYOFF(A1:A2,B1:B2,C1:C2,D1:D2,600,"snowball")',
    });
    const avalanche = val(d, 'H4') as number;
    const snowball = val(d, 'L4') as number;
    expect(val(d, 'F4')).toBe('Total');
    expect(avalanche).toBeLessThanOrEqual(snowball);
    expect(val(d, 'G2') as number).toBeLessThan(val(d, 'G3') as number);
  });

  it('projects savings growth', () => {
    const d = makeDoc({ A1: '=SAVINGS.PROJECTION(0, 100, 0%, 2)' });
    expect([val(d, 'A3'), val(d, 'B3'), val(d, 'D3')]).toEqual([2, 2400, 2400]);
  });

  it('gets live quotes and exchange rates', async () => {
    const d = makeDoc({ A1: '=STOCK("voo")', A2: '=STOCK("VOO","Change %")', A3: '=FX(100,"USD","EUR")', A4: '=STOCK("ZZZZ")' });
    await settle();
    expect(val(d, 'A1')).toBe(500);
    expect(val(d, 'A2') as number).toBeCloseTo(10 / 490);
    expect(val(d, 'A3')).toBe(90);
    expect(val(d, 'A4')).toBe('#N/A');
  });
});
