# Gridline

A dark-mode desktop spreadsheet (Electron + TypeScript) that opens and saves **.xlsx**, **.csv** and **.tsv** files.

## Run

```
npm install
npm run dev      # app with hot reload
npm start        # production build, then launch
npm test         # 82 engine / file-format / NFL / finance tests
npm run dist     # Windows installer + portable exe in release/
```

## Highlights

- 250+ functions (XLOOKUP, FILTER, SORT, UNIQUE, LET, SUMIFS, financial, date, text…) with dynamic-array spilling.
- Errors explain themselves in plain English, point to the cell that caused them, and offer one-click fixes.
- No ribbon: one slim toolbar, plus **Ctrl+K** to search every command, cell, sheet and function.
- Formula editing: colored references, click/drag to insert ranges, autocomplete, argument hints, F4 for `$`.
- Formatting, borders, merges, number formats with live previews, conditional formatting, sort, filter, freeze panes, and native Excel charts.
- Leading zeros and long ID numbers stay as text; CSV export keeps full precision.

## NFL data

Functions that start with `NFL.` pull stats from [nflverse](https://github.com/nflverse) (1999 to today). Files download once and are cached on disk, so they keep working offline. **Ctrl+Alt+F5** refreshes them.

| Function | What it returns |
| --- | --- |
| `NFL.PLAYER(player, stat, [season], [week], [season_type])` | One stat, e.g. `=NFL.PLAYER("Josh Allen", "Passing TDs", 2024)` |
| `NFL.PLAYERSTATS(player, stats, [first], [last])` | A season-by-season table, e.g. `=NFL.PLAYERSTATS("Tom Brady", "Passing TDs, Rushing TDs, Passing Yards", 2015, 2020)` |
| `NFL.GAMELOG(player, [season])` | Week-by-week stats |
| `NFL.TEAM(team, stat, [season], [week])` | Record, points, yards, "… Allowed" stats |
| `NFL.GAME(team, season, week, [field])` | Result, score, spread, over/under… |
| `NFL.SCHEDULE`, `NFL.STANDINGS`, `NFL.LEADERS`, `NFL.ROSTER`, `NFL.PLAYERINFO`, `NFL.TEAMS`, `NFL.STATS` | Tables |

How fill works with these functions:
- Arguments can be ranges. `=NFL.PLAYER("Tom Brady", B1:D1, A2:A7)` fills a seasons × stats grid.
- Dragging a formula down with a typed season (or week) counts it up: 2015 becomes 2016, 2017 and so on.
- The **NFL** toolbar button (or Ctrl+K → "NFL data") builds these grids for you.
- While typing, stat names, teams and players autocomplete inside the quotes.

## Money

- **Loans:** `AMORTIZATION`, `LOANBALANCE`, `PAYOFFMONTHS`, `PAYOFFINTEREST`.
- **Debt:** `DEBTPAYOFF` (avalanche / snowball).
- **Saving:** `SAVINGSGOAL`, `SAVINGS.PROJECTION`, `CAGR`.
- **Live prices:** `STOCK`, `STOCKHISTORY`, `CRYPTO`, `FXRATE`, `FX`.

**File → New from template** has starter workbooks: monthly budget, loan and mortgage planner, debt payoff, net worth and savings, investment portfolio, and an NFL season tracker (schedule, standings, QB comparison, weekly fantasy points).

Gridline-only functions (NFL, STOCK…) are saved into .xlsx files along with their last values. Excel doesn't have these functions, so it may show `#NAME?` in those cells if it recalculates the file.
