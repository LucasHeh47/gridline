# Gridline

A dark-mode desktop spreadsheet (Electron + TypeScript) that opens and saves **.xlsx**, **.csv** and **.tsv** files.

## Run

```
npm install
npm run dev      # app with hot reload
npm start        # production build, then launch
npm test         # 86 engine / file-format / NFL / finance / formatting tests
npm run dist     # Windows installer + portable exe in release/
npm run release  # build and publish a GitHub release (needs GH_TOKEN)
```

## Highlights

- 250+ functions (XLOOKUP, FILTER, SORT, UNIQUE, LET, SUMIFS, financial, date, text…) with dynamic-array spilling.
- Errors explain themselves in plain English, point to the cell that caused them, and offer one-click fixes.
- A ribbon with plain-English tabs (Home, Insert, Formulas, Data, View, Money & NFL, Help). Drag its bottom edge, or double-click a tab, to switch between full, compact and tabs-only. **Ctrl+K** searches every command, cell, sheet and function.
- **Settings** (click the Gridline logo, or Ctrl+,): 125 options in 13 sections, with search, export and import.
- **Advanced format** (Home tab, Ctrl+Alt+F): detects headers, totals, dates, money, percentages, years and IDs and formats the whole table in one step. You can keep it, try another style, or revert.
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
- **Insert → NFL data** (or Ctrl+K → "NFL data") builds these grids for you.
- While typing, stat names, teams and players autocomplete inside the quotes.

## Money

- **Loans:** `AMORTIZATION`, `LOANBALANCE`, `PAYOFFMONTHS`, `PAYOFFINTEREST`.
- **Debt:** `DEBTPAYOFF` (avalanche / snowball).
- **Saving:** `SAVINGSGOAL`, `SAVINGS.PROJECTION`, `CAGR`.
- **Live prices:** `STOCK`, `STOCKHISTORY`, `CRYPTO`, `FXRATE`, `FX`.

**File → New from template** has starter workbooks: monthly budget, loan and mortgage planner, debt payoff, net worth and savings, investment portfolio, and an NFL season tracker (schedule, standings, QB comparison, weekly fantasy points).

Gridline-only functions (NFL, STOCK…) are saved into .xlsx files along with their last values. Excel doesn't have these functions, so it may show `#NAME?` in those cells if it recalculates the file.

## Versions and updates

Every change ships as a new version (1.0.1, 1.0.2…), recorded in [CHANGELOG.md](CHANGELOG.md). Installed copies check GitHub Releases when they start and ask before updating. You can turn this off in Settings → Updates. The portable exe can't replace itself, so it opens the download page instead.

To publish a release:

1. Bump `version` in `package.json` and add a section to `CHANGELOG.md`.
2. Commit, then tag and push: `git tag v1.0.2 && git push origin main --tags`.
3. The **Release** GitHub Action builds the installer and publishes it, along with `latest.yml`, which installed copies read.

You can also publish from this computer with `npm run release`. Set `GH_TOKEN` to a GitHub token with access to the repository first.

To test the update prompt without publishing, serve a folder containing `latest.yml` and the installer locally, then start the packaged app with `GRIDLINE_UPDATE_FEED=http://127.0.0.1:8123/`. Only localhost addresses are accepted.
