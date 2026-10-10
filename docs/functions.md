# Gridline function reference

Gridline supports the standard spreadsheet functions you know from Excel (250+, including XLOOKUP, FILTER, SORT, UNIQUE, LET and SUMIFS). To browse them, use **Insert → Function** or the category menus on the **Formulas** tab.

This page covers the functions that are **unique to Gridline**:

- [Personal finance](#personal-finance): loans, debt, savings.
- [Live prices](#live-prices): stocks, crypto, currencies.
- [Sports (NFL)](#sports-nfl): players, teams, games, standings.
- [Weather and world](#weather-and-world): weather, holidays, country data.

Arguments in `[brackets]` are optional. Text arguments go in quotes, and while you type inside the quotes Gridline suggests valid values (stat names, teams, fields, currencies…).

> **Sharing files with Excel users:** these functions are saved in `.xlsx` files together with their last results, so the numbers still show up. Excel doesn't know these functions, though, and may show `#NAME?` if it recalculates the file.

## How live data works

- The first time a formula needs data, the cell shows **Loading…** while it downloads, then fills in.
- Downloaded data is saved on your computer, so it keeps working offline.
- Prices refresh after a couple of minutes. Current NFL stats refresh every few hours. Older seasons are kept for weeks.
- **Live Data → Refresh all** (Ctrl+Alt+F5) downloads everything again.
- Errors you may see:
  - `#CONNECT!`: the data couldn't be downloaded, usually because there's no internet connection.
  - `#N/A`: the data doesn't exist, for example a game that hasn't been played yet or a ticker that isn't real.

## Personal finance

Rates can be typed as `6.5%`, `0.065` or `6.5`. Gridline treats any rate of 1 or more as a percentage.

| Function | Returns | Example |
| --- | --- | --- |
| `AMORTIZATION(annual_rate, years, loan_amount, [start_date], [extra_payment])` | A month-by-month table: payment, principal, interest, extra and remaining balance. | `=AMORTIZATION(6.5%, 30, 350000, DATE(2026,1,1), 200)` |
| `LOANBALANCE(annual_rate, years, loan_amount, payments_made)` | What you still owe after a number of payments. | `=LOANBALANCE(6.5%, 30, 350000, 60)` |
| `PAYOFFMONTHS(annual_rate, balance, monthly_payment)` | Months until a balance is paid off. | `=PAYOFFMONTHS(24.99%, 5000, 200)` |
| `PAYOFFINTEREST(annual_rate, balance, monthly_payment)` | Total interest paid until the balance is gone. | `=PAYOFFINTEREST(24.99%, 5000, 200)` |
| `DEBTPAYOFF(names, balances, annual_rates, min_payments, monthly_budget, [method])` | When each debt is paid off and the interest paid. `method` is `"avalanche"` (highest rate first, the default) or `"snowball"` (smallest balance first). | `=DEBTPAYOFF(A2:A5, B2:B5, C2:C5, D2:D5, 1000, "snowball")` |
| `SAVINGSGOAL(goal, years, annual_rate, [current_savings])` | How much to save each month to reach a goal. | `=SAVINGSGOAL(25000, 3, 4.5%, 2000)` |
| `SAVINGS.PROJECTION(starting_balance, monthly_contribution, annual_rate, years)` | A year-by-year table: contributed, growth and balance. | `=SAVINGS.PROJECTION(10000, 500, 7%, 30)` |
| `CAGR(start_value, end_value, years)` | Compound annual growth rate. | `=CAGR(10000, 18000, 5)` returns about 12.5% |

## Live prices

| Function | Returns | Example |
| --- | --- | --- |
| `STOCK(ticker, [field])` | A stock, ETF or index quote, about 15 minutes delayed. Fields: Price (default), Change, Change %, Previous Close, Day High, Day Low, Volume, 52 Week High, 52 Week Low, Name, Currency, Exchange, Updated. | `=STOCK("AAPL")`, `=STOCK("VOO", "Change %")` |
| `STOCKHISTORY(ticker, start_date, [end_date], [interval])` | A table of Date, Open, High, Low, Close and Volume. `interval` is `"daily"` (default), `"weekly"` or `"monthly"`. | `=STOCKHISTORY("SPY", TODAY()-90)` |
| `CRYPTO(symbol, [currency], [date])` | A cryptocurrency price, now or on a past date. The currency defaults to the one in Settings → Live data. | `=CRYPTO("BTC")`, `=CRYPTO("ETH", "EUR", DATE(2025,1,1))` |
| `FXRATE(from_currency, to_currency, [date])` | An exchange rate (European Central Bank reference rates). | `=FXRATE("USD", "EUR")` |
| `FX(amount, from_currency, to_currency, [date])` | Converts an amount between currencies. | `=FX(100, "USD", "JPY")` |

## Sports (NFL)

The data comes from [nflverse](https://github.com/nflverse) and covers 1999 to today, regular season and playoffs.

- **Players:** use the name as it's usually written: `"Patrick Mahomes"`, `"Ja'Marr Chase"`. If two players share a name, add the team or position: `"Josh Allen (BUF)"`.
- **Teams:** a nickname, code or full name all work: `"Chiefs"`, `"KC"`, `"Kansas City Chiefs"`.
- **Seasons and weeks:**
  - `season` defaults to the current season.
  - `week` is a number. For playoffs you can also use `"WC"`, `"DIV"`, `"CON"` or `"SB"`.
  - `season_type` is `"REG"` (default), `"POST"` or `"ALL"`.
- **Stats:** use friendly names like `"Passing Yards"`, `"Rushing TDs"`, `"Receptions"`, `"Fantasy Points PPR"`, `"Passer Rating"`, `"Sacks"` or `"FG %"`. `=NFL.STATS()` lists all of them. For team stats you can add **Allowed**, as in `"Passing Yards Allowed"`.

| Function | Returns | Example |
| --- | --- | --- |
| `NFL.PLAYER(player, stat, [season], [week], [season_type])` | One stat for a season, or for one week. | `=NFL.PLAYER("Josh Allen", "Passing TDs", 2024)` |
| `NFL.PLAYERSTATS(player, stats, [first_season], [last_season], [season_type])` | A season-by-season table with one column per stat. | `=NFL.PLAYERSTATS("Tom Brady", "Passing TDs, Passing Yards", 2015, 2020)` |
| `NFL.GAMELOG(player, [season], [stats], [season_type])` | Week-by-week stats with opponent and result. | `=NFL.GAMELOG("Saquon Barkley", 2024)` |
| `NFL.TEAM(team, stat, [season], [week], [season_type])` | A team stat or record: Record, Wins, Win %, Points For, Point Differential, Passing Yards, Takeaways… | `=NFL.TEAM("Lions", "Record", 2023)` |
| `NFL.GAME(team, season, week, [field])` | One game. Fields include Result, Score, Opponent, Spread, Covered, Over/Under, Kickoff, QB, Coach. | `=NFL.GAME("Chiefs", 2023, "SB", "Score")` |
| `NFL.SCHEDULE(team, [season], [season_type])` | The full schedule with results, scores and betting lines (bye weeks included). | `=NFL.SCHEDULE("Bills")` |
| `NFL.STANDINGS([season], [group])` | Standings for the league, a conference (`"AFC"`) or a division (`"NFC North"`). | `=NFL.STANDINGS(2024, "NFC North")` |
| `NFL.LEADERS(stat, [season], [count], [position], [season_type])` | The top players for a stat. | `=NFL.LEADERS("Rushing Yards", 2024, 10)` |
| `NFL.ROSTER(team, [season], [position])` | A team's players. | `=NFL.ROSTER("Eagles", 2024, "WR")` |
| `NFL.PLAYERINFO(player, field)` | Profile details: Position, Team, College, Height, Weight, Age, Draft Year, Draft Pick… | `=NFL.PLAYERINFO("Travis Kelce", "College")` |
| `NFL.TEAMS([season])` | All 32 teams with codes, names, conference and division. | `=NFL.TEAMS()` |
| `NFL.STATS([kind])` | Every stat name you can use (`"player"` or `"team"`). | `=NFL.STATS("team")` |

**Filling:**
- Ranges fill whole tables. `=NFL.PLAYER("Tom Brady", B1:D1, A2:A7)` gives seasons down the side and stats across.
- When you drag a formula that has a typed season or week, the number counts up: 2015, 2016, 2017…
- **Live Data → NFL data** builds these grids for you.

## Weather and world

| Function | Returns | Example |
| --- | --- | --- |
| `WEATHER(location, [field], [day])` | Current weather or a forecast for up to 15 days ahead. Fields: Temperature (default), Feels like, Condition, High, Low, Rain chance, Precipitation, Humidity, Wind, UV index, Sunrise, Sunset, Place. `day` is 0 for today, 1 for tomorrow, and so on. | `=WEATHER("Chicago")`, `=WEATHER("Paris, France", "High", 1)` |
| `WEATHER.FORECAST(location, [days])` | A forecast table: date, conditions, high, low, chance of rain and precipitation. | `=WEATHER.FORECAST("Denver", 7)` |
| `HOLIDAYS(country, [year])` | Public holidays: date, name, local name, and whether they're nationwide. | `=HOLIDAYS("Germany", 2026)` |
| `COUNTRYDATA(country, field, [year])` | World Bank facts. Fields: Population, GDP, GDP per person, GDP growth, Inflation, Unemployment, Life expectancy, Internet users, Land area, Urban population, Capital, Region, Income level. Without a year you get the latest value. | `=COUNTRYDATA("Japan", "Population")` |

- **Locations** can be a city (`"Chicago"`), a city with its state or country (`"Springfield, IL"`, `"Paris, France"`), or coordinates (`"41.88, -87.63"`).
- **Units** follow your region. You can change them in **Settings → Live data → Units**.
