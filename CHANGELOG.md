# Changelog

## 1.0.5

- **AI assistant (Beta, Developer branch).** Ask about your workbook or describe a job in plain English, for example: "Find duplicate transactions, but don't delete anything" or "Add a formula that categorizes these transactions". Open it from **Home → Ask AI** or with **Ctrl+Shift+A**.
  - **Works with the whole workbook**, not just what's on screen. It reads the sheets, formulas, charts, highlight rules and dropdowns it needs, and traces how cells depend on each other across sheets. Every read is listed in the chat.
  - **Nothing changes without your approval.** Changes come as proposals that show every cell that changes (before → after, with the formula), plus formula results elsewhere that would change. Click **Apply** or **Reject**, or tell it what to do instead.
  - **One-click undo.** Each applied change is a single step with an **Undo** button. The button still works after you've made other edits, and keeps any cells you changed yourself since.
  - Use **Claude** with your own API key, or a **model on your computer** with Ollama or LM Studio, so the workbook never leaves it. Choose in **Settings → AI assistant**.
  - **Your API key stays secret.** It's encrypted with your computer's keychain and never shown again. Only Gridline's background process uses it, and only to talk to Anthropic. Keys pasted into the chat aren't sent, and keys found in a workbook are hidden from the AI.
  - Cell addresses in answers are links that jump to those cells.

## 1.0.4

- **Calculations match Excel exactly.** Tested against an Excel workbook with 9,719 formulas: every result is identical to Excel's, down to the last digit, and all 19,638 cells display the same text. Gridline now:
  - adds up SUM, SUMIFS and SUMPRODUCT in the same order as Excel;
  - rounds the last digit of multiplications and divisions like Excel;
  - returns exactly 0 when a final + or − lands a hair away from zero (=0.1+0.2-0.3 is 0, like Excel);
  - keeps 15 significant digits of numbers you type, like Excel;
  - rounds displayed values like Excel, so $1,087.005 shows as $1,087.01.
- **Data validation.** **Data → Data validation** limits what can be typed in cells:
  - whole numbers, decimals, dates, times, text length, dropdown lists or a custom formula;
  - an input message and a Stop, Warning or Information alert;
  - **Circle invalid data** shows cells that break a rule.
  - Rules from Excel files are kept and saved.
- **Excel files:**
  - Fixed workbooks that Excel couldn't open after saving in Gridline. This affected a chart on a sheet with data bars, formulas using LET, and spill references like A1#.
  - LAMBDA and other recent Excel functions keep working in Excel after you save.
  - Formulas that some apps save in a way Excel shows as #NAME? (CONCAT, LET, XLOOKUP) now calculate.

## 1.0.3

- **Macros.** Gridline now runs VBA macros, the same kind Excel uses. Open the new **Automate** tab to run a macro, or open the **macro editor** (Alt+F11) to write one. The editor sits under the grid, so you can see your sheet while you work. It has syntax colors, red underlines for mistakes, and F5 to run. Most everyday Excel macros work: Range, Cells, Worksheets, loops, If, MsgBox, InputBox, WorksheetFunction, sorting, finding, formatting and more. A whole macro run undoes with one Ctrl+Z.
- **.xlsm files.** Macro-enabled workbooks open with their code, and the code is saved back so it keeps working in Excel. A workbook with code saved as .xlsx gets an offer to save as .xlsm instead.
- **Step macros (Beta, Developer branch).** Build a macro without writing code, from plain-English steps such as "Color B2:B50 that is more than 100 in Yellow" or "For each row… If D{Row} is Paid…". Start from one of five examples, and click **Show code** to see the VBA the steps turn into.
- **Run macros your way.** Put a button on the sheet, give a macro a keyboard shortcut, or run it automatically when the file opens, when cells change, when a sheet is shown or before saving. VBA events like Workbook_Open and Worksheet_Change work too.
- **Safe by default.** Macros in files you open stay off until you click **Enable macros**, and **Settings → Macros** controls this. Macros can only work with the workbook: they can't open files, run programs or reach the internet.
- Clearer messages when checking for updates fails (for example, no internet connection).
- Releases are now published in a single step, so each version appears exactly once on GitHub.

## 1.0.2

- **Home page.** Gridline now opens to a Home page where you can start a blank workbook, pick a preset, or reopen a recent file. Pin the files you use most and search the list. The **File** tab (Ctrl+Shift+H) brings you back to it any time, together with Save, Save as and Export.
- **10 presets with live previews.** Monthly budget, loan and mortgage calculator, debt payoff, net worth and savings, stock portfolio, crypto portfolio, NFL team tracker, fantasy football, trip planner and project tracker. The stock, crypto, NFL, fantasy and trip presets are new or rebuilt.
- **New Live Data tab.** Finance, Sports, and Weather & world data in one place. It replaces the "Money & NFL" tab.
- **Weather, holidays and country data.**
  - `WEATHER` and `WEATHER.FORECAST` give current conditions and up to 16 days of forecast for any city.
  - `HOLIDAYS` lists public holidays for 100+ countries.
  - `COUNTRYDATA` returns population, GDP, inflation and more from the World Bank.
- **Developer branch.** Turn it on in Settings → Updates to get development builds with the newest features before they're released.
- **Documentation.** A user guide, a function reference and install instructions, all linked from the README.
- **Fixes:** chart axis labels at the right edge are no longer cut off.

## 1.0.1

- **Ribbon with tabs.** Home, Insert, Formulas, Data, View, Money & NFL and Help. Every tab is organized into small groups with plain-English labels. Drag the bottom edge of the ribbon (or double-click a tab) to switch between full, compact and tabs-only sizes.
- **Settings.** Click the Gridline logo, then Settings (or press Ctrl+,). More than 120 options in 13 sections, with search: theme and accent color, density, what Enter does, date order, autosave, CSV options, live-data refresh, chart colors, status-bar totals and more.
- **Advanced format.** Select a table and click Advanced format (Home tab). Gridline works out headers, dates, money, percentages, years and IDs, then formats everything in one step. Keep it, try another style, or revert.
- **Automatic updates.** Gridline checks for new versions when it starts and asks before installing.
- **Quick chart (F11)**, cell styles (Title, Heading, Total, Good, Bad, Input…), a row and column highlight for wide tables, and an option to hide zero values.

## 1.0.0

- First release: .xlsx, .csv and .tsv files, 250+ functions, charts, conditional formatting, sort, filter and freeze panes.
- NFL.* functions with live nflverse data, personal-finance functions (loans, debt payoff, savings, live stock, crypto and currency prices) and starter templates.
