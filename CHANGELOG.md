# Changelog

## 1.0.6

- **AI assistant (Beta, Developer branch).** Ask about your workbook or describe a job in plain English, for example: "Find duplicate transactions, but don't delete anything" or "Add a formula that categorizes these transactions". Open it from **Home → Ask AI** or with **Ctrl+Shift+A**.
  - **Works with the whole workbook**, not just what's on screen. It reads the sheets, formulas, charts, highlight rules and dropdowns it needs, and traces how cells depend on each other across sheets. Every read is listed in the chat.
  - **See changes right on the sheet.** Changes go straight into the workbook and are selected, so you can look at them. A card lists every cell that changed (before → after, with the formula) and formula results elsewhere that changed, with **Undo** and **Advanced format** buttons. To approve each change first, turn on **Ask before changing the workbook**.
  - **Checked before it lands.** A change whose formulas would show errors (#SPILL!, #NAME?, #REF!…) isn't applied: the assistant is told exactly what's wrong and fixes it.
  - **One-click undo.** Each applied change is a single step with an **Undo** button. The button still works after you've made other edits, and keeps any cells you changed yourself since.
  - Use **Claude** with your own API key, or a **model on your computer** with Ollama or LM Studio, so the workbook never leaves it. Choose in **Settings → AI assistant**.
  - **Your API key stays secret.** It's encrypted with your computer's keychain and never shown again. Only Gridline's background process uses it, and only to talk to Anthropic. Keys pasted into the chat aren't sent, and keys found in a workbook are hidden from the AI.
  - Cell addresses in answers are links that jump to those cells.
  - **Uses real data, not made-up numbers.** The assistant knows Gridline's live-data functions, so "make a table with NFL quarterback stats" builds a live `NFL.LEADERS` table instead of typing in invented stats. It can also look up live data while it works.
  - **Local models get the full instructions.** Ollama's default context (often 4,096 tokens) silently cut off over half of what the assistant needs to know, so models didn't realize they were in a spreadsheet. Gridline now sets it to 16,384 tokens (**Settings → AI assistant → Context size**) and warns when a chat outgrows it. A tip suggests a larger model when you pick a very small one.
  - **Works with more local models.** Some models, like Qwen3-Coder in Ollama, write their actions as text instead of sending them properly. Gridline now recognizes these, so the changes show up as proposals instead of code in the chat. Reasoning blocks (`<think>…</think>`) are hidden.
  - Proposal previews show "Loading…" while live data downloads and fill in when it arrives. A formula that fills a whole table shows how much of the sheet it covers.

## 1.0.5

- **Your VBA functions work in cells.** A `Function` in a workbook's macros can be used in formulas, like in Excel: `=VAULTVALUE("Rent", "$")`, `=ESTTAX(B2)`.
  - Cell references passed to `Variant` or `Range` parameters arrive as Range objects, so code can read `.NumberFormat`, `.Text` and more.
  - Optional parameters, error values (`CVErr`), arrays that spill, `Application.Volatile` and `Application.Caller` all work.
  - Like in Excel, these functions can read the whole workbook but not change it. They run only while the workbook's macros are on; until then, their cells show the values saved in the file.
- `IsError(CVErr(…))` is now True, and a macro writing `CVErr(xlErrNA)` to a cell puts `#N/A` there instead of the number 2042.
- **Step macros (Beta) can now use variables.** Each step macro has a **Variables** list above its steps:
  - click **New variable** to make as many as you need, each with its own name (no two alike) and a starting value: text, a number, **the value in a cell**, or a formula's result;
  - click a variable to rename it (the steps using it follow) or change what it starts as; hover it to see the steps that use it;
  - steps pick a variable from a dropdown, which also has **New variable…**;
  - use any variable in a box as `{Name}`.
- A new **Variables** group of steps:
  - **Set a variable** to a value, the value in a cell, or the result of a formula with variables in it (`=ROUND({Price}*{Qty}, 2)`);
  - **Change a number** (add, subtract, multiply, divide, round, remainder, limit), **Change text** (add to it, replace, uppercase, trim, keep the first or last characters) and **Change a date** (add days, working days, weeks, months or years);
  - **Get a total or count** of cells, the last row with data or the next empty row;
  - **Find a value** and remember its row, its address or the value next to it;
  - **Pick a random number**.
- **More logic in step macros:**
  - **If** can check a `{Variable}`, not just a cell, and can take a second rule joined with *and* or *or*;
  - new loops: **Repeat while/until**, **For each cell** (`{Cell}` and `{Value}`), **For each item in a list** (`{Item}`) and **For each sheet** (`{Sheet}`);
  - **Stop repeating** leaves a loop early; **If a step fails** can skip errors and keep going; **Note** adds a comment; **Write to the output** prints to the output panel.
  - Comparisons work the way you'd expect: numbers as numbers (an answer of "1,200" counts as 1200), dates as dates, and text ignoring upper and lower case.
- **A better step builder:**
  - drag a step by its icon to move it, including into and out of If and loop blocks;
  - the **{ }** button next to a box inserts a variable, listing the ones from the blocks around the step, your own, and `{Today}`/`{Now}`;
  - the variables a macro uses are shown at the top; click one to point out the steps that use it;
  - mistakes show above the steps, like a variable that's used but never set ("Did you mean {Total}?");
  - boxes that don't apply are hidden (the replacement text only shows for *replace*).
  - Two new examples: **Add up paid amounts** and **A sheet per quarter**.
- **Tables (Ctrl+T), made for sorting.** Click in your data and press **Ctrl+T** (or **Insert → Table**). Every column name gets an arrow:
  - **one click** sorts smallest first (A to Z, oldest first), **two** largest first, **three** puts the rows back in their original order, which Excel can't do once you've sorted;
  - **Shift+click** another arrow to sort by it too; the arrows show 1, 2… in order;
  - sorting understands your data: numbers typed as text (`1,200`, `$5`) sort as numbers, `Item 2` comes before `Item 10`, capitals don't matter, and empty cells go last;
  - the sorted column is tinted, and a dot on its arrow says rows changed since sorting; click to sort again;
  - in a long table, the column names and arrows take the place of the column letters once you scroll past the header.
- **Tables are easier than in Excel:**
  - no dialog, and **nothing in your cells changes**: values, formulas, formats, colors, column widths and rows stay exactly as they were. The table adds arrows to its first row and a thin outline (colors and banded rows are optional, in **Table options**);
  - selecting whole columns, or an empty row or a title above your data, makes the table from just the data and its header row;
  - typing just below or right of a table makes it bigger; new rows get the formulas of calculated columns, and formulas that use a whole column (`=SUM(D2:D41)`) grow with it;
  - a total typed under a table stays outside it, so sorting never mixes it in;
  - each column has a **▾** menu (or **Alt+↓**) with sort choices, **Original order** and a filter; **Table options** sets the name, color and banded rows, or turns the table back into normal cells.
- **Excel's structured references work:** `=SUM(Sales[Amount])`, `[@Price]*[@Qty]`, `Sales[#All]`, `Sales[[#Headers],[Qty]]` and more. Renaming a table or a column updates the formulas that use it. Before, these showed #VALUE!.
- Tables are saved as real Excel tables in `.xlsx` and `.xlsm` files (sort order, filters, color and banding included), and tables in Excel files open as tables.
- **Ctrl+T** now makes a table, like in Excel; **Ctrl+N** still opens a new workbook tab.
- **Tabs for workbooks.** Open several files in one window, each in its own tab at the top, like a web browser:
  - **Ctrl+N** opens a new workbook in a new tab;
  - files you open, from Gridline or from File Explorer, open in tabs too;
  - **Ctrl+Tab** switches tabs, **Ctrl+W** closes one and **Ctrl+Shift+T** reopens the last file you closed;
  - drag tabs to reorder them, and right-click one for more options;
  - a dot shows which ones have unsaved changes;
  - copy and paste works between tabs, formatting included;
  - AutoSave covers every open file.
- **Menus close when you click elsewhere.** Right-click menus, ribbon dropdowns, color and font pickers and filter menus now close when you click the sheet, the title bar or anywhere else, when you scroll, when the window loses focus, and when it's resized. Clicking a dropdown's button again closes it.
- **Better keyboard use in menus.** → opens a submenu and ← goes back; Home and End jump to the first and last item; typing a letter jumps to the next item starting with it; disabled items are skipped. Escape closes one level at a time, and focus goes back to where you were.
- **A fresh look.** Gridline now looks less like a copy of Excel:
  - ribbon tools sit in rounded groups instead of labeled sections (the group names can be turned back on in **Settings → Ribbon**);
  - pill-shaped tabs, and a quieter File button;
  - the sheet, side panels and macro editor are rounded cards;
  - selected row and column headings get a soft highlight instead of a bar;
  - smoother hover effects, and menus fade in.
- The Advanced format card closes (keeping the formatting) when you click elsewhere.
- Fixed a light outline around right-click menus.
- **Gridline is part of a family of apps.** The new dotted button at the top left switches to [Pixelline](https://github.com/LucasHeh47/pixelline) (images and pixel art), [Frameline](https://github.com/LucasHeh47/frameline) (video editing) and [Baseline](https://github.com/LucasHeh47/baseline) (a home for your apps and files), or links to their download page. Baseline can start a blank workbook or any preset in Gridline and reopen your recent workbooks.

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
