// Every setting in Gridline. Each one is wired to real behavior: `apply` runs when it changes
// (and at startup) and `refresh` lists the parts of the UI that redraw.
import * as L from 'lucide';
import { config } from '../../core/config';
import { TEAMS } from '../../core/data/nflTeams';
import { setChartPalette } from '../../core/charts';
import { platform } from '../platform';
import { uiOptions } from '../widgets';
import type { SectionDef, SettingDef, SettingOption } from './store';

export const SECTIONS: SectionDef[] = [
  { id: 'appearance', label: 'Appearance', icon: L.Palette, description: 'Theme, colors, fonts and what is shown on screen.' },
  { id: 'ribbon', label: 'Ribbon', icon: L.PanelTop, description: 'Size, tabs and labels of the ribbon.' },
  { id: 'grid', label: 'Grid & sheets', icon: L.Grid3x3, description: 'How cells look and behave, and defaults for new workbooks.' },
  { id: 'editing', label: 'Editing', icon: L.Pencil, description: 'Typing, Enter and Tab, autocomplete and undo.' },
  { id: 'formulas', label: 'Formulas & errors', icon: L.SquareFunction, description: 'Formula bar and error explanations.' },
  { id: 'smart', label: 'Advanced format', icon: L.Sparkles, description: 'How Advanced format styles your data.' },
  { id: 'files', label: 'Files & saving', icon: L.FileSpreadsheet, description: 'Opening, autosave and CSV options.' },
  { id: 'live', label: 'Live data', icon: L.CloudDownload, description: 'NFL stats, stock prices and exchange rates.' },
  { id: 'nfl', label: 'NFL', icon: L.Trophy, description: 'Defaults for the NFL data panel and templates.' },
  { id: 'charts', label: 'Charts', icon: L.ChartColumn, description: 'Colors and defaults for new charts.' },
  { id: 'status', label: 'Status bar', icon: L.PanelBottom, description: 'Totals and controls in the bottom bar.' },
  { id: 'keyboard', label: 'Keyboard & search', icon: L.Keyboard, description: 'Command search and shortcuts.' },
  { id: 'updates', label: 'Updates', icon: L.RefreshCw, description: 'Version and automatic updates.' },
];

const opts = (...pairs: [string, string][]): SettingOption[] => pairs.map(([value, label]) => ({ value, label }));
const minutes = (...m: number[]): SettingOption[] => [{ value: '0', label: 'Off' }, ...m.map((n) => ({ value: String(n), label: n === 1 ? 'Every minute' : `Every ${n} minutes` }))];

export const CURRENCIES = opts(['USD', 'US dollar'], ['EUR', 'Euro'], ['GBP', 'British pound'], ['CAD', 'Canadian dollar'], ['AUD', 'Australian dollar'], ['JPY', 'Japanese yen'], ['CHF', 'Swiss franc'], ['INR', 'Indian rupee'], ['MXN', 'Mexican peso'], ['BRL', 'Brazilian real']);

export const SETTINGS: SettingDef[] = [
  // ---------- Appearance ----------
  { id: 'appearance.theme', section: 'appearance', group: 'Theme', label: 'Theme', description: 'Dark is easiest on the eyes; "Match Windows" follows your system setting.', type: 'select', default: 'dark', options: opts(['dark', 'Dark'], ['light', 'Light'], ['system', 'Match Windows']), refresh: ['theme'], keywords: 'dark mode light mode' },
  { id: 'appearance.accent', section: 'appearance', group: 'Theme', label: 'Accent color', description: 'Used for the selection, buttons and highlights.', type: 'select', default: 'blue', options: opts(['blue', 'Blue'], ['teal', 'Teal'], ['green', 'Green'], ['purple', 'Purple'], ['pink', 'Pink'], ['orange', 'Orange'], ['red', 'Red'], ['gold', 'Gold']), refresh: ['theme'], keywords: 'color highlight selection' },
  { id: 'appearance.adaptColors', section: 'appearance', group: 'Theme', label: 'Dark-friendly cell colors', description: 'Softens bright cell fills in dark mode so imported spreadsheets stay readable. Files are never changed.', type: 'toggle', default: true, refresh: ['theme'] },
  { id: 'appearance.contrast', section: 'appearance', group: 'Theme', label: 'Contrast', description: 'High contrast makes borders and secondary text stronger.', type: 'select', default: 'normal', options: opts(['normal', 'Normal'], ['high', 'High']), refresh: ['theme'], keywords: 'accessibility' },
  { id: 'appearance.density', section: 'appearance', group: 'Layout', label: 'Density', description: 'Compact fits more on screen; comfortable has larger click targets.', type: 'select', default: 'comfortable', options: opts(['comfortable', 'Comfortable'], ['compact', 'Compact']), refresh: ['chrome', 'layout'], keywords: 'spacing size' },
  { id: 'appearance.uiFont', section: 'appearance', group: 'Layout', label: 'Interface font', type: 'select', default: 'system', options: opts(['system', 'Segoe UI (default)'], ['arial', 'Arial'], ['verdana', 'Verdana'], ['tahoma', 'Tahoma'], ['trebuchet', 'Trebuchet MS'], ['georgia', 'Georgia'], ['consolas', 'Consolas']), refresh: ['chrome'] },
  { id: 'appearance.uiZoom', section: 'appearance', group: 'Layout', label: 'Interface size', description: 'Scales the whole window: menus, ribbon and grid.', type: 'number', default: 100, min: 80, max: 150, step: 10, unit: '%', desktopOnly: true, apply: (v: number) => platform.setZoomFactor(v / 100), keywords: 'zoom scale bigger smaller' },
  { id: 'appearance.reduceMotion', section: 'appearance', group: 'Layout', label: 'Reduce motion', description: 'Turns off animations, including the moving border around copied cells.', type: 'toggle', default: false, refresh: ['chrome', 'grid'], keywords: 'animation accessibility' },
  { id: 'appearance.scrollbars', section: 'appearance', group: 'Layout', label: 'Scrollbars', type: 'select', default: 'normal', options: opts(['normal', 'Normal'], ['thin', 'Thin']), refresh: ['chrome'] },
  { id: 'appearance.showFormulaBar', section: 'appearance', group: 'Show on screen', label: 'Formula bar', type: 'toggle', default: true, refresh: ['chrome', 'layout'] },
  { id: 'appearance.showStatusBar', section: 'appearance', group: 'Show on screen', label: 'Status bar', type: 'toggle', default: true, refresh: ['chrome', 'layout'] },
  { id: 'appearance.showSheetTabs', section: 'appearance', group: 'Show on screen', label: 'Sheet tabs', type: 'toggle', default: true, refresh: ['chrome', 'layout'] },
  { id: 'appearance.showSearch', section: 'appearance', group: 'Show on screen', label: 'Search box in the title bar', description: 'Ctrl+K works either way.', type: 'toggle', default: true, refresh: ['titlebar'] },
  { id: 'appearance.showQuickAccess', section: 'appearance', group: 'Show on screen', label: 'Save, undo and redo in the title bar', type: 'toggle', default: true, refresh: ['titlebar'], keywords: 'quick access toolbar' },
  { id: 'appearance.showThemeButton', section: 'appearance', group: 'Show on screen', label: 'Light/dark button in the title bar', type: 'toggle', default: true, refresh: ['titlebar'] },
  { id: 'appearance.showShortcutsButton', section: 'appearance', group: 'Show on screen', label: 'Keyboard shortcuts button in the title bar', type: 'toggle', default: true, refresh: ['titlebar'] },
  { id: 'appearance.titlePath', section: 'appearance', group: 'Show on screen', label: 'Show the file location in the window title', type: 'toggle', default: false, refresh: ['titlebar'], keywords: 'path folder' },
  { id: 'appearance.tooltips', section: 'appearance', group: 'Tooltips & messages', label: 'Tooltips', description: 'Short descriptions when you hover over a button.', type: 'toggle', default: true, apply: (v: boolean) => (uiOptions.tooltips = v) },
  { id: 'appearance.tooltipDelay', section: 'appearance', group: 'Tooltips & messages', label: 'Tooltip delay', type: 'number', default: 450, min: 0, max: 2000, step: 50, unit: 'ms', apply: (v: number) => (uiOptions.tooltipDelay = v) },
  { id: 'appearance.tooltipKeys', section: 'appearance', group: 'Tooltips & messages', label: 'Show keyboard shortcuts in tooltips', type: 'toggle', default: true, apply: (v: boolean) => (uiOptions.tooltipShortcuts = v) },
  { id: 'appearance.toastSeconds', section: 'appearance', group: 'Tooltips & messages', label: 'How long messages stay on screen', description: 'Error messages stay 3 seconds longer.', type: 'number', default: 4.5, min: 2, max: 20, step: 0.5, unit: 'seconds', apply: (v: number) => (uiOptions.toastMs = v * 1000), keywords: 'toast notification' },

  // ---------- Ribbon ----------
  { id: 'ribbon.size', section: 'ribbon', group: 'Size', label: 'Ribbon size', description: 'You can also drag the bottom edge of the ribbon, or double-click a tab.', type: 'select', default: 'full', options: opts(['full', 'Full (large buttons with labels)'], ['compact', 'Compact (one row)'], ['tabs', 'Tabs only (opens when you click a tab)']), refresh: ['ribbon', 'layout'], keywords: 'collapse hide toolbar resize' },
  { id: 'ribbon.labels', section: 'ribbon', group: 'Size', label: 'Button labels', description: 'Show text next to icons in the full ribbon.', type: 'toggle', default: true, refresh: ['ribbon', 'layout'] },
  { id: 'ribbon.groupLabels', section: 'ribbon', group: 'Size', label: 'Group names', description: 'The small names under each group (Font, Alignment…).', type: 'toggle', default: true, refresh: ['ribbon', 'layout'] },
  { id: 'ribbon.largeIcons', section: 'ribbon', group: 'Size', label: 'Large icons', type: 'toggle', default: false, refresh: ['ribbon', 'layout'] },
  { id: 'ribbon.tabs', section: 'ribbon', group: 'Tabs', label: 'Tabs to show', description: 'Home is always shown.', type: 'multi', default: ['insert', 'formulas', 'data', 'view', 'mine', 'help'], options: opts(['insert', 'Insert'], ['formulas', 'Formulas'], ['data', 'Data'], ['view', 'View'], ['mine', 'Money & NFL'], ['help', 'Help']), refresh: ['ribbon'] },
  { id: 'ribbon.startTab', section: 'ribbon', group: 'Tabs', label: 'Tab shown when Gridline opens', type: 'select', default: 'home', options: opts(['home', 'Home'], ['last', 'The last tab I used']), refresh: [] },
  { id: 'ribbon.wheelTabs', section: 'ribbon', group: 'Tabs', label: 'Scroll over the tabs to switch between them', type: 'toggle', default: true },
  { id: 'ribbon.autoClose', section: 'ribbon', group: 'Tabs', label: 'Close the pop-up ribbon after a command', description: 'Applies when the ribbon is set to "Tabs only".', type: 'toggle', default: true },

  // ---------- Grid & sheets ----------
  { id: 'grid.lines', section: 'grid', group: 'Look', label: 'Gridline color', type: 'select', default: 'normal', options: opts(['subtle', 'Subtle'], ['normal', 'Normal'], ['strong', 'Strong']), refresh: ['theme'] },
  { id: 'grid.headings', section: 'grid', group: 'Look', label: 'Row and column headings', description: 'The A, B, C… and 1, 2, 3… labels.', type: 'toggle', default: true, refresh: ['layout'], keywords: 'headers' },
  { id: 'grid.highlightHeadings', section: 'grid', group: 'Look', label: 'Highlight the headings of selected cells', type: 'toggle', default: true, refresh: ['grid'] },
  { id: 'grid.crosshair', section: 'grid', group: 'Look', label: 'Highlight the active row and column', description: 'A faint band that makes it easy to read across wide tables.', type: 'toggle', default: false, refresh: ['grid'], keywords: 'crosshair reading guide' },
  { id: 'grid.errorMarkers', section: 'grid', group: 'Look', label: 'Red corner on cells with errors', type: 'toggle', default: true, refresh: ['grid'] },
  { id: 'grid.noteMarkers', section: 'grid', group: 'Look', label: 'Orange corner on cells with notes', type: 'toggle', default: true, refresh: ['grid'] },
  { id: 'grid.spillBorder', section: 'grid', group: 'Look', label: 'Outline results that spill into several cells', type: 'toggle', default: true, refresh: ['grid'], keywords: 'dynamic array' },
  { id: 'grid.frozenLine', section: 'grid', group: 'Look', label: 'Line below frozen rows and columns', type: 'toggle', default: true, refresh: ['grid'] },
  { id: 'grid.negativeRed', section: 'grid', group: 'Numbers', label: 'Show negative numbers in red', description: 'For cells without their own text color. Only changes how they look.', type: 'toggle', default: false, refresh: ['grid'] },
  { id: 'grid.showZeros', section: 'grid', group: 'Numbers', label: 'Show zero values', description: 'Turn off to leave cells that equal 0 blank.', type: 'toggle', default: true, refresh: ['grid'] },
  { id: 'grid.loadingText', section: 'grid', group: 'Numbers', label: 'Cells waiting for live data show', type: 'select', default: 'loading', options: opts(['loading', '"Loading…"'], ['busy', '#BUSY!']), refresh: ['grid'] },
  { id: 'grid.fillHandle', section: 'grid', group: 'Mouse & scrolling', label: 'Fill handle', description: 'The small square at the corner of the selection that you drag to fill.', type: 'toggle', default: true, refresh: ['grid'], keywords: 'drag autofill' },
  { id: 'grid.doubleClickEdit', section: 'grid', group: 'Mouse & scrolling', label: 'Double-click a cell to edit it', type: 'toggle', default: true },
  { id: 'grid.cellTips', section: 'grid', group: 'Mouse & scrolling', label: 'Hover tips for notes, links and errors', type: 'toggle', default: true },
  { id: 'grid.scrollSpeed', section: 'grid', group: 'Mouse & scrolling', label: 'Scroll speed', type: 'number', default: 1, min: 0.5, max: 3, step: 0.25, unit: '×' },
  { id: 'grid.ctrlWheelZoom', section: 'grid', group: 'Mouse & scrolling', label: 'Ctrl + mouse wheel zooms', type: 'toggle', default: true },
  { id: 'grid.zoomStep', section: 'grid', group: 'Mouse & scrolling', label: 'Zoom step', type: 'select', default: '10', options: opts(['5', '5%'], ['10', '10%'], ['25', '25%']) },
  { id: 'grid.defaultZoom', section: 'grid', group: 'New workbooks', label: 'Zoom', type: 'number', default: 100, min: 50, max: 200, step: 10, unit: '%' },
  { id: 'grid.colWidth', section: 'grid', group: 'New workbooks', label: 'Column width', type: 'number', default: 88, min: 40, max: 300, step: 4, unit: 'px' },
  { id: 'grid.rowHeight', section: 'grid', group: 'New workbooks', label: 'Row height', type: 'number', default: 22, min: 16, max: 60, step: 1, unit: 'px' },
  { id: 'grid.sheetCount', section: 'grid', group: 'New workbooks', label: 'Number of sheets', type: 'number', default: 1, min: 1, max: 10, step: 1 },
  { id: 'grid.sheetPrefix', section: 'grid', group: 'New workbooks', label: 'Name for new sheets', description: 'New sheets are named this plus a number.', type: 'text', default: 'Sheet' },
  { id: 'grid.gridlines', section: 'grid', group: 'New workbooks', label: 'Show gridlines on new sheets', type: 'toggle', default: true },

  // ---------- Editing ----------
  { id: 'editing.enterDirection', section: 'editing', group: 'Enter & Tab', label: 'After pressing Enter, move', type: 'select', default: 'down', options: opts(['down', 'Down'], ['right', 'Right'], ['up', 'Up'], ['left', 'Left'], ['none', "Don't move"]) },
  { id: 'editing.tabReturn', section: 'editing', group: 'Enter & Tab', label: 'Enter after Tab returns to the starting column', description: 'Type across a row with Tab, press Enter, and land at the start of the next row.', type: 'toggle', default: true },
  { id: 'editing.recognizeDates', section: 'editing', group: 'Typing', label: 'Turn typed dates into dates', description: 'Off: "3/4" stays as text.', type: 'toggle', default: true, apply: (v: boolean) => (config.recognizeDates = v) },
  { id: 'editing.dateOrder', section: 'editing', group: 'Typing', label: 'Read typed dates as', type: 'select', default: 'MDY', options: opts(['MDY', 'Month/Day/Year (3/14/2026)'], ['DMY', 'Day/Month/Year (14/3/2026)']), apply: (v: 'MDY' | 'DMY') => (config.dateOrder = v), keywords: 'region locale' },
  { id: 'editing.leadingZeros', section: 'editing', group: 'Typing', label: 'Keep leading zeros', description: '00123 stays 00123 (ZIP codes, IDs) instead of becoming 123.', type: 'toggle', default: true, apply: (v: boolean) => (config.keepLeadingZeros = v) },
  { id: 'editing.longDigits', section: 'editing', group: 'Typing', label: 'Keep long numbers exact', description: 'More than 15 digits (card or account numbers) stay as text instead of being rounded.', type: 'toggle', default: true, apply: (v: boolean) => (config.longDigitsAsText = v) },
  { id: 'editing.autoLink', section: 'editing', group: 'Typing', label: 'Turn web addresses into links', type: 'toggle', default: true, apply: (v: boolean) => (config.autoLinkUrls = v) },
  { id: 'editing.autoDecimal', section: 'editing', group: 'Typing', label: 'Automatically insert a decimal point', description: 'For fast entry of money: with 2 places, typing 1999 gives 19.99.', type: 'number', default: 0, min: 0, max: 4, step: 1, unit: 'places (0 = off)', keywords: 'fixed decimal' },
  { id: 'editing.autoFix', section: 'editing', group: 'Typing', label: 'Fix missing parentheses and quotes', description: 'When a formula is missing a closing ) or ", add it instead of showing an error.', type: 'toggle', default: true },
  { id: 'editing.autocomplete', section: 'editing', group: 'Formula help', label: 'Suggest function names while typing', type: 'toggle', default: true, keywords: 'autocomplete' },
  { id: 'editing.argHints', section: 'editing', group: 'Formula help', label: 'Show what each function argument means', type: 'toggle', default: true },
  { id: 'editing.argValues', section: 'editing', group: 'Formula help', label: 'Suggest values inside quotes', description: 'Stat names, teams, players, currencies…', type: 'toggle', default: true },
  { id: 'editing.refColors', section: 'editing', group: 'Formula help', label: 'Color the cells a formula uses', type: 'toggle', default: true },
  { id: 'editing.fillCounts', section: 'editing', group: 'Fill & undo', label: 'Dragging a single number counts up', description: 'Off (like Excel): dragging 5 copies 5. On: 5, 6, 7… Hold Ctrl while dragging to do the opposite.', type: 'toggle', default: false, apply: (v: boolean) => (config.fillCountsSingleNumber = v) },
  { id: 'editing.undoLimit', section: 'editing', group: 'Fill & undo', label: 'Undo steps remembered', type: 'number', default: 200, min: 50, max: 1000, step: 50, apply: (v: number) => (config.undoLimit = v) },

  // ---------- Formulas ----------
  { id: 'formulas.explain', section: 'formulas', group: 'Errors', label: 'Explain errors in plain English', description: 'Shows a card explaining what went wrong when you select a cell with an error.', type: 'toggle', default: true },
  { id: 'formulas.quickFixes', section: 'formulas', group: 'Errors', label: 'Offer one-click fixes', type: 'toggle', default: true },
  { id: 'formulas.barFont', section: 'formulas', group: 'Formula bar', label: 'Formula bar font', type: 'select', default: 'mono', options: opts(['mono', 'Monospace (easier to count parentheses)'], ['ui', 'Same as the interface']), refresh: ['chrome'] },
  { id: 'formulas.barSize', section: 'formulas', group: 'Formula bar', label: 'Formula bar text size', type: 'number', default: 13, min: 11, max: 20, step: 1, unit: 'px', refresh: ['chrome', 'layout'] },
  { id: 'formulas.volatileMinutes', section: 'formulas', group: 'Calculation', label: 'Update NOW(), TODAY() and RAND()', description: 'Recalculate time-based formulas on a timer, not just when something changes.', type: 'select', default: '0', options: minutes(1, 5, 15, 60), refresh: ['timers'], keywords: 'volatile clock' },

  // ---------- Advanced format ----------
  { id: 'smart.style', section: 'smart', group: 'Style', label: 'First style to try', description: '"Try another style" cycles through the others.', type: 'select', default: 'banded', options: opts(['banded', 'Banded rows'], ['clean', 'Clean lines'], ['bold', 'Bold header'], ['grid', 'Boxed grid']) },
  { id: 'smart.color', section: 'smart', group: 'Style', label: 'Header color', type: 'select', default: 'blue', options: opts(['accent', 'Match accent color'], ['blue', 'Blue'], ['green', 'Green'], ['gray', 'Gray'], ['orange', 'Orange'], ['purple', 'Purple'], ['red', 'Red']) },
  { id: 'smart.header', section: 'smart', group: 'Style', label: 'Header row', type: 'select', default: 'auto', options: opts(['auto', 'Detect automatically'], ['always', 'The first row is always a header'], ['never', 'No header row']) },
  { id: 'smart.boldTotals', section: 'smart', group: 'Style', label: 'Style "Total" rows', description: 'Bold with a line above.', type: 'toggle', default: true },
  { id: 'smart.convertText', section: 'smart', group: 'Data', label: 'Convert numbers stored as text', description: 'Text like "$1,200" or "45%" becomes a real number so it can be added up.', type: 'toggle', default: true },
  { id: 'smart.currency', section: 'smart', group: 'Data', label: 'Currency symbol', type: 'select', default: '$', options: opts(['$', '$ Dollar'], ['€', '€ Euro'], ['£', '£ Pound'], ['¥', '¥ Yen'], ['₹', '₹ Rupee'], ['', 'None']) },
  { id: 'smart.dateFormat', section: 'smart', group: 'Data', label: 'Date format', type: 'select', default: 'mmm d, yyyy', options: opts(['mmm d, yyyy', 'Mar 14, 2026'], ['m/d/yyyy', '3/14/2026'], ['d/m/yyyy', '14/3/2026'], ['yyyy-mm-dd', '2026-03-14'], ['dddd, mmmm d, yyyy', 'Saturday, March 14, 2026']) },
  { id: 'smart.negatives', section: 'smart', group: 'Data', label: 'Negative money', type: 'select', default: 'minus', options: opts(['minus', '-$1,234.10'], ['red', '-$1,234.10 in red'], ['parens', '($1,234.10)'], ['redParens', '($1,234.10) in red']) },
  { id: 'smart.thousands', section: 'smart', group: 'Data', label: 'Thousands separator', description: '1,234,567 instead of 1234567 (years and IDs are never separated).', type: 'toggle', default: true },
  { id: 'smart.decimals', section: 'smart', group: 'Data', label: 'Decimal places for numbers', type: 'select', default: 'auto', options: opts(['auto', 'Automatic (match the data)'], ['0', '0'], ['1', '1'], ['2', '2'], ['3', '3']) },
  { id: 'smart.autofit', section: 'smart', group: 'Layout', label: 'Resize columns to fit', type: 'toggle', default: true },
  { id: 'smart.freeze', section: 'smart', group: 'Layout', label: 'Freeze the header row', description: 'Only when the sheet has nothing frozen yet.', type: 'toggle', default: true },
  { id: 'smart.filter', section: 'smart', group: 'Layout', label: 'Add filter buttons to the header', type: 'toggle', default: false },

  // ---------- Files ----------
  { id: 'files.startup', section: 'files', group: 'Opening', label: 'When Gridline starts', type: 'select', default: 'blank', options: opts(['blank', 'Open a blank workbook'], ['templates', 'Show the template gallery'], ['last', 'Reopen the last file']) },
  { id: 'files.openIn', section: 'files', group: 'Opening', label: 'Open files in', type: 'select', default: 'auto', options: opts(['auto', 'This window if it is empty, otherwise a new one'], ['new', 'Always a new window'], ['same', 'Always this window']) },
  { id: 'files.recentCount', section: 'files', group: 'Opening', label: 'Recent files to show', type: 'number', default: 10, min: 3, max: 30, step: 1 },
  { id: 'files.autosave', section: 'files', group: 'Saving', label: 'AutoSave', description: 'Saves files that already have a location. New files still need a first save.', type: 'select', default: '0', options: minutes(1, 2, 5, 10), refresh: ['timers'] },
  { id: 'files.csvWarning', section: 'files', group: 'Saving', label: 'Warn before saving formatting into a CSV', description: 'CSV files keep only values; formatting, formulas and extra sheets are lost.', type: 'toggle', default: true },
  { id: 'files.csvDelimiter', section: 'files', group: 'CSV files', label: 'Separator when saving CSV', type: 'select', default: ',', options: opts([',', 'Comma ,'], [';', 'Semicolon ;'], ['|', 'Pipe |']), apply: (v: ',' | ';' | '|') => (config.csv.delimiter = v) },
  { id: 'files.csvBom', section: 'files', group: 'CSV files', label: 'Add a byte order mark (BOM)', description: 'Helps Excel read accented characters. Turn off for tools that dislike it.', type: 'toggle', default: true, apply: (v: boolean) => (config.csv.bom = v) },
  { id: 'files.csvLineEnding', section: 'files', group: 'CSV files', label: 'Line endings', type: 'select', default: 'crlf', options: opts(['crlf', 'Windows (CRLF)'], ['lf', 'Mac / Linux (LF)']), apply: (v: string) => (config.csv.lineEnding = v === 'lf' ? '\n' : '\r\n') },
  { id: 'files.csvDates', section: 'files', group: 'CSV files', label: 'Write dates', type: 'select', default: 'display', options: opts(['display', 'As shown in the cell'], ['iso', 'ISO format (2026-03-14)']), apply: (v: 'display' | 'iso') => (config.csv.dates = v) },
  { id: 'files.csvDetect', section: 'files', group: 'CSV files', label: 'Recognize numbers and dates when opening CSV', description: 'Off: every value is kept exactly as text.', type: 'toggle', default: true, apply: (v: boolean) => (config.csv.detectTypes = v) },

  // ---------- Live data ----------
  { id: 'live.enabled', section: 'live', group: 'Connection', label: 'Live data', description: 'NFL, STOCK, CRYPTO and FX functions download data from the internet.', type: 'toggle', default: true, refresh: ['live'] },
  { id: 'live.offline', section: 'live', group: 'Connection', label: 'Offline mode', description: 'Only use data that was already downloaded.', type: 'toggle', default: false, desktopOnly: true, refresh: ['live'] },
  { id: 'live.loadingPill', section: 'live', group: 'Connection', label: 'Show "Loading live data" in the status bar', type: 'toggle', default: true, refresh: ['status'] },
  { id: 'live.quoteRefresh', section: 'live', group: 'Prices', label: 'Refresh stock, crypto and currency prices', type: 'select', default: '0', options: minutes(1, 5, 15, 30, 60), refresh: ['timers'] },
  { id: 'live.quoteMaxAge', section: 'live', group: 'Prices', label: 'Reuse a downloaded price for', type: 'number', default: 120, min: 15, max: 3600, step: 15, unit: 'seconds', apply: (v: number) => (config.quoteMaxAgeSec = v) },
  { id: 'live.currency', section: 'live', group: 'Prices', label: 'Default currency for CRYPTO', type: 'select', default: 'USD', options: CURRENCIES, apply: (v: string) => (config.defaultCurrency = v) },
  { id: 'live.nflHours', section: 'live', group: 'NFL data', label: 'Check for new current-season NFL stats every', description: 'Past seasons rarely change and are kept for 30 days.', type: 'number', default: 3, min: 1, max: 48, step: 1, unit: 'hours', apply: (v: number) => (config.nflCurrentSeasonHours = v) },
  { id: 'live.refreshNow', section: 'live', group: 'Downloaded data', label: 'Refresh all live data now', type: 'action', action: { label: 'Refresh', run: (app) => app.commands.run('data.refreshLive') } },
  { id: 'live.cache', section: 'live', group: 'Downloaded data', label: 'Downloaded data on this computer', type: 'info', desktopOnly: true, info: async () => formatBytes(await platform.dataCacheSize()) },
  { id: 'live.clearCache', section: 'live', group: 'Downloaded data', label: 'Delete downloaded data', description: "It's downloaded again when a formula needs it.", type: 'action', desktopOnly: true, action: { label: 'Delete', run: (app) => app.commands.run('data.clearCache') } },

  // ---------- NFL ----------
  { id: 'nfl.team', section: 'nfl', group: 'Defaults', label: 'Favorite team', description: 'Used by the NFL data panel and the NFL season tracker template.', type: 'select', default: 'Chiefs', options: () => [...TEAMS].sort((a, b) => `${a.city} ${a.nickname}`.localeCompare(`${b.city} ${b.nickname}`)).map((t) => ({ value: t.nickname, label: `${t.city} ${t.nickname}` })) },
  { id: 'nfl.fantasy', section: 'nfl', group: 'Defaults', label: 'Fantasy scoring', type: 'select', default: 'ppr', options: opts(['ppr', 'PPR'], ['half', 'Half PPR'], ['standard', 'Standard']) },
  { id: 'nfl.seasonType', section: 'nfl', group: 'Defaults', label: 'Games to include', type: 'select', default: 'REG', options: opts(['REG', 'Regular season'], ['POST', 'Playoffs'], ['ALL', 'Regular season + playoffs']) },
  { id: 'nfl.layout', section: 'nfl', group: 'Defaults', label: 'NFL data panel inserts', type: 'select', default: 'cells', options: opts(['cells', 'Cells you can drag'], ['table', 'One table formula']) },
  { id: 'nfl.seasons', section: 'nfl', group: 'Defaults', label: 'Seasons to show', type: 'number', default: 5, min: 1, max: 25, step: 1 },

  // ---------- Charts ----------
  { id: 'charts.palette', section: 'charts', group: 'Colors', label: 'Color palette', type: 'select', default: 'gridline', options: opts(['gridline', 'Gridline'], ['office', 'Office'], ['colorblind', 'Colorblind-friendly'], ['pastel', 'Pastel'], ['vivid', 'Vivid'], ['mono', 'Shades of one color']), apply: (v: string) => setChartPalette(v), refresh: ['charts'] },
  { id: 'charts.gridlines', section: 'charts', group: 'Look', label: 'Gridlines in charts', type: 'toggle', default: true, refresh: ['charts'] },
  { id: 'charts.piePercent', section: 'charts', group: 'Look', label: 'Show percentages on pie slices', type: 'toggle', default: true, refresh: ['charts'] },
  { id: 'charts.type', section: 'charts', group: 'New charts', label: 'Quick chart type (F11)', type: 'select', default: 'column', options: opts(['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie'], ['doughnut', 'Doughnut'], ['scatter', 'Scatter']) },
  { id: 'charts.legend', section: 'charts', group: 'New charts', label: 'Legend position', type: 'select', default: 'bottom', options: opts(['bottom', 'Bottom'], ['right', 'Right'], ['top', 'Top'], ['none', 'Hidden']) },
  { id: 'charts.size', section: 'charts', group: 'New charts', label: 'Size', type: 'select', default: 'medium', options: opts(['small', 'Small'], ['medium', 'Medium'], ['large', 'Large']) },

  // ---------- Status bar ----------
  { id: 'status.stats', section: 'status', group: 'Totals', label: 'Totals for selected cells', type: 'multi', default: ['sum', 'average', 'count', 'min', 'max'], options: opts(['sum', 'Sum'], ['average', 'Average'], ['count', 'Count'], ['numbers', 'Numerical count'], ['min', 'Min'], ['max', 'Max']), refresh: ['status'] },
  { id: 'status.copy', section: 'status', group: 'Totals', label: 'Click a total to copy it', type: 'toggle', default: true, refresh: ['status'] },
  { id: 'status.zoom', section: 'status', group: 'Controls', label: 'Zoom slider', type: 'toggle', default: true, refresh: ['status'] },
  { id: 'status.mode', section: 'status', group: 'Controls', label: 'Ready / Edit indicator', type: 'toggle', default: true, refresh: ['status'] },

  // ---------- Keyboard & search ----------
  { id: 'keyboard.recent', section: 'keyboard', group: 'Command search (Ctrl+K)', label: 'Show recently used commands first', type: 'toggle', default: true },
  { id: 'keyboard.remember', section: 'keyboard', group: 'Command search (Ctrl+K)', label: 'Remember the commands I use', type: 'toggle', default: true },
  { id: 'keyboard.cells', section: 'keyboard', group: 'Command search (Ctrl+K)', label: 'Search cell contents', type: 'toggle', default: true },
  { id: 'keyboard.functions', section: 'keyboard', group: 'Command search (Ctrl+K)', label: 'Search functions', type: 'toggle', default: true },
  { id: 'keyboard.clearRecent', section: 'keyboard', group: 'Command search (Ctrl+K)', label: 'Forget recently used commands', type: 'action', action: { label: 'Clear', run: (app) => {
    app.prefs.recentCommands = [];
    app.savePrefs();
  } } },
  { id: 'keyboard.list', section: 'keyboard', group: 'Shortcuts', label: 'All keyboard shortcuts', type: 'action', action: { label: 'Show', run: (app) => app.commands.run('help.shortcuts') } },

  // ---------- Updates ----------
  { id: 'updates.version', section: 'updates', group: 'Version', label: 'Installed version', type: 'info', info: async () => `Gridline ${await platform.appVersion()}` },
  { id: 'updates.check', section: 'updates', group: 'Version', label: 'Check for a newer version', type: 'action', desktopOnly: true, action: { label: 'Check now', run: (app) => app.updates.check(false) } },
  { id: 'updates.whatsNew', section: 'updates', group: 'Version', label: "What's new in this version", type: 'action', action: { label: 'Show', run: (app) => app.updates.showWhatsNew() } },
  { id: 'updates.onStartup', section: 'updates', group: 'Automatic updates', label: 'Check for updates when Gridline starts', type: 'toggle', default: true, desktopOnly: true },
  { id: 'updates.autoDownload', section: 'updates', group: 'Automatic updates', label: 'Download updates in the background', description: "You're still asked before anything is installed.", type: 'toggle', default: false, desktopOnly: true },
  { id: 'updates.channel', section: 'updates', group: 'Automatic updates', label: 'Update channel', description: 'Beta gets new features earlier.', type: 'select', default: 'stable', options: opts(['stable', 'Stable'], ['beta', 'Beta']), desktopOnly: true },
  { id: 'updates.showNotes', section: 'updates', group: 'Automatic updates', label: "Show what's new after updating", type: 'toggle', default: true },
];

function formatBytes(n: number): string {
  if (n < 1024) return `${n} bytes`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}
