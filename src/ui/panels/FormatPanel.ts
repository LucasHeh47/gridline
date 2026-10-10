// Format cells: number, alignment, font, border, fill — applied live.
import * as L from 'lucide';
import type { IconNode } from 'lucide';
import { applyBorders, type BorderPreset } from '../../core/ops/cells';
import { formatValue, isDateFormat } from '../../core/numfmt';
import type { BorderStyle, HAlign, StylePatch, VAlign } from '../../core/styles';
import type { App } from '../app';
import { h, icon } from '../dom';
import { FONTS } from '../toolbar';
import { clampRange } from '../commands';
import { colorPicker, tooltip } from '../widgets';
import type { PanelContent } from './SidePanel';

type Tab = 'number' | 'alignment' | 'font' | 'border' | 'fill';

const CATEGORIES = ['Automatic', 'Number', 'Currency', 'Accounting', 'Percent', 'Scientific', 'Fraction', 'Date', 'Time', 'Text', 'Custom'] as const;
type Category = (typeof CATEGORIES)[number];

const DATE_FORMATS = ['m/d/yyyy', 'mm/dd/yyyy', 'd/m/yyyy', 'yyyy-mm-dd', 'mmm d, yyyy', 'mmmm d, yyyy', 'd-mmm-yy', 'dddd, mmmm d, yyyy', 'mmm-yy', 'mmmm yyyy', 'm/d/yyyy h:mm', 'yyyy-mm-dd hh:mm:ss'];
const TIME_FORMATS = ['h:mm AM/PM', 'h:mm:ss AM/PM', 'hh:mm', 'hh:mm:ss', '[h]:mm:ss', 'mm:ss.0'];
const CUSTOM_EXAMPLES = ['0', '0.00', '#,##0', '#,##0.00', '0%', '0.0%', '#,##0,"K"', '#,##0,,"M"', '"$"#,##0.00;[Red]-"$"#,##0.00', '000-00-0000', '(000) 000-0000', '0 "units"', '@" (text)"'];

function categorize(fmt: string): Category {
  if (!fmt || fmt === 'General') return 'Automatic';
  if (fmt === '@') return 'Text';
  if (/^_\(/.test(fmt)) return 'Accounting';
  if (isDateFormat(fmt)) return /[dy]/i.test(fmt.replace(/\[[^\]]*\]/g, '')) || /m{3,}/.test(fmt) ? 'Date' : 'Time';
  if (/%/.test(fmt)) return 'Percent';
  if (/E[+-]/i.test(fmt)) return 'Scientific';
  if (/\?\/\?|#\s*\?/.test(fmt)) return 'Fraction';
  if (/[$€£¥]/.test(fmt)) return 'Currency';
  if (/^[#,0.]+$/.test(fmt) || /^[#,0.]+;/.test(fmt)) return 'Number';
  return 'Custom';
}

export class FormatPanel implements PanelContent {
  title = 'Format cells';
  readonly el: HTMLElement;
  private tab: Tab;
  private body: HTMLDivElement;
  private tabsEl: HTMLDivElement;
  private category: Category | null = null;
  private borderStyle: BorderStyle = 'thin';
  private borderColor: string | undefined;

  constructor(
    private app: App,
    initial?: string,
  ) {
    this.tab = (initial as Tab) ?? 'number';
    this.tabsEl = h('div', { class: 'sp-tabs' });
    this.body = h('div', { class: 'sp-body' });
    this.el = h('div', { style: { display: 'flex', flexDirection: 'column', flex: '1', minHeight: '0' } }, this.tabsEl, this.body);
    this.render();
  }

  setTab(t: string): void {
    this.tab = t as Tab;
    this.render();
  }

  refresh(): void {
    if (document.activeElement && this.body.contains(document.activeElement) && (document.activeElement as HTMLElement).tagName === 'INPUT') return;
    this.render();
  }

  private patch(label: string, p: StylePatch): void {
    this.app.applyStyle(label, p);
  }

  private render(): void {
    this.tabsEl.innerHTML = '';
    const tabs: [Tab, string][] = [
      ['number', 'Number'],
      ['alignment', 'Alignment'],
      ['font', 'Font'],
      ['border', 'Border'],
      ['fill', 'Fill'],
    ];
    for (const [t, label] of tabs) {
      this.tabsEl.append(h('button', { class: `sp-tab${this.tab === t ? ' on' : ''}`, onclick: () => this.setTab(t) }, label));
    }
    this.body.innerHTML = '';
    const sel = this.app.selectionLabel();
    this.body.append(h('div', { class: 'hint', style: { marginBottom: '10px' } }, `Applies to ${sel}. Changes take effect immediately (Ctrl+Z to undo).`));
    switch (this.tab) {
      case 'number':
        this.renderNumber();
        break;
      case 'alignment':
        this.renderAlignment();
        break;
      case 'font':
        this.renderFont();
        break;
      case 'border':
        this.renderBorder();
        break;
      case 'fill':
        this.renderFill();
        break;
    }
  }

  private renderNumber(): void {
    const app = this.app;
    const st = app.styleAtActive();
    const fmt = st.numFmt ?? 'General';
    const cat = this.category ?? categorize(fmt);
    const a = app.sel.active;
    const v = app.valueAt(a.row, a.col);
    const sample = typeof v === 'number' ? v : cat === 'Date' || cat === 'Time' ? 45730.6354 : -1234.5678;
    const preview = h('div', { class: 'preview-box' });
    const showPreview = (f: string) => {
      const out = formatValue(typeof v === 'string' && cat === 'Text' ? v : sample, f, 30);
      preview.textContent = out.text || ' ';
      preview.style.color = out.color ?? '';
    };
    const setFmt = (f: string) => {
      this.patch('Number format', { numFmt: f === 'General' ? null : f });
      showPreview(f);
    };

    const catList = h('div', { class: 'chips', style: { marginBottom: '12px' } });
    for (const c of CATEGORIES) {
      catList.append(
        h('button', {
          class: `chip${c === cat ? ' on' : ''}`,
          onclick: () => {
            this.category = c;
            const defaults: Record<Category, string> = {
              Automatic: 'General',
              Number: '#,##0.00',
              Currency: '"$"#,##0.00',
              Accounting: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
              Percent: '0.00%',
              Scientific: '0.00E+00',
              Fraction: '# ?/?',
              Date: 'm/d/yyyy',
              Time: 'h:mm AM/PM',
              Text: '@',
              Custom: fmt,
            };
            if (c !== 'Custom') setFmt(defaults[c]);
            this.render();
          },
        }, c),
      );
    }
    this.body.append(catList, h('div', { class: 'field' }, h('label', null, 'Preview'), preview));
    showPreview(fmt);

    const decimalsOf = (f: string) => {
      const m = /\.(0+)/.exec(f);
      return m ? m[1].length : 0;
    };
    if (cat === 'Number' || cat === 'Currency' || cat === 'Percent' || cat === 'Scientific' || cat === 'Accounting') {
      const dec = h('input', { class: 'input', type: 'number', min: '0', max: '15', value: String(decimalsOf(fmt)), style: { width: '80px' } }) as HTMLInputElement;
      const thousands = h('input', { type: 'checkbox', checked: fmt.includes(',') || cat === 'Currency' || cat === 'Accounting' }) as HTMLInputElement;
      const symbol = h('select', { class: 'input' }) as HTMLSelectElement;
      for (const s of ['$', '€', '£', '¥', '₹', 'CHF ', 'kr ']) symbol.append(h('option', { value: s }, s.trim()));
      const curSym = /"([^"]+)"/.exec(fmt)?.[1] ?? /\[\$([^-\]]+)/.exec(fmt)?.[1] ?? '$';
      symbol.value = curSym;
      const neg = h('select', { class: 'input' }) as HTMLSelectElement;
      for (const [val, label] of [
        ['minus', '-1,234.10'],
        ['red', '1,234.10 in red'],
        ['paren', '(1,234.10)'],
        ['redparen', '(1,234.10) in red'],
      ])
        neg.append(h('option', { value: val }, label));
      neg.value = /\[Red\]\(/.test(fmt) ? 'redparen' : /\[Red\]/.test(fmt) ? 'red' : /;\(/.test(fmt) || /;"?\$?"?\(/.test(fmt) ? 'paren' : 'minus';
      const build = () => {
        const d = Math.max(0, Math.min(15, parseInt(dec.value, 10) || 0));
        const decPart = d ? '.' + '0'.repeat(d) : '';
        if (cat === 'Percent') return `0${decPart}%`;
        if (cat === 'Scientific') return `0${decPart}E+00`;
        if (cat === 'Accounting') {
          const s = `"${symbol.value.trim()}"`;
          return `_(${s}* #,##0${decPart}_);_(${s}* (#,##0${decPart});_(${s}* "-"??_);_(@_)`;
        }
        const body = `${thousands.checked ? '#,##0' : '0'}${decPart}`;
        const pos = cat === 'Currency' ? `"${symbol.value.trim()}"${body}` : body;
        switch (neg.value) {
          case 'red':
            return `${pos};[Red]-${pos}`;
          case 'paren':
            return `${pos};(${pos})`;
          case 'redparen':
            return `${pos};[Red](${pos})`;
          default:
            return pos;
        }
      };
      const update = () => setFmt(build());
      dec.addEventListener('change', update);
      thousands.addEventListener('change', update);
      symbol.addEventListener('change', update);
      neg.addEventListener('change', update);
      this.body.append(h('div', { class: 'field' }, h('label', null, 'Decimal places'), dec));
      if (cat === 'Number') this.body.append(h('label', { class: 'check' }, thousands, 'Use thousands separator (1,000)'));
      if (cat === 'Currency' || cat === 'Accounting') this.body.append(h('div', { class: 'field' }, h('label', null, 'Symbol'), symbol));
      if (cat === 'Number' || cat === 'Currency') this.body.append(h('div', { class: 'field', style: { marginTop: '10px' } }, h('label', null, 'Negative numbers'), neg));
    }
    if (cat === 'Date' || cat === 'Time') {
      const list = h('div', { class: 'list' });
      for (const f of cat === 'Date' ? DATE_FORMATS : TIME_FORMATS) {
        list.append(h('div', { class: `list-item${f === fmt ? ' on' : ''}`, onclick: () => {
          setFmt(f);
          this.render();
        } }, formatValue(sample, f).text, h('span', { class: 'sub' }, f)));
      }
      this.body.append(h('div', { class: 'field' }, h('label', null, 'Style'), list));
    }
    if (cat === 'Fraction') {
      const list = h('div', { class: 'list' });
      for (const f of ['# ?/?', '# ??/??', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100']) {
        list.append(h('div', { class: `list-item${f === fmt ? ' on' : ''}`, onclick: () => {
          setFmt(f);
          this.render();
        } }, formatValue(Math.abs(sample), f).text, h('span', { class: 'sub' }, f)));
      }
      this.body.append(list);
    }
    if (cat === 'Text') this.body.append(h('p', { class: 'hint' }, 'Cells formatted as text keep exactly what you type — leading zeros, long ID numbers and things that look like dates stay as they are.'));
    if (cat === 'Automatic') this.body.append(h('p', { class: 'hint' }, 'Numbers are shown as typed, with as many decimals as fit in the cell.'));
    if (cat === 'Custom' || cat !== 'Automatic') {
      const code = h('input', { class: 'input mono', value: fmt, spellcheck: false }) as HTMLInputElement;
      code.addEventListener('input', () => showPreview(code.value || 'General'));
      code.addEventListener('change', () => setFmt(code.value || 'General'));
      code.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') setFmt(code.value || 'General');
      });
      this.body.append(h('div', { class: 'field', style: { marginTop: '12px' } }, h('label', null, 'Format code (advanced)'), code));
      if (cat === 'Custom') {
        const ex = h('div', { class: 'list' });
        for (const f of CUSTOM_EXAMPLES) ex.append(h('div', { class: 'list-item', onclick: () => {
          code.value = f;
          setFmt(f);
        } }, formatValue(sample, f).text, h('span', { class: 'sub' }, f)));
        this.body.append(h('div', { class: 'field' }, h('label', null, 'Examples'), ex));
      }
    }
  }

  private renderAlignment(): void {
    const st = this.app.styleAtActive();
    const seg = <T extends string>(items: [T, IconNode, string][], cur: T, onPick: (v: T) => void) => {
      const box = h('div', { class: 'segmented' });
      for (const [v, ic, label] of items) {
        const b = h('button', { class: `btn${cur === v ? ' on' : ''}`, onclick: () => {
          onPick(v);
          this.render();
        } }, icon(ic, 16));
        tooltip(b, label);
        box.append(b);
      }
      return box;
    };
    this.body.append(
      h('div', { class: 'field' }, h('label', null, 'Horizontal'), seg<HAlign>([
        ['general', L.TextCursor, 'Automatic (numbers right, text left)'],
        ['left', L.AlignLeft, 'Left'],
        ['center', L.AlignCenter, 'Center'],
        ['right', L.AlignRight, 'Right'],
        ['centerContinuous', L.MoveHorizontal, 'Center across selection (no merge)'],
      ], st.hAlign ?? 'general', (v) => this.patch('Alignment', { hAlign: v === 'general' ? null : v }))),
      h('div', { class: 'field' }, h('label', null, 'Vertical'), seg<VAlign>([
        ['top', L.AlignVerticalJustifyStart, 'Top'],
        ['middle', L.AlignVerticalJustifyCenter, 'Middle'],
        ['bottom', L.AlignVerticalJustifyEnd, 'Bottom'],
      ], st.vAlign ?? 'bottom', (v) => this.patch('Alignment', { vAlign: v === 'bottom' ? null : v }))),
    );
    const wrap = h('input', { type: 'checkbox', checked: !!st.wrap }) as HTMLInputElement;
    wrap.addEventListener('change', () => this.app.commands.run('fmt.wrap'));
    const indent = h('input', { class: 'input', type: 'number', min: '0', max: '15', value: String(st.indent ?? 0), style: { width: '80px' } }) as HTMLInputElement;
    indent.addEventListener('change', () => this.patch('Indent', { indent: Math.max(0, Math.min(15, +indent.value || 0)) || null }));
    this.body.append(h('label', { class: 'check' }, wrap, 'Wrap text onto multiple lines'), h('div', { class: 'field', style: { marginTop: '12px' } }, h('label', null, 'Indent'), indent));
    const mergeRow = h('div', { class: 'row wrap' });
    for (const [id, label] of [
      ['fmt.merge.center', 'Merge & center'],
      ['fmt.merge.merge', 'Merge'],
      ['fmt.merge.across', 'Merge across'],
      ['fmt.merge.unmerge', 'Unmerge'],
    ])
      mergeRow.append(h('button', { class: 'btn outline', onclick: () => this.app.commands.run(id) }, label));
    this.body.append(h('div', { class: 'field' }, h('label', null, 'Merge'), mergeRow));
  }

  private renderFont(): void {
    const st = this.app.styleAtActive();
    const family = h('select', { class: 'input' }) as HTMLSelectElement;
    for (const f of FONTS) family.append(h('option', { value: f, style: { fontFamily: f } }, f));
    family.value = st.fontName ?? 'Calibri';
    family.addEventListener('change', () => this.patch('Font', { fontName: family.value === 'Calibri' ? null : family.value }));
    const size = h('input', { class: 'input', type: 'number', min: '1', max: '409', value: String(st.fontSize ?? 11), style: { width: '80px' } }) as HTMLInputElement;
    size.addEventListener('change', () => this.patch('Font size', { fontSize: +size.value === 11 ? null : +size.value }));
    const toggles = h('div', { class: 'segmented' });
    for (const [key, ic, label] of [
      ['bold', L.Bold, 'Bold'],
      ['italic', L.Italic, 'Italic'],
      ['underline', L.Underline, 'Underline'],
      ['strike', L.Strikethrough, 'Strikethrough'],
    ] as const) {
      const b = h('button', { class: `btn${st[key] ? ' on' : ''}`, onclick: () => {
        this.patch(label, { [key]: !st[key] || null });
        this.render();
      } }, icon(ic, 16));
      tooltip(b, label);
      toggles.append(b);
    }
    this.body.append(
      h('div', { class: 'field' }, h('label', null, 'Font'), family),
      h('div', { class: 'field' }, h('label', null, 'Size (points)'), size),
      h('div', { class: 'field' }, h('label', null, 'Style'), toggles),
      h('div', { class: 'field' }, h('label', null, 'Text color'), colorPicker((c) => {
        this.patch('Text color', { color: c });
        this.render();
      }, { noneLabel: 'Automatic', current: st.color })),
    );
  }

  private renderBorder(): void {
    const app = this.app;
    const styleSel = h('select', { class: 'input' }) as HTMLSelectElement;
    for (const [v, label] of [
      ['thin', 'Thin'],
      ['medium', 'Medium'],
      ['thick', 'Thick'],
      ['dashed', 'Dashed'],
      ['dotted', 'Dotted'],
      ['double', 'Double'],
      ['hair', 'Hairline'],
    ])
      styleSel.append(h('option', { value: v }, label));
    styleSel.value = this.borderStyle;
    styleSel.addEventListener('change', () => (this.borderStyle = styleSel.value as BorderStyle));
    const grid = h('div', { class: 'preset-grid' });
    const presets: [BorderPreset, IconNode, string][] = [
      ['all', L.Grid3x3, 'All'],
      ['outer', L.Square, 'Outside'],
      ['inner', L.Plus, 'Inside'],
      ['top', L.PanelTop, 'Top'],
      ['bottom', L.PanelBottom, 'Bottom'],
      ['left', L.PanelLeft, 'Left'],
      ['right', L.PanelRight, 'Right'],
      ['horizontal', L.Rows3, 'Horizontal'],
      ['vertical', L.Columns3, 'Vertical'],
      ['none', L.SquareDashed, 'None'],
    ];
    for (const [p, ic, label] of presets) {
      grid.append(
        h('button', {
          class: 'preset',
          onclick: () => app.transact(`Borders: ${label}`, (tx) => {
            for (const r of app.sel.ranges) applyBorders(tx, app.doc.wb, app.sheet, clampRange(app, r), p, { style: this.borderStyle, color: this.borderColor });
          }),
        }, icon(ic, 18), label),
      );
    }
    this.body.append(
      h('div', { class: 'field' }, h('label', null, 'Line style'), styleSel),
      h('div', { class: 'field' }, h('label', null, 'Line color'), colorPicker((c) => {
        this.borderColor = c ?? undefined;
      }, { noneLabel: 'Automatic', current: this.borderColor })),
      h('div', { class: 'field' }, h('label', null, 'Apply'), grid),
    );
  }

  private renderFill(): void {
    const st = this.app.styleAtActive();
    this.body.append(
      h('div', { class: 'field' }, h('label', null, 'Background color'), colorPicker((c) => {
        this.patch('Fill color', { fill: c });
        this.render();
      }, { noneLabel: 'No fill', current: st.fill })),
      h('p', { class: 'hint' }, 'In dark mode, very light fills are shown as darker tints so text stays readable. The real colors are saved in the file. Toggle this in the command palette: "Dark-friendly cell colors".'),
    );
  }
}
