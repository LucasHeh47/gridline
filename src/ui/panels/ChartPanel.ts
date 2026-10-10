// Edit the selected chart.
import * as L from 'lucide';
import type { IconNode } from 'lucide';
import { formatRange } from '../../core/address';
import { CHART_COLORS, chartData, guessSource, resolveRef, seriesFromSource } from '../../core/charts';
import type { ChartDef, ChartType } from '../../core/workbook';
import type { App } from '../app';
import { h, icon } from '../dom';
import { colorPicker, Popover, toast, tooltip } from '../widgets';
import type { PanelContent } from './SidePanel';

const TYPES: [ChartType, IconNode, string][] = [
  ['column', L.ChartColumn, 'Column'],
  ['bar', L.ChartBar, 'Bar'],
  ['line', L.ChartLine, 'Line'],
  ['area', L.ChartArea, 'Area'],
  ['pie', L.ChartPie, 'Pie'],
  ['doughnut', L.Donut, 'Doughnut'],
  ['scatter', L.ChartScatter, 'Scatter'],
];

export class ChartPanel implements PanelContent {
  title = 'Chart';
  readonly el: HTMLElement;

  constructor(private app: App) {
    this.el = h('div', { class: 'sp-body' });
    this.render();
  }

  refresh(): void {
    if (document.activeElement && this.el.contains(document.activeElement) && /INPUT|SELECT/.test((document.activeElement as HTMLElement).tagName)) return;
    this.render();
  }

  private render(): void {
    const app = this.app;
    const b = this.el;
    b.innerHTML = '';
    const def = app.charts.selectedDef();
    if (!def) {
      b.append(h('p', { class: 'hint' }, 'Select a chart to edit it, or insert one from the toolbar after selecting your data.'));
      return;
    }
    const up = (patch: Partial<ChartDef>, label?: string) => app.charts.update(def, patch, label);

    // Type
    const types = h('div', { class: 'segmented', style: { flexWrap: 'wrap' } });
    for (const [t, ic, label] of TYPES) {
      const btn = h('button', { class: `btn${def.type === t ? ' on' : ''}`, onclick: () => {
        let series = def.series;
        if ((t === 'pie' || t === 'doughnut') && series.length > 1) series = series.slice(0, 1);
        else if (def.source && (def.type === 'pie' || def.type === 'doughnut')) series = seriesFromSource(app.doc, def.source, t);
        up({ type: t, series }, 'Change chart type');
      } }, icon(ic, 16));
      tooltip(btn, label);
      types.append(btn);
    }
    b.append(h('div', { class: 'field' }, h('label', null, 'Chart type'), types));

    // Title
    const title = h('input', { class: 'input', value: def.title ?? '', placeholder: 'Chart title' }) as HTMLInputElement;
    title.addEventListener('change', () => up({ title: title.value.trim() || undefined }, 'Chart title'));
    b.append(h('div', { class: 'field' }, h('label', null, 'Title'), title));

    // Data
    const src = def.source;
    const rangeInput = h('input', { class: 'input mono', value: src?.range ?? '', placeholder: 'e.g. Sheet1!A1:C12' }) as HTMLInputElement;
    const useSel = h('button', { class: 'btn outline', onclick: () => {
      const r = app.dataRegion();
      const source = guessSource(app.doc, app.sheet, r);
      up({ source, series: seriesFromSource(app.doc, source, def.type) }, 'Chart data');
    } }, 'Use selection');
    tooltip(useSel, 'Chart the currently selected cells');
    rangeInput.addEventListener('change', () => {
      const res = resolveRef(app.doc, rangeInput.value, app.sheet);
      if (!res) {
        toast('Enter a range like A1:C12 or Sheet2!A1:B20.', { kind: 'error' });
        return;
      }
      const source = { ...guessSource(app.doc, res.sheet, res.range) };
      up({ source, series: seriesFromSource(app.doc, source, def.type) }, 'Chart data');
    });
    b.append(h('div', { class: 'field' }, h('label', null, 'Data range'), h('div', { class: 'row' }, rangeInput, useSel)));
    if (src) {
      const regen = (patch: Partial<typeof src>) => {
        const source = { ...src, ...patch };
        up({ source, series: seriesFromSource(app.doc, source, def.type) }, 'Chart data');
      };
      const byRows = h('div', { class: 'segmented' },
        h('button', { class: `btn${!src.byRows ? ' on' : ''}`, onclick: () => regen({ byRows: false }) }, 'Columns'),
        h('button', { class: `btn${src.byRows ? ' on' : ''}`, onclick: () => regen({ byRows: true }) }, 'Rows'),
      );
      const hr = h('input', { type: 'checkbox', checked: src.headerRow }) as HTMLInputElement;
      hr.addEventListener('change', () => regen({ headerRow: hr.checked }));
      const hc = h('input', { type: 'checkbox', checked: src.headerCol }) as HTMLInputElement;
      hc.addEventListener('change', () => regen({ headerCol: hc.checked }));
      b.append(
        h('div', { class: 'field' }, h('label', null, 'Each series is a'), byRows),
        h('label', { class: 'check' }, hr, 'First row contains series names'),
        h('label', { class: 'check' }, hc, 'First column contains labels'),
      );
    } else {
      b.append(h('p', { class: 'hint' }, `${def.series.length} series from the file: ${def.series.map((s) => s.values).join(', ')}`));
    }

    // Series colors
    if (def.type !== 'pie' && def.type !== 'doughnut') {
      const data = chartData(app.doc, def, app.sheet);
      const list = h('div', { class: 'list' });
      def.series.forEach((se, i) => {
        const color = se.color ?? CHART_COLORS[i % CHART_COLORS.length];
        const sw = h('span', { style: { width: '16px', height: '16px', borderRadius: '4px', background: color, display: 'inline-block', flex: 'none' } });
        const item = h('div', { class: 'list-item', title: 'Change color' }, sw, h('span', null, data.series[i]?.name ?? `Series ${i + 1}`), h('span', { class: 'sub' }, se.values.replace(/^.*!/, '')));
        item.addEventListener('click', () => {
          const pop = new Popover(colorPicker((c) => {
            pop.close();
            const series = def.series.map((x, k) => (k === i ? { ...x, color: c ?? undefined } : x));
            up({ series }, 'Series color');
          }, { noneLabel: 'Default', current: color }), item);
          pop.open(item);
        });
        list.append(item);
      });
      b.append(h('div', { class: 'field', style: { marginTop: '12px' } }, h('label', null, 'Series'), list));
    }

    // Layout
    const legend = h('select', { class: 'input' }) as HTMLSelectElement;
    for (const [v, label] of [
      ['bottom', 'Bottom'],
      ['top', 'Top'],
      ['right', 'Right'],
      ['none', 'Hidden'],
    ])
      legend.append(h('option', { value: v }, label));
    legend.value = def.legend;
    legend.addEventListener('change', () => up({ legend: legend.value as ChartDef['legend'] }, 'Chart legend'));
    b.append(h('div', { class: 'field' }, h('label', null, 'Legend'), legend));
    if (def.type !== 'pie' && def.type !== 'doughnut') {
      const xt = h('input', { class: 'input', value: def.xTitle ?? '', placeholder: 'Horizontal axis title' }) as HTMLInputElement;
      const yt = h('input', { class: 'input', value: def.yTitle ?? '', placeholder: 'Vertical axis title' }) as HTMLInputElement;
      xt.addEventListener('change', () => up({ xTitle: xt.value.trim() || undefined }, 'Axis title'));
      yt.addEventListener('change', () => up({ yTitle: yt.value.trim() || undefined }, 'Axis title'));
      b.append(h('div', { class: 'field' }, h('label', null, 'Axis titles'), xt, yt));
      if (def.type !== 'scatter') {
        const st = h('input', { type: 'checkbox', checked: !!def.stacked }) as HTMLInputElement;
        st.addEventListener('change', () => up({ stacked: st.checked }, 'Stack series'));
        b.append(h('label', { class: 'check' }, st, 'Stack series on top of each other'));
      }
    }
    const res = def.source ? resolveRef(app.doc, def.source.range, app.sheet) : null;
    b.append(
      h('div', { class: 'row', style: { marginTop: '18px', justifyContent: 'space-between' } },
        res ? h('button', { class: 'btn outline', onclick: () => {
          if (res.sheet.id !== app.sheet.id) app.setActiveSheet(res.sheet.id, false);
          app.select(res.range);
        } }, icon(L.Crosshair, 14), `Select ${formatRange(res.range)}`) : h('span'),
        h('button', { class: 'btn outline danger', onclick: () => app.charts.deleteSelected() }, icon(L.Trash2, 14), 'Delete chart'),
      ),
    );
  }
}
