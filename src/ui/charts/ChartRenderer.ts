// Renders charts as SVG.
import type { ResolvedChart } from '../../core/charts';
import { CHART_COLORS } from '../../core/charts';
import { formatValue } from '../../core/numfmt';
import type { ChartDef } from '../../core/workbook';

export interface ChartTheme {
  dark: boolean;
  bg: string;
  text: string;
  muted: string;
  grid: string;
  axis: string;
}

export interface HitInfo {
  series: string;
  category: string;
  value: string;
  color: string;
}

const NS = 'http://www.w3.org/2000/svg';

function el<K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number> = {}, text?: string): SVGElementTagNameMap[K] {
  const e = document.createElementNS(NS, tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
  if (text !== undefined) e.textContent = text;
  return e;
}

/** "Nice" axis bounds and step (1, 2, 2.5, 5 × 10^n). */
export function niceScale(min: number, max: number, ticks = 5): { min: number; max: number; step: number } {
  if (min === max) {
    if (min === 0) return { min: 0, max: 1, step: 0.2 };
    const pad = Math.abs(min) * 0.1;
    min -= pad;
    max += pad;
  }
  const range = max - min;
  const rough = range / ticks;
  const mag = Math.pow(10, Math.floor(Math.log10(rough)));
  const norm = rough / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  return { min: Math.floor(min / step + 1e-9) * step, max: Math.ceil(max / step - 1e-9) * step, step };
}

function compact(v: number, fmt?: string): string {
  if (fmt && fmt !== 'General') {
    const abs = Math.abs(v);
    if (abs >= 1e6 && /^[^%]*$/.test(fmt) && !/[dmyhs]/i.test(fmt)) {
      return formatValue(v / 1e6, fmt.replace(/\.0+/, '.0')).text + 'M';
    }
    return formatValue(v, fmt).text;
  }
  const abs = Math.abs(v);
  if (abs >= 1e9) return (v / 1e9).toFixed(abs >= 1e10 ? 0 : 1).replace(/\.0$/, '') + 'B';
  if (abs >= 1e6) return (v / 1e6).toFixed(abs >= 1e7 ? 0 : 1).replace(/\.0$/, '') + 'M';
  if (abs >= 1e4) return (v / 1e3).toFixed(abs >= 1e5 ? 0 : 1).replace(/\.0$/, '') + 'K';
  return String(Number(v.toPrecision(6)));
}

function measure(text: string, size: number): number {
  return text.length * size * 0.56;
}

export function renderChart(svg: SVGSVGElement, chart: ChartDef, data: ResolvedChart, width: number, height: number, t: ChartTheme): void {
  while (svg.firstChild) svg.removeChild(svg.firstChild);
  svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
  svg.setAttribute('font-family', 'Segoe UI, system-ui, sans-serif');
  svg.append(el('rect', { x: 0, y: 0, width, height, fill: t.bg }));
  const pie = chart.type === 'pie' || chart.type === 'doughnut';
  let top = 12;
  let bottom = height - 10;
  let left = 12;
  let right = width - 14;

  if (chart.title) {
    svg.append(el('text', { x: width / 2, y: 26, 'text-anchor': 'middle', 'font-size': 15, 'font-weight': 600, fill: t.text }, chart.title));
    top = 40;
  }

  // Legend entries
  // Pie slices: skip empty rows (no label and no value) so spare rows in a list don't show up as "Item 7"
  const legendItems = pie
    ? data.categories
        .map((c, i) => ({ name: c || `Item ${i + 1}`, color: CHART_COLORS[i % CHART_COLORS.length], empty: !c && !data.series[0]?.values[i] }))
        .filter((x) => !x.empty)
    : data.series.map((s) => ({ name: s.name, color: s.color }));
  const showLegend = chart.legend !== 'none' && (legendItems.length > 1 || pie);
  if (showLegend) {
    const fs = 12;
    if (chart.legend === 'right') {
      const lw = Math.min(width * 0.35, Math.max(...legendItems.map((i) => measure(i.name, fs))) + 26);
      let y = top + 6;
      const g = el('g');
      for (const it of legendItems.slice(0, Math.floor((bottom - top) / 20))) {
        g.append(el('rect', { x: right - lw + 4, y: y - 9, width: 10, height: 10, rx: 2, fill: it.color }));
        g.append(el('text', { x: right - lw + 20, y, 'font-size': fs, fill: t.muted }, truncate(it.name, lw - 24, fs)));
        y += 20;
      }
      svg.append(g);
      right -= lw + 4;
    } else {
      // Lay items out in rows
      const rows: { name: string; color: string; x: number }[][] = [[]];
      let x = 0;
      const maxW = right - left;
      for (const it of legendItems) {
        const w = Math.min(160, measure(it.name, fs)) + 28;
        if (x + w > maxW && rows[rows.length - 1].length) {
          rows.push([]);
          x = 0;
        }
        rows[rows.length - 1].push({ ...it, x });
        x += w;
      }
      const shownRows = rows.slice(0, 3);
      const lh = 18;
      const startY = chart.legend === 'top' ? top + 4 : bottom - shownRows.length * lh + 6;
      shownRows.forEach((row, ri) => {
        const rowW = row.length ? row[row.length - 1].x + Math.min(160, measure(row[row.length - 1].name, fs)) + 28 : 0;
        const ox = left + (maxW - rowW) / 2;
        for (const it of row) {
          const y = startY + ri * lh + 10;
          svg.append(el('rect', { x: ox + it.x, y: y - 9, width: 10, height: 10, rx: 2, fill: it.color }));
          svg.append(el('text', { x: ox + it.x + 16, y, 'font-size': fs, fill: t.muted }, truncate(it.name, 150, fs)));
        }
      });
      if (chart.legend === 'top') top += shownRows.length * lh + 8;
      else bottom -= shownRows.length * lh + 6;
    }
  }

  const hasData = data.series.some((s) => s.values.some((v) => v !== null));
  if (!hasData) {
    svg.append(el('text', { x: (left + right) / 2, y: (top + bottom) / 2, 'text-anchor': 'middle', 'font-size': 13, fill: t.muted }, 'No numbers to chart in this range'));
    return;
  }
  if (pie) renderPie(svg, chart, data, left, top, right, bottom, t);
  else renderCartesian(svg, chart, data, left, top, right, bottom, t);
}

function truncate(s: string, maxW: number, fs: number): string {
  if (measure(s, fs) <= maxW) return s;
  const n = Math.max(1, Math.floor(maxW / (fs * 0.56)) - 1);
  return s.slice(0, n) + '…';
}

function hit(e: SVGElement, info: HitInfo): void {
  e.setAttribute('data-hit', JSON.stringify(info));
}

function renderPie(svg: SVGSVGElement, chart: ChartDef, data: ResolvedChart, l: number, t0: number, r: number, b: number, t: ChartTheme): void {
  const s = data.series[0];
  const vals = s.values.map((v) => (v !== null && v > 0 ? v : 0));
  const total = vals.reduce((a, c) => a + c, 0);
  if (!total) return;
  const cx = (l + r) / 2;
  const cy = (t0 + b) / 2;
  const radius = Math.max(10, Math.min(r - l, b - t0) / 2 - 8);
  const inner = chart.type === 'doughnut' ? radius * 0.55 : 0;
  let angle = -Math.PI / 2;
  vals.forEach((v, i) => {
    if (!v) return;
    const a2 = angle + (v / total) * Math.PI * 2;
    const large = a2 - angle > Math.PI ? 1 : 0;
    const p = (ang: number, rad: number) => `${cx + Math.cos(ang) * rad},${cy + Math.sin(ang) * rad}`;
    let d: string;
    if (v === total) {
      d = `M ${cx - radius} ${cy} a ${radius} ${radius} 0 1 0 ${radius * 2} 0 a ${radius} ${radius} 0 1 0 ${-radius * 2} 0`;
      if (inner) d += ` M ${cx - inner} ${cy} a ${inner} ${inner} 0 1 1 ${inner * 2} 0 a ${inner} ${inner} 0 1 1 ${-inner * 2} 0`;
    } else if (inner) d = `M ${p(angle, radius)} A ${radius} ${radius} 0 ${large} 1 ${p(a2, radius)} L ${p(a2, inner)} A ${inner} ${inner} 0 ${large} 0 ${p(angle, inner)} Z`;
    else d = `M ${cx} ${cy} L ${p(angle, radius)} A ${radius} ${radius} 0 ${large} 1 ${p(a2, radius)} Z`;
    const color = CHART_COLORS[i % CHART_COLORS.length];
    const path = el('path', { d, fill: color, stroke: t.bg, 'stroke-width': 1.5, 'fill-rule': 'evenodd' });
    const pct = (v / total) * 100;
    hit(path, { series: s.name, category: data.categories[i] ?? String(i + 1), value: `${formatValue(v, s.format).text} (${pct.toFixed(1)}%)`, color });
    svg.append(path);
    if (pct >= 5) {
      const mid = (angle + a2) / 2;
      const lr = inner ? (radius + inner) / 2 : radius * 0.65;
      svg.append(el('text', { x: cx + Math.cos(mid) * lr, y: cy + Math.sin(mid) * lr + 4, 'text-anchor': 'middle', 'font-size': 11, 'font-weight': 600, fill: '#0b1220', 'pointer-events': 'none' }, `${Math.round(pct)}%`));
    }
    angle = a2;
  });
}

function renderCartesian(svg: SVGSVGElement, chart: ChartDef, data: ResolvedChart, l: number, t0: number, r: number, b: number, t: ChartTheme): void {
  const type = chart.type;
  const horizontal = type === 'bar';
  const scatter = type === 'scatter';
  const stacked = !!chart.stacked && !scatter;
  const n = Math.max(...data.series.map((s) => s.values.length), 0);
  const fmt = data.series[0]?.format;

  // Value range
  let vmin = Infinity;
  let vmax = -Infinity;
  if (stacked) {
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      for (const s of data.series) {
        const v = s.values[i] ?? 0;
        if (v >= 0) pos += v;
        else neg += v;
      }
      vmin = Math.min(vmin, neg);
      vmax = Math.max(vmax, pos);
    }
  } else {
    for (const s of data.series)
      for (const v of s.values) {
        if (v === null) continue;
        vmin = Math.min(vmin, v);
        vmax = Math.max(vmax, v);
      }
  }
  if (type !== 'line' && type !== 'scatter') {
    vmin = Math.min(0, vmin);
    vmax = Math.max(0, vmax);
  }
  const vs = niceScale(vmin, vmax);

  // X range for scatter
  let xs = { min: 0, max: 1, step: 0.2 };
  if (scatter) {
    const xvals = data.xs.filter((x): x is number => x !== null);
    xs = niceScale(Math.min(...xvals), Math.max(...xvals));
  }

  const fs = 11;
  const valLabels: string[] = [];
  for (let v = vs.min; v <= vs.max + vs.step / 2; v += vs.step) valLabels.push(compact(Number(v.toPrecision(12)), fmt));
  const valLabelW = Math.max(...valLabels.map((s) => measure(s, fs))) + 8;
  const catLabelW = horizontal ? Math.min((r - l) * 0.3, Math.max(...data.categories.map((c) => measure(c, fs))) + 10) : 0;

  let left = l + (horizontal ? catLabelW : valLabelW) + (chart.yTitle ? 18 : 0);
  let bottom = b - 20 - (chart.xTitle ? 18 : 0);
  const top = t0 + 6;
  const right = r - 6;
  if (horizontal) bottom = b - 20 - (chart.xTitle ? 18 : 0);
  if (left > right - 40) left = right - 40;

  // Axis titles
  if (chart.yTitle) {
    const yt = el('text', { x: l + 10, y: (top + bottom) / 2, 'text-anchor': 'middle', 'font-size': 11.5, fill: t.muted, transform: `rotate(-90 ${l + 10} ${(top + bottom) / 2})` }, chart.yTitle);
    svg.append(yt);
  }
  if (chart.xTitle) svg.append(el('text', { x: (left + right) / 2, y: b - 4, 'text-anchor': 'middle', 'font-size': 11.5, fill: t.muted }, chart.xTitle));

  const plotW = Math.max(10, right - left);
  const plotH = Math.max(10, bottom - top);
  const valToPx = (v: number) => (horizontal ? left + ((v - vs.min) / (vs.max - vs.min)) * plotW : bottom - ((v - vs.min) / (vs.max - vs.min)) * plotH);

  // Gridlines + value labels
  const grid = el('g');
  valLabels.forEach((label, i) => {
    const v = vs.min + i * vs.step;
    const p = valToPx(v);
    if (horizontal) {
      grid.append(el('line', { x1: p, x2: p, y1: top, y2: bottom, stroke: Math.abs(v) < 1e-12 ? t.axis : t.grid, 'stroke-width': 1 }));
      grid.append(el('text', { x: p, y: bottom + 15, 'text-anchor': 'middle', 'font-size': fs, fill: t.muted }, label));
    } else {
      grid.append(el('line', { x1: left, x2: right, y1: p, y2: p, stroke: Math.abs(v) < 1e-12 ? t.axis : t.grid, 'stroke-width': 1 }));
      grid.append(el('text', { x: left - 6, y: p + 4, 'text-anchor': 'end', 'font-size': fs, fill: t.muted }, label));
    }
  });
  svg.append(grid);

  if (scatter) {
    const xToPx = (x: number) => left + ((x - xs.min) / (xs.max - xs.min)) * plotW;
    for (let x = xs.min; x <= xs.max + xs.step / 2; x += xs.step) {
      const px = xToPx(x);
      svg.append(el('line', { x1: px, x2: px, y1: top, y2: bottom, stroke: t.grid }));
      svg.append(el('text', { x: px, y: bottom + 15, 'text-anchor': 'middle', 'font-size': fs, fill: t.muted }, compact(Number(x.toPrecision(12)))));
    }
    for (const s of data.series) {
      s.values.forEach((v, i) => {
        const x = data.xs[i];
        if (v === null || x === null || x === undefined) return;
        const c = el('circle', { cx: xToPx(x), cy: valToPx(v), r: 4.5, fill: s.color, 'fill-opacity': 0.85, stroke: t.bg, 'stroke-width': 1 });
        hit(c, { series: s.name, category: `x = ${compact(x)}`, value: formatValue(v, s.format).text, color: s.color });
        svg.append(c);
      });
    }
    return;
  }

  // Categories
  const band = (horizontal ? plotH : plotW) / Math.max(1, n);
  const catCenter = (i: number) => (horizontal ? top + band * (i + 0.5) : left + band * (i + 0.5));
  const maxLabels = Math.max(1, Math.floor((horizontal ? plotH / 16 : plotW / 60)));
  const every = Math.ceil(n / maxLabels);
  const labels = el('g');
  data.categories.slice(0, n).forEach((c, i) => {
    if (i % every) return;
    const p = catCenter(i);
    if (horizontal) labels.append(el('text', { x: left - 6, y: p + 4, 'text-anchor': 'end', 'font-size': fs, fill: t.muted }, truncate(c, catLabelW - 8, fs)));
    else labels.append(el('text', { x: p, y: bottom + 15, 'text-anchor': 'middle', 'font-size': fs, fill: t.muted }, truncate(c, band * every - 4, fs)));
  });
  svg.append(labels);
  // Axis line
  svg.append(horizontal ? el('line', { x1: left, x2: left, y1: top, y2: bottom, stroke: t.axis }) : el('line', { x1: left, x2: right, y1: bottom, y2: bottom, stroke: t.axis }));

  const zero = valToPx(Math.max(vs.min, Math.min(vs.max, 0)));
  if (type === 'column' || type === 'bar') {
    const groupW = band * 0.72;
    const k = data.series.length;
    const barW = stacked ? groupW : groupW / Math.max(1, k);
    const posAcc = new Array(n).fill(0);
    const negAcc = new Array(n).fill(0);
    data.series.forEach((s, si) => {
      s.values.forEach((v, i) => {
        if (v === null) return;
        let from = 0;
        if (stacked) {
          if (v >= 0) {
            from = posAcc[i];
            posAcc[i] += v;
          } else {
            from = negAcc[i];
            negAcc[i] += v;
          }
        }
        const p1 = valToPx(from);
        const p2 = valToPx(from + v);
        const offset = catCenter(i) - groupW / 2 + (stacked ? 0 : si * barW);
        const rect = horizontal
          ? el('rect', { x: Math.min(p1, p2), y: offset + 1, width: Math.max(1, Math.abs(p2 - p1)), height: Math.max(1, barW - 2), rx: 2, fill: s.color })
          : el('rect', { x: offset + 1, y: Math.min(p1, p2), width: Math.max(1, barW - 2), height: Math.max(1, Math.abs(p2 - p1)), rx: 2, fill: s.color });
        hit(rect, { series: s.name, category: data.categories[i] ?? String(i + 1), value: formatValue(v, s.format).text, color: s.color });
        svg.append(rect);
      });
    });
    void zero;
    return;
  }

  // Line & area
  const acc = new Array(n).fill(0);
  const layers = data.series.map((s) => {
    const pts = s.values.map((v, i) => {
      if (v === null) return null;
      const base = stacked ? acc[i] : 0;
      const y = stacked ? base + v : v;
      if (stacked) acc[i] += v;
      return { x: catCenter(i), y: valToPx(y), base: valToPx(base), v, i };
    });
    return { s, pts };
  });
  const order = type === 'area' && !stacked ? [...layers].reverse() : layers;
  for (const { s, pts } of order) {
    const valid = pts.filter((p): p is NonNullable<typeof p> => !!p);
    if (!valid.length) continue;
    if (type === 'area') {
      const d = `M ${valid[0].x} ${valid[0].base} ` + valid.map((p) => `L ${p.x} ${p.y}`).join(' ') + ' ' + [...valid].reverse().map((p) => `L ${p.x} ${p.base}`).join(' ') + ' Z';
      svg.append(el('path', { d, fill: s.color, 'fill-opacity': 0.3, stroke: 'none' }));
    }
    // Lines break at gaps
    let d = '';
    let pen = false;
    for (const p of pts) {
      if (!p) {
        pen = false;
        continue;
      }
      d += `${pen ? 'L' : 'M'} ${p.x} ${p.y} `;
      pen = true;
    }
    svg.append(el('path', { d, fill: 'none', stroke: s.color, 'stroke-width': 2.2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
    for (const p of valid) {
      const c = el('circle', { cx: p.x, cy: p.y, r: valid.length <= 40 ? 3.5 : 6, fill: valid.length <= 40 ? s.color : 'transparent', stroke: valid.length <= 40 ? t.bg : 'none', 'stroke-width': 1.2 });
      hit(c, { series: s.name, category: data.categories[p.i] ?? String(p.i + 1), value: formatValue(p.v, s.format).text, color: s.color });
      svg.append(c);
    }
  }
}
