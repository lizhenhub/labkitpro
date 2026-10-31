/* LabKitPro Excel add-in - charts.js (v1.3)
   Self-contained canvas chart engine, ported from the Chrome extension popup
   so the Excel task pane can render the SAME business-quality charts.

   In the Excel task pane we cannot show a live <canvas> the way the browser
   popup does, so every chart function returns a PNG dataURL string. The
   app injects it as <img src="data:image/png;base64,...">. Fully offline:
   no external libs, no network, WebView2-safe.

   Exposes window.LKPCharts = { capabilityChart, capabilityRun, spcChart,
     grrChart, grrRangeChart, grrAppChart, grrRangeMulti, varianceChart,
     runChart, devHist, fitChart, biasTrendChart, agreeBars, offscreenChart }. */
(function (root) {
  'use strict';

  // ---------- chart framework: axes, ticks, gridlines, axis titles ----------
  function niceStep(range, n) {
    const raw = range / Math.max(1, n);
    const mag = Math.pow(10, Math.floor(Math.log10(raw || 1)));
    for (const m of [1, 2, 2.5, 5, 10]) if (raw <= m * mag) return m * mag;
    return 10 * mag;
  }
  function tickFmt(v, step) {
    const d = step >= 1 ? 0 : Math.min(4, Math.max(0, Math.ceil(-Math.log10(step))));
    return String(+(+v).toFixed(d));
  }
  // draws a full frame (grid + tick labels + axes + titles); returns X/Y mappers
  function frame(ctx, W, H, o) {
    const L = o.L == null ? 64 : o.L, R = W - (o.R == null ? 14 : o.R);
    const T = o.T == null ? 26 : o.T, B = H - (o.B == null ? 54 : o.B);
    const X = (v) => L + ((v - o.x0) * (R - L)) / (o.x1 - o.x0);
    const Y = (v) => B - ((v - o.y0) * (B - T)) / (o.y1 - o.y0);
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    ctx.font = '13px sans-serif';
    // horizontal grid + y tick labels
    let ys = niceStep(o.y1 - o.y0, o.yN || 5);
    if (o.yInt) ys = Math.max(1, Math.ceil(ys));
    ctx.textAlign = 'right'; ctx.fillStyle = '#7a8090';
    for (let v = Math.ceil(o.y0 / ys) * ys; v <= o.y1 + ys * 1e-6; v += ys) {
      ctx.strokeStyle = '#eef0f4'; ctx.beginPath(); ctx.moveTo(L, Y(v)); ctx.lineTo(R, Y(v)); ctx.stroke();
      ctx.fillText(tickFmt(v, ys), L - 6, Y(v) + 4.5);
    }
    // x tick labels
    ctx.textAlign = 'center';
    if (o.xLabels) {
      const n = o.xLabels.length, every = Math.max(1, Math.ceil(n / 12));
      for (let i = 0; i < n; i += every) ctx.fillText(o.xLabels[i], X(i + 1), B + 18);
    } else {
      const xs = niceStep(o.x1 - o.x0, o.xN || 6);
      for (let v = Math.ceil(o.x0 / xs) * xs; v <= o.x1 + xs * 1e-6; v += xs) ctx.fillText(tickFmt(v, xs), X(v), B + 18);
    }
    // axes
    ctx.strokeStyle = '#b5bac4'; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(L, T); ctx.lineTo(L, B); ctx.lineTo(R, B); ctx.stroke();
    // axis titles
    ctx.fillStyle = '#444a56'; ctx.font = '14px sans-serif';
    if (o.xlabel) { ctx.textAlign = 'right'; ctx.fillText(o.xlabel, R, B + 34); }
    if (o.ylabel) {
      ctx.save(); ctx.translate(14, (T + B) / 2); ctx.rotate(-Math.PI / 2);
      ctx.textAlign = 'center'; ctx.fillText(o.ylabel, 0, 0); ctx.restore();
    }
    ctx.textAlign = 'left';
    return { X, Y, L, R, T, B };
  }
  // right-aligned label with translucent white backing (keeps text readable over lines/points)
  function labelBg(ctx, xRight, yBase, text, color) {
    ctx.font = 'bold 14px sans-serif';
    const w = ctx.measureText(text).width;
    ctx.fillStyle = 'rgba(255,255,255,0.85)';
    ctx.fillRect(xRight - w - 4, yBase - 13, w + 7, 18);
    ctx.fillStyle = color; ctx.textAlign = 'right';
    ctx.fillText(text, xRight, yBase);
    ctx.textAlign = 'left';
  }
  // dashed horizontal reference line with right-side label
  function hline(ctx, f, v, color, dash, label) {
    if (!isFinite(v)) return;
    ctx.strokeStyle = color; ctx.lineWidth = 1.6; ctx.setLineDash(dash || []);
    ctx.beginPath(); ctx.moveTo(f.L, f.Y(v)); ctx.lineTo(f.R, f.Y(v)); ctx.stroke(); ctx.setLineDash([]);
    if (label) labelBg(ctx, f.R - 4, f.Y(v) - 4, label, color);
  }
  // dashed vertical reference line with top label
  function vline(ctx, f, v, color, label, dash, yOff) {
    if (v == null || !isFinite(v)) return;
    ctx.strokeStyle = color; ctx.lineWidth = 1.8; ctx.setLineDash(dash || [6, 4]);
    ctx.beginPath(); ctx.moveTo(f.X(v), f.T); ctx.lineTo(f.X(v), f.B); ctx.stroke(); ctx.setLineDash([]);
    ctx.fillStyle = color; ctx.font = 'bold 15px sans-serif';
    ctx.fillText(label, Math.min(f.X(v) + 5, f.R - 60), f.T + 14 + (yOff || 0));
  }
  function ptsLine(ctx, f, xs, ys, color) {
    ctx.strokeStyle = color; ctx.lineWidth = 1.8; ctx.beginPath();
    xs.forEach((x, i) => (i ? ctx.lineTo(f.X(x), f.Y(ys[i])) : ctx.moveTo(f.X(x), f.Y(ys[i]))));
    ctx.stroke();
  }
  function drawPts(ctx, f, xs, ys, color, r) {
    ctx.fillStyle = color;
    xs.forEach((x, i) => { ctx.beginPath(); ctx.arc(f.X(x), f.Y(ys[i]), r || 4, 0, Math.PI * 2); ctx.fill(); });
  }
  // per-point value labels: auto-thinned to maxLabels, side chosen so labels never sit on the line
  function labelVals(ctx, f, xs, ys, maxLabels, belowIdx) {
    const n = ys.length;
    const step = Math.max(1, Math.ceil(n / (maxLabels || 16)));
    const yMid = (Math.min(...ys) + Math.max(...ys)) / 2;
    const forceBelow = new Set(belowIdx || []);
    for (let i = 0; i < n; i += step) {
      const y = ys[i], px = f.X(xs[i]), py = f.Y(y);
      const txt = String(+y.toFixed(3));
      ctx.font = 'bold 12px sans-serif';
      const w = ctx.measureText(txt).width;
      const up = !forceBelow.has(i) && y < yMid; // lower-half points label above, upper-half below
      const by = up ? py - 23 : py + 7;
      ctx.fillStyle = 'rgba(255,255,255,0.85)';
      ctx.fillRect(px - w / 2 - 3, by, w + 6, 15);
      ctx.fillStyle = '#33415c';
      ctx.fillText(txt, px - w / 2, by + 12);
    }
  }
  // generic offscreen canvas chart factory (720px print-quality dataURL)
  function offscreenChart(H, draw) {
    const W = 720, cv = document.createElement('canvas');
    cv.width = W; cv.height = H;
    const ctx = cv.getContext('2d');
    draw(ctx, W, H);
    return cv.toDataURL('image/png');
  }

  // ---------- SPC painter (shared renderer; also used by spcChart wrapper) ----------
  function paintSpc(ctx, W, H, pts, cl, ucl, lcl, viol) {
    const isU = Array.isArray(ucl), isL = Array.isArray(lcl);
    const U = (i) => (isU ? ucl[i] : ucl), Lo = (i) => (isL ? lcl[i] : lcl);
    let lo = Math.min(...pts, cl), hi = Math.max(...pts, cl);
    pts.forEach((_, i) => { lo = Math.min(lo, Lo(i)); hi = Math.max(hi, U(i)); });
    const pad = (hi - lo) * 0.12 || 1;
    const f = frame(ctx, W, H, { x0: 0.5, x1: pts.length + 0.5, y0: lo - pad, y1: hi + pad, xLabels: pts.map((_, i) => i + 1), xlabel: 'Subgroup', ylabel: 'Value' });
    // limit lines: stepped polylines when limits vary per point (p/u charts), straight lines otherwise
    const drawLim = (val, isArr, color, label) => {
      ctx.strokeStyle = color; ctx.lineWidth = 1.2; ctx.setLineDash([5, 3]); ctx.beginPath();
      if (isArr) val.forEach((v, i) => (i ? ctx.lineTo(f.X(i + 1), f.Y(v)) : ctx.moveTo(f.X(i + 1), f.Y(v))));
      else { ctx.moveTo(f.L, f.Y(val)); ctx.lineTo(f.R, f.Y(val)); }
      ctx.stroke(); ctx.setLineDash([]);
      const last = isArr ? val[val.length - 1] : val;
      labelBg(ctx, f.R - 4, label === 'LCL' ? f.Y(last) + 18 : f.Y(last) - 4, label + ' ' + (+last.toFixed(4)), color);
    };
    drawLim(ucl, isU, '#b03a2e', 'UCL');
    drawLim(cl, false, '#1d7a45', 'CL');
    drawLim(lcl, isL, '#b03a2e', 'LCL');
    ptsLine(ctx, f, pts.map((_, i) => i + 1), pts, '#0c447c');
    drawPts(ctx, f, pts.map((_, i) => i + 1), pts, '#0c447c', 4);
    labelVals(ctx, f, pts.map((_, i) => i + 1), pts, W < 500 ? 8 : 14, (viol || []).map((v) => v.i));
    // violations: red ring + rule tag
    const vmap = {};
    (viol || []).forEach((v) => { (vmap[v.i] = vmap[v.i] || []).push(v.rule); });
    ctx.textAlign = 'center';
    Object.keys(vmap).forEach((k) => {
      const idx = +k, p = pts[idx];
      ctx.beginPath(); ctx.arc(f.X(idx + 1), f.Y(p), 6.5, 0, Math.PI * 2);
      ctx.fillStyle = '#b03a2e'; ctx.fill();
      ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.6; ctx.stroke();
      const tags = [...new Set(vmap[idx])].map((r) => 'R' + r).join(' ');
      ctx.fillStyle = '#b03a2e'; ctx.font = 'bold 13px sans-serif';
      ctx.fillText(tags, f.X(idx + 1), f.Y(p) - 13);
    });
    ctx.textAlign = 'left';
  }
  // convenience: SPC control chart as PNG dataURL (Excel injection)
  function spcChart(pts, cl, ucl, lcl, viol, H) {
    return offscreenChart(H || 360, (ctx, W, H2) => paintSpc(ctx, W, H2, pts, cl, ucl, lcl, viol));
  }

  // ---------- histogram core (bars + axes; curve overlaid by callers) ----------
  function histogram(ctx, W, H, vals, x0, x1, xlabel) {
    const nbins = Math.min(14, Math.max(6, Math.ceil(Math.sqrt(vals.length))));
    const bw = (x1 - x0) / nbins;
    const bins = new Array(nbins).fill(0);
    vals.forEach((v) => { bins[Math.min(nbins - 1, Math.max(0, Math.floor((v - x0) / bw)))]++; });
    const bmax = Math.max(...bins);
    const f = frame(ctx, W, H, { x0, x1, y0: 0, y1: bmax * 1.12 || 1, yN: 5, yInt: true, xlabel: xlabel || '', ylabel: 'Frequency' });
    bins.forEach((c, i) => {
      const bx = f.X(x0 + i * bw), by = f.Y(c), wb = (f.R - f.L) / nbins - 2;
      ctx.fillStyle = '#b5d4f4'; ctx.fillRect(bx + 1, by, wb, f.B - by);
      ctx.strokeStyle = '#7fb2e3'; ctx.strokeRect(bx + 1, by, wb, f.B - by);
    });
    return { f, bins, bw };
  }

  // Capability histogram: bars + fitted normal curve + LSL/USL/mean lines
  function capabilityChart(vals, lsl, usl, mean, std) {
    return offscreenChart(380, (ctx, W, H) => {
      const lo = Math.min(...vals, lsl != null ? lsl : Infinity, usl != null ? usl : Infinity);
      const hi = Math.max(...vals, lsl != null ? lsl : -Infinity, usl != null ? usl : -Infinity);
      const pad = (hi - lo) * 0.10 || 1;
      const x0 = lo - pad, x1 = hi + pad;
      const { f, bins, bw } = histogram(ctx, W, H, vals, x0, x1, 'Measured value');
      // normal curve scaled to bar units
      const k = (Math.max(...bins) * bw) / (std * Math.sqrt(2 * Math.PI));
      ctx.strokeStyle = '#0c447c'; ctx.lineWidth = 2; ctx.beginPath();
      for (let px = x0; px <= x1; px += (x1 - x0) / 200) {
        const y = f.Y(k * Math.exp(-((px - mean) ** 2) / (2 * std * std)));
        px === x0 ? ctx.moveTo(f.X(px), y) : ctx.lineTo(f.X(px), y);
      }
      ctx.stroke(); ctx.lineWidth = 1;
      vline(ctx, f, lsl, '#b03a2e', 'LSL');
      vline(ctx, f, usl, '#b03a2e', 'USL');
      vline(ctx, f, mean, '#1d7a45', 'x̄', []);
    });
  }
  // Run chart of individual values in data order vs spec limits (companion to the histogram)
  function capabilityRun(vals, lsl, usl, mean) {
    return offscreenChart(300, (ctx, W, H) => {
      const xs = vals.map((_, i) => i + 1);
      const lo = Math.min(...vals, lsl != null ? lsl : Infinity, usl != null ? usl : Infinity);
      const hi = Math.max(...vals, lsl != null ? lsl : -Infinity, usl != null ? usl : -Infinity);
      const pad = (hi - lo) * 0.12 || 1;
      const f = frame(ctx, W, H, { x0: 0.5, x1: vals.length + 0.5, y0: lo - pad, y1: hi + pad, xLabels: xs, xlabel: 'Run (data order)', ylabel: 'Measured value' });
      hline(ctx, f, mean, '#1d7a45', [], 'Mean');
      hline(ctx, f, lsl, '#b03a2e', [5, 3], 'LSL');
      hline(ctx, f, usl, '#b03a2e', [5, 3], 'USL');
      ptsLine(ctx, f, xs, vals, '#0c447c');
      drawPts(ctx, f, xs, vals, '#0c447c', 4);
      labelVals(ctx, f, xs, vals, 15);
    });
  }
  // Run chart of repeated measurements with reference line (bias / Cg studies)
  function runChart(vals, ref, meanV, stdV) {
    return offscreenChart(320, (ctx, W, H) => {
      const xs = vals.map((_, i) => i + 1);
      const extras = meanV != null && stdV > 0 ? [meanV - 3 * stdV, meanV + 3 * stdV] : [];
      const lo = Math.min(...vals, ref, ...extras), hi = Math.max(...vals, ref, ...extras);
      const pad = (hi - lo) * 0.15 || 1;
      const f = frame(ctx, W, H, { x0: 0.5, x1: vals.length + 0.5, y0: lo - pad, y1: hi + pad, xLabels: xs, xlabel: 'Run #', ylabel: 'Measured value' });
      hline(ctx, f, ref, '#1d7a45', [5, 3], 'Reference');
      if (meanV != null && stdV > 0) {
        hline(ctx, f, meanV + 3 * stdV, '#ef9f27', [4, 3], 'x̄+3s');
        hline(ctx, f, meanV, '#5b6b7d', [], 'x̄');
        hline(ctx, f, meanV - 3 * stdV, '#ef9f27', [4, 3], 'x̄−3s');
      }
      ptsLine(ctx, f, xs, vals, '#0c447c');
      drawPts(ctx, f, xs, vals, '#0c447c', 4);
      labelVals(ctx, f, xs, vals, 20);
    });
  }
  // Distribution of deviations from the reference (bias / Cg companion chart)
  function devHist(vals, ref, biasVal) {
    return offscreenChart(300, (ctx, W, H) => {
      const devs = vals.map((v) => v - ref);
      const lo = Math.min(...devs, 0), hi = Math.max(...devs, 0);
      const pad = (hi - lo) * 0.10 || 1;
      const { f } = histogram(ctx, W, H, devs, lo - pad, hi + pad, 'Deviation from reference (measured − ref)');
      const close = Math.abs(f.X(biasVal) - f.X(0)) < 110; // labels would overlap → stack rows
      vline(ctx, f, 0, '#7a8090', close ? '0' : '0 (no bias)', [4, 3]);
      if (biasVal != null && isFinite(biasVal)) vline(ctx, f, biasVal, '#ef9f27', 'bias', null, close ? 36 : 0);
    });
  }
  // Measurements by part, one line per trial (Gage R&R repeatability view)
  function grrChart(parts) {
    return offscreenChart(340, (ctx, W, H) => {
      const all = parts.flat();
      const lo = Math.min(...all), hi = Math.max(...all);
      const pad = (hi - lo) * 0.15 || 1;
      const f = frame(ctx, W, H, { x0: 0.5, x1: parts.length + 0.5, y0: lo - pad, y1: hi + pad, xLabels: parts.map((_, i) => 'P' + (i + 1)), xlabel: 'Part', ylabel: 'Measured value' });
      const colors = ['#0c447c', '#ef9f27', '#1d7a45'];
      const trials = parts[0].length;
      for (let t = 0; t < trials; t++) {
        const ys = parts.map((p) => p[t]);
        ptsLine(ctx, f, parts.map((_, i) => i + 1), ys, colors[t % 3]);
        drawPts(ctx, f, parts.map((_, i) => i + 1), ys, colors[t % 3], 3.2);
      }
      ctx.font = 'bold 14px sans-serif';
      for (let t = 0; t < Math.min(3, trials); t++) {
        ctx.fillStyle = colors[t % 3];
        ctx.fillText('Trial ' + (t + 1), f.L + 8 + t * 76, f.T + 12);
      }
    });
  }
  // Range within part vs R̄ and UCL = D4·R̄ (AIAG repeatability check)
  function grrRangeChart(parts) {
    return offscreenChart(300, (ctx, W, H) => {
      const trials = parts[0].length;
      const D4v = trials === 2 ? 3.267 : 2.574;
      const ranges = parts.map((p) => Math.max(...p) - Math.min(...p));
      const rbar = ranges.reduce((s, r) => s + r, 0) / ranges.length;
      const ucl = D4v * rbar;
      const hi = Math.max(...ranges, ucl);
      const f = frame(ctx, W, H, { x0: 0.5, x1: parts.length + 0.5, y0: 0, y1: hi * 1.1 || 1, xLabels: parts.map((_, i) => 'P' + (i + 1)), xlabel: 'Part', ylabel: 'Range (max − min)' });
      const bwid = (f.R - f.L) / parts.length * 0.5;
      ranges.forEach((r, i) => {
        const cx = f.X(i + 1);
        ctx.fillStyle = r > ucl ? '#b03a2e' : '#7fb2e3';
        ctx.fillRect(cx - bwid / 2, f.Y(r), bwid, f.B - f.Y(r));
      });
      hline(ctx, f, ucl, '#b03a2e', [5, 3], 'UCL');
      hline(ctx, f, rbar, '#1d7a45', [], 'R̄');
    });
  }
  const APP_COLORS = ['#0c447c', '#ef9f27', '#1d7a45'];
  // AIAG classic: every measurement as a point, colored by appraiser, + appraiser mean profile per part
  function grrAppChart(appData) {
    return offscreenChart(380, (ctx, W, H) => {
      const all = appData.flatMap((a) => a.trials.flat());
      const lo = Math.min(...all), hi = Math.max(...all);
      const pad = (hi - lo) * 0.10 || 1;
      const nP = appData[0].trials[0].length;
      const f = frame(ctx, W, H, { x0: 0.5, x1: nP + 0.5, y0: lo - pad, y1: hi + pad, xLabels: Array.from({ length: nP }, (_, i) => 'P' + (i + 1)), xlabel: 'Part', ylabel: 'Measured value' });
      appData.forEach((a, ai) => {
        const col = APP_COLORS[ai % 3];
        const means = a.trials[0].map((_, p) => a.trials.reduce((s, t) => s + t[p], 0) / a.trials.length);
        ptsLine(ctx, f, means.map((_, i) => i + 1), means, col);
        a.trials.forEach((t) => drawPts(ctx, f, t.map((_, i) => i + 1), t, col, 3.4));
      });
      // legend
      ctx.font = 'bold 14px sans-serif';
      appData.forEach((a, ai) => {
        ctx.fillStyle = APP_COLORS[ai % 3];
        ctx.fillText('Appraiser ' + a.name, f.L + 8 + ai * 92, f.T + 12);
      });
    });
  }
  // within-part range bars grouped by appraiser, with R̄̄ and UCL = D4·R̄̄
  function grrRangeMulti(appData, rbar, ucl) {
    return offscreenChart(320, (ctx, W, H) => {
      const nP = appData[0].ranges.length;
      const hi = Math.max(...appData.flatMap((a) => a.ranges), ucl);
      const f = frame(ctx, W, H, { x0: 0.5, x1: nP + 0.5, y0: 0, y1: hi * 1.1 || 1, xLabels: Array.from({ length: nP }, (_, i) => 'P' + (i + 1)), xlabel: 'Part', ylabel: 'Range within part' });
      const slot = (f.R - f.L) / nP;
      const bwid = Math.min(slot / (appData.length + 1), 16);
      appData.forEach((a, ai) => {
        a.ranges.forEach((r, p) => {
          const cx = f.X(p + 1) - slot / 2 + slot * (ai + 0.5) / appData.length;
          ctx.fillStyle = r > ucl ? '#b03a2e' : APP_COLORS[ai % 3];
          ctx.fillRect(cx - bwid / 2, f.Y(r), bwid, f.B - f.Y(r));
        });
      });
      hline(ctx, f, ucl, '#b03a2e', [5, 3], 'UCL = D4·R̄̄');
      hline(ctx, f, rbar, '#1d7a45', [], 'R̄̄');
      ctx.font = 'bold 13px sans-serif';
      appData.forEach((a, ai) => {
        ctx.fillStyle = APP_COLORS[ai % 3];
        ctx.fillText(a.name, f.L + 8 + ai * 24, f.T + 12);
      });
    });
  }
  // variance components: %EV / %AV / %GRR / %PV with 10% and 30% reference lines
  function varianceChart(r) {
    return offscreenChart(300, (ctx, W, H) => {
      const items = [['%EV', r.pctEV], ['%AV', r.pctAV], ['%GRR', r.pctGRR], ['%PV', r.pctPV]].filter(([, v]) => v != null && isFinite(v));
      const hi = Math.max(...items.map(([, v]) => v), 30) * 1.15;
      const f = frame(ctx, W, H, { x0: 0.5, x1: items.length + 0.5, y0: 0, y1: hi, yN: 5, xLabels: items.map(([k]) => k), xlabel: '', ylabel: '% of tolerance' });
      const bwid = Math.min(90, (f.R - f.L) / items.length * 0.5);
      ctx.textAlign = 'center';
      items.forEach(([k, v], i) => {
        const cx = f.X(i + 1);
        const col = k === '%GRR' ? (v <= 10 ? '#1d7a45' : v <= 30 ? '#ef9f27' : '#b03a2e') : '#7fb2e3';
        ctx.fillStyle = col;
        ctx.fillRect(cx - bwid / 2, f.Y(v), bwid, f.B - f.Y(v));
        ctx.fillStyle = '#0c447c'; ctx.font = 'bold 15px sans-serif';
        ctx.fillText(v + '%', cx, f.Y(v) - 6);
      });
      ctx.textAlign = 'left';
      hline(ctx, f, 10, '#1d7a45', [5, 3], '10%');
      hline(ctx, f, 30, '#b03a2e', [5, 3], '30%');
    });
  }
  // Bias vs reference with fitted trend (linearity)
  function fitChart(pts, slope, intercept, r2) {
    return offscreenChart(340, (ctx, W, H) => {
      const xs = pts.map((p) => p.ref), ys = pts.map((p) => p.bias);
      const y0 = Math.min(...ys, 0), y1 = Math.max(...ys, 0);
      const xpad = ((Math.max(...xs) - Math.min(...xs)) || 1) * 0.1;
      const ypad = ((y1 - y0) || 1) * 0.2;
      const f = frame(ctx, W, H, { x0: Math.min(...xs) - xpad, x1: Math.max(...xs) + xpad, y0: y0 - ypad, y1: y1 + ypad, xlabel: 'Reference value', ylabel: 'Bias (measured − ref)' });
      hline(ctx, f, 0, '#b5bac4', [4, 3], '0');
      ctx.strokeStyle = '#ef9f27'; ctx.lineWidth = 2; ctx.beginPath();
      const xr0 = Math.min(...xs) - xpad, xr1 = Math.max(...xs) + xpad;
      ctx.moveTo(f.X(xr0), f.Y(intercept + slope * xr0)); ctx.lineTo(f.X(xr1), f.Y(intercept + slope * xr1));
      ctx.stroke(); ctx.lineWidth = 1;
      drawPts(ctx, f, xs, ys, '#0c447c', 5);
      labelVals(ctx, f, xs, ys, 10);
      ctx.fillStyle = '#444a56'; ctx.font = '14px sans-serif';
      ctx.fillText('slope = ' + slope + ', R² = ' + r2, f.L + 8, f.T + 14);
    });
  }
  // Bias over time (stability): X̄ chart of time-group means with ±A₂R̄ limits when available
  function biasTrendChart(r, ref) {
    return offscreenChart(320, (ctx, W, H) => {
      const ys = r.means || r.rows.map((x) => x.mean);
      const xs = ys.map((_, i) => i + 1);
      const hasLim = r.ucl != null;
      const lo = Math.min(...ys, hasLim ? r.lcl : 0), hi = Math.max(...ys, hasLim ? r.ucl : 0);
      const pad = ((hi - lo) || 1) * 0.18;
      const f = frame(ctx, W, H, { x0: 0.5, x1: ys.length + 0.5, y0: lo - pad, y1: hi + pad, xLabels: r.rows.map((x) => x.label), xlabel: 'Time order', ylabel: 'Group mean' });
      if (hasLim) {
        hline(ctx, f, r.ucl, '#b03a2e', [5, 3], 'UCL');
        hline(ctx, f, r.cl, '#1d7a45', [], 'CL');
        hline(ctx, f, r.lcl, '#b03a2e', [5, 3], 'LCL');
      } else {
        hline(ctx, f, ref, '#1d7a45', [5, 3], 'Reference');
      }
      ptsLine(ctx, f, xs, ys, '#0c447c');
      drawPts(ctx, f, xs, ys, '#0c447c', 4.5);
      // violation rings
      const vmap = {};
      (r.violations || []).forEach((v) => { (vmap[v.i] = vmap[v.i] || []).push(v.rule); });
      ctx.textAlign = 'center';
      Object.keys(vmap).forEach((k) => {
        const idx = +k;
        ctx.beginPath(); ctx.arc(f.X(idx + 1), f.Y(ys[idx]), 6.5, 0, Math.PI * 2);
        ctx.fillStyle = '#b03a2e'; ctx.fill();
        ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.6; ctx.stroke();
      });
      ctx.textAlign = 'left';
      labelVals(ctx, f, xs, ys, 10, Object.keys(vmap).map(Number));
    });
  }
  // Agreement bars per appraiser with 90% target (attribute agreement)
  function agreeBars(perApp) {
    return offscreenChart(320, (ctx, W, H) => {
      const n = perApp.length;
      const f = frame(ctx, W, H, { x0: 0.5, x1: n + 0.5, y0: 0, y1: 105, yN: 5, xLabels: perApp.map((a) => 'Appraiser ' + a.app), xlabel: '', ylabel: 'Agreement with reference (%)' });
      const bwid = Math.min(90, (f.R - f.L) / n * 0.45);
      ctx.textAlign = 'center';
      perApp.forEach((a, i) => {
        const cx = f.X(i + 1);
        ctx.fillStyle = a.pct >= 90 ? '#1d7a45' : '#b03a2e';
        ctx.fillRect(cx - bwid / 2, f.Y(a.pct), bwid, f.B - f.Y(a.pct));
        ctx.fillStyle = '#0c447c'; ctx.font = 'bold 15px sans-serif';
        ctx.fillText(a.pct + '%', cx, f.Y(a.pct) - 6);
      });
      ctx.textAlign = 'left';
      hline(ctx, f, 90, '#ef9f27', [5, 3], '90% target');
    });
  }

  root.LKPCharts = {
    capabilityChart, capabilityRun, spcChart,
    grrChart, grrRangeChart, grrAppChart, grrRangeMulti, varianceChart,
    runChart, devHist, fitChart, biasTrendChart, agreeBars, offscreenChart, paintSpc,
  };
})(typeof self !== 'undefined' ? self : this);
