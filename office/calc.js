/* LabKitPro Quality Tools — core calculation library (pure JS, offline).
   Shared by Chrome extension v1 and (future) Excel add-in. No dependencies. */
(function (root) {
  'use strict';

  // ---------- basic stats ----------
  const mean = (a) => a.reduce((s, x) => s + x, 0) / a.length;
  const sampleStd = (a) => {
    if (a.length < 2) return 0;
    const m = mean(a);
    return Math.sqrt(a.reduce((s, x) => s + (x - m) * (x - m), 0) / (a.length - 1));
  };
  const round = (x, d = 3) => (isFinite(x) ? +x.toFixed(d) : NaN);

  // ---------- capability (AIAG dual basis, same engine as the website CpkSuite) ----------
  // values: flat readings. subgroups: rows of equal length (2–10) → within σ = R̄/d₂ → Cp/Cpk;
  // overall s → Pp/Ppk. Without subgroup structure only the overall basis is shown.
  function capability(values, lsl, usl, target, subgroups) {
    const n = values.length;
    if (n < 2) return { error: 'Need at least 2 values' };
    const m = mean(values);
    const s = sampleStd(values); // overall (long-term) sigma
    if (s === 0) return { error: 'All values identical — zero variation' };
    const twoSided = lsl !== null && usl !== null;
    let cp = null, cpk = null, cpu = null, cpl = null, pp = null, ppk = null, cpm = null;
    if (twoSided) { pp = (usl - lsl) / (6 * s); ppk = Math.min(usl - m, m - lsl) / (3 * s); }
    else if (usl !== null) ppk = (usl - m) / (3 * s);
    else ppk = (m - lsl) / (3 * s);
    let sigW = null;
    if (Array.isArray(subgroups) && subgroups.length >= 2) {
      const k0 = subgroups[0].length;
      if (k0 >= 2 && k0 <= 10 && subgroups.every((g) => g.length === k0)) {
        const rbar = mean(subgroups.map((g) => Math.max(...g) - Math.min(...g)));
        sigW = rbar / D2[k0];
      }
    }
    if (sigW) {
      if (twoSided) {
        cp = (usl - lsl) / (6 * sigW);
        cpu = (usl - m) / (3 * sigW);
        cpl = (m - lsl) / (3 * sigW);
        cpk = Math.min(cpu, cpl);
      } else if (usl !== null) { cpu = (usl - m) / (3 * sigW); cpk = cpu; }
      else { cpl = (m - lsl) / (3 * sigW); cpk = cpl; }
    }
    if (target != null && isFinite(target) && twoSided)
      cpm = (usl - lsl) / (6 * Math.sqrt(s * s + (m - target) ** 2));
    const eff = cpk != null ? cpk : ppk;
    const ppm = twoSided
      ? 1e6 * (normCdf((lsl - m) / s) + 1 - normCdf((usl - m) / s))
      : usl !== null ? 1e6 * (1 - normCdf((usl - m) / s)) : 1e6 * normCdf((lsl - m) / s);
    return {
      n, mean: round(m), std: round(s), sigW: sigW != null ? round(sigW) : null,
      cp: cp != null ? round(cp) : null, cpk: cpk != null ? round(cpk) : null,
      cpu: cpu != null ? round(cpu) : null, cpl: cpl != null ? round(cpl) : null,
      pp: pp != null ? round(pp) : null, ppk: round(ppk), cpm: cpm != null ? round(cpm) : null,
      ppm: Math.round(ppm), overallOnly: !sigW,
      verdict: eff >= 1.33 ? 'Capable (≥ 1.33)' : eff >= 1.0 ? 'Marginal (1.00–1.33)' : 'Not capable (< 1.00)',
      ok: eff >= 1.33,
    };
  }
  function normCdf(z) { return 0.5 * (1 + erf(z / Math.SQRT2)); }
  function erf(x) { // Abramowitz-Stegun 7.1.26
    const s = x < 0 ? -1 : 1; x = Math.abs(x);
    const t = 1 / (1 + 0.3275911 * x);
    const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x);
    return s * y;
  }

  // ---------- SPC: 8 chart types ----------
  // Constants (AIAG / ASTM)
  const D3 = { 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0.076, 8: 0.136, 9: 0.184, 10: 0.223 };
  const D4 = { 2: 3.267, 3: 2.574, 4: 2.282, 5: 2.115, 6: 2.004, 7: 1.924, 8: 1.864, 9: 1.816, 10: 1.777 };
  const D2 = { 2: 1.128, 3: 1.693, 4: 2.059, 5: 2.326, 6: 2.534, 7: 2.704, 8: 2.847, 9: 2.97, 10: 3.078 };
  const A2 = { 2: 1.88, 3: 1.023, 4: 0.729, 5: 0.577, 6: 0.483, 7: 0.419, 8: 0.373, 9: 0.337, 10: 0.308 };
  // Ã₂ (median chart coefficient, AIAG SPC Median Charts table) — applied directly to R̄.
  // Using A2 here understates the limits ~42% at n=5 and mass-alarms (engine-audit fix).
  const M3 = { 2: 1.88, 3: 1.187, 4: 0.796, 5: 0.691, 6: 0.549, 7: 0.509, 8: 0.432, 9: 0.412, 10: 0.363 };
  const A3 = { 2: 2.659, 3: 1.954, 4: 1.628, 5: 1.427, 6: 1.287, 7: 1.182, 8: 1.099, 9: 1.032, 10: 0.975 };
  const B3 = { 2: 0, 3: 0, 4: 0, 5: 0, 6: 0.03, 7: 0.118, 8: 0.185, 9: 0.239, 10: 0.284 };
  const B4 = { 2: 3.267, 3: 2.568, 4: 2.266, 5: 2.089, 6: 1.97, 7: 1.882, 8: 1.815, 9: 1.761, 10: 1.716 };
  const C4 = { 2: 0.7979, 3: 0.8862, 4: 0.9213, 5: 0.94, 6: 0.9515, 7: 0.9594, 8: 0.965, 9: 0.9693, 10: 0.9727 };

  const SPC_CHARTS = ['xbar-r', 'xbar-s', 'i-mr', 'p', 'np', 'c', 'u', 'xbar-r-median'];

  function spcLimits(type, groups) {
    // groups: array of arrays (variable) or array of {n, defects} (attribute)
    const k = groups.length;
    if (k < 2) return { error: 'Need at least 2 subgroups' };
    const out = { type, k };
    const allVals = [];
    groups.forEach((g) => { if (Array.isArray(g)) allVals.push(...g); });

    if (type === 'xbar-r' || type === 'xbar-r-median') {
      const sizes = groups.map((g) => g.length);
      const n0 = sizes[0];
      if (!A2[n0]) return { error: 'Constant subgroup size 2–10 required' };
      const xbars = groups.map(mean), ranges = groups.map((g) => Math.max(...g) - Math.min(...g));
      const xb = mean(xbars), rb = mean(ranges);
      const isMedian = type === 'xbar-r-median';
      const meds = isMedian ? groups.map(medianOf) : null;
      out.points = isMedian ? meds : xbars;
      out.cl = isMedian ? mean(meds) : xb;
      // median chart: Ã₂×R̄ per AIAG (NOT A2 — that understates limits ~42% at n=5)
      const kCoef = isMedian ? M3[n0] : A2[n0];
      out.ucl = out.cl + kCoef * rb; out.lcl = out.cl - kCoef * rb;
      out.extra = { chart: 'R', points: ranges, cl: rb, ucl: D4[n0] * rb, lcl: D3[n0] * rb };
      out.sizeNote = 'subgroup size ' + n0;
    } else if (type === 'xbar-s') {
      const n0 = groups[0].length;
      if (!A3[n0]) return { error: 'Constant subgroup size 2–10 required' };
      const xbars = groups.map(mean), ss = groups.map(sampleStd);
      const xb = mean(xbars), sb = mean(ss);
      out.points = xbars; out.cl = xb; out.ucl = xb + A3[n0] * sb; out.lcl = xb - A3[n0] * sb;
      out.extra = { chart: 's', points: ss, cl: sb, ucl: B4[n0] * sb, lcl: B3[n0] * sb };
      out.sizeNote = 'subgroup size ' + n0;
    } else if (type === 'i-mr') {
      const x = allVals;
      if (x.length < 4) return { error: 'Need at least 4 individual values' };
      const mr = []; for (let i = 1; i < x.length; i++) mr.push(Math.abs(x[i] - x[i - 1]));
      const xb = mean(x), mrbar = mean(mr);
      out.points = x; out.cl = xb; out.ucl = xb + 2.66 * mrbar; out.lcl = xb - 2.66 * mrbar;
      out.extra = { chart: 'MR', points: mr, cl: mrbar, ucl: 3.267 * mrbar, lcl: 0 };
      out.sizeNote = 'individuals (n=' + x.length + ')';
    } else if (type === 'p' || type === 'np' || type === 'c' || type === 'u') {
      // attribute input: {n, d} per subgroup (d = defects/defectives)
      // note: c-chart counts defects per unit — d may exceed n (1 unit), so only defectives are bounded
      if (k < 4) return { error: 'Need at least 4 subgroups' };
      const valid = groups.every((g) => g && g.n > 0 && g.d !== undefined && g.d >= 0 && (type === 'c' || g.d <= g.n));
      if (!valid) return { error: 'Each subgroup needs sample size n and defect count d' };
      const ns = groups.map((g) => g.n), ds = groups.map((g) => g.d);
      const N = ns.reduce((a, b) => a + b, 0), D = ds.reduce((a, b) => a + b, 0);
      if (type === 'p') {
        const pb = D / N;
        out.points = ds.map((d, i) => d / ns[i]); out.cl = pb;
        out.ucl = ds.map((_, i) => pb + 3 * Math.sqrt(pb * (1 - pb) / ns[i]));
        out.lcl = ds.map((_, i) => Math.max(0, pb - 3 * Math.sqrt(pb * (1 - pb) / ns[i])));
      } else if (type === 'np') {
        const n0 = ns[0]; if (!ns.every((n) => n === n0)) return { error: 'np chart needs constant n' };
        const pb = D / N;
        out.points = ds; out.cl = n0 * pb;
        out.ucl = n0 * pb + 3 * Math.sqrt(n0 * pb * (1 - pb));
        out.lcl = Math.max(0, n0 * pb - 3 * Math.sqrt(n0 * pb * (1 - pb)));
      } else if (type === 'c') {
        const cb = D / k;
        out.points = ds; out.cl = cb; out.ucl = cb + 3 * Math.sqrt(cb); out.lcl = Math.max(0, cb - 3 * Math.sqrt(cb));
      } else { // u
        const ub = D / N;
        out.points = ds.map((d, i) => d / ns[i]); out.cl = ub;
        out.ucl = ds.map((_, i) => ub + 3 * Math.sqrt(ub / ns[i]));
        out.lcl = ds.map((_, i) => Math.max(0, ub - 3 * Math.sqrt(ub / ns[i])));
      }
      out.sizeNote = type === 'p' || type === 'u' ? 'variable n allowed' : 'constant n';
    } else {
      return { error: 'Unknown chart type' };
    }
    out.violations = nelson14(out.points, out.cl, out.ucl, out.lcl);
    return out;
  }
  function medianOf(a) { const s = [...a].sort((x, y) => x - y); const m = Math.floor(s.length / 2); return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }

  // Western Electric rules (same set as the website SpcSuite):
  // ① beyond limits ② 2 of 3 beyond 2σ ③ 4 of 5 beyond 1σ ④ 8 same side ⑤ 6 trending
  // Variable-limit charts (p/u): each point standardized by its own σ = (UCL−CL)/3.
  function nelson14(pts, cl, ucl, lcl) {
    const v = [];
    const isArr = Array.isArray(ucl);
    const U = (i) => (isArr ? ucl[i] : ucl), L = (i) => (isArr ? lcl[i] : lcl);
    const z = (i) => { const sig = (U(i) - cl) / 3; return sig > 0 ? (pts[i] - cl) / sig : 0; };
    pts.forEach((p, i) => { if (p > U(i) || p < L(i)) v.push({ i, rule: 1, txt: 'Beyond control limits' }); });
    for (let i = 2; i < pts.length; i++) {
      for (const side of [1, -1]) {
        const w = [z(i), z(i - 1), z(i - 2)];
        if (w.filter((x) => x * side > 2).length >= 2) { v.push({ i, rule: 2, txt: '2 of 3 beyond 2σ' }); break; }
      }
    }
    for (let i = 4; i < pts.length; i++) {
      for (const side of [1, -1]) {
        const w = [z(i), z(i - 1), z(i - 2), z(i - 3), z(i - 4)];
        if (w.filter((x) => x * side > 1).length >= 4) { v.push({ i, rule: 3, txt: '4 of 5 beyond 1σ' }); break; }
      }
    }
    for (let i = 7; i < pts.length; i++) {
      const side = Math.sign(pts[i] - cl);
      if (side !== 0) {
        let ok = true;
        for (let j = i - 7; j <= i; j++) if (Math.sign(pts[j] - cl) !== side) { ok = false; break; }
        if (ok) v.push({ i, rule: 4, txt: '8 consecutive on one side of CL' });
      }
    }
    for (let i = 5; i < pts.length; i++) {
      const w = pts.slice(i - 5, i + 1);
      if (w.every((x, kk) => kk === 0 || x > w[kk - 1]) || w.every((x, kk) => kk === 0 || x < w[kk - 1])) v.push({ i, rule: 5, txt: '6 consecutive trending' });
    }
    return v;
  }

  // ---------- MSA: 5 studies ----------
  const K1 = { 2: 0.8862, 3: 0.5906 }; // 1/d2 by trials
  const K2 = { 2: 0.7071, 3: 0.5231 }; // by appraisers
  const K3 = { 2: 0.7071, 3: 0.5231, 4: 0.4467, 5: 0.4030, 6: 0.3742, 7: 0.3534, 8: 0.3375, 9: 0.3249, 10: 0.3146 };

  function gageRR(trials, tolerance) {
    // trials: [appraiser][trial][part] -> measurements (1–3 appraisers, 2–3 trials, 3–10 parts)
    const appraisers = trials.length;
    const nTrials = trials[0].length;
    const nParts = trials[0][0].length;
    if (!K1[nTrials] || !K3[nParts] || (appraisers > 1 && !K2[appraisers]) || (appraisers < 1))
      return { error: 'Need 2–3 trials × parts (1–3 appraisers)' };
    // Equipment variation (within): avg range across parts, averaged over appraisers
    let evVar = 0;
    const partMeans = Array(nParts).fill(0);
    const appMeans = [];
    const rangesByApp = [];
    for (let a = 0; a < appraisers; a++) {
      const partRanges = [];
      for (let p = 0; p < nParts; p++) {
        const vals = trials[a].map((t) => t[p]);
        vals.forEach((v) => (partMeans[p] += v / (appraisers * nTrials)));
        partRanges.push(Math.max(...vals) - Math.min(...vals));
      }
      evVar += Math.pow(mean(partRanges) * K1[nTrials], 2) / appraisers;
      appMeans.push(mean(trials[a].flat()));
      rangesByApp.push(partRanges);
    }
    const ev = Math.sqrt(evVar);
    // Appraiser variation
    const xDiffMax = Math.max(...appMeans) - Math.min(...appMeans);
    const av = appraisers > 1 ? Math.sqrt(Math.max(0, Math.pow(xDiffMax * K2[appraisers], 2) - ev * ev / (nParts * nTrials))) : 0;
    const grr = Math.sqrt(ev * ev + av * av);
    const partRange = Math.max(...partMeans) - Math.min(...partMeans);
    const pv = partRange * K3[nParts];
    const tv = Math.sqrt(grr * grr + pv * pv);
    // %GRR basis: vs tolerance per AIAG uses the 6σ spread (6·σ/Tol × 100);
    // vs total variation the σ basis cancels, so the ratio is unchanged.
    const base = tolerance > 0 ? tolerance : tv;
    const pct = (x) => (base > 0 ? round((tolerance > 0 ? 600 * x : 100 * x) / base, 1) : NaN);
    // verdict on the same basis as the displayed %GRR (10% / 30% AIAG thresholds)
    const grrPct = tolerance > 0 ? (600 * grr) / tolerance : (100 * grr) / tv;
    return {
      ev: round(ev), av: round(av), grr: round(grr), pv: round(pv), tv: round(tv),
      pctEV: pct(ev), pctAV: pct(av), pctGRR: pct(grr), pctPV: pct(pv),
      ndc: Math.floor((pv / grr) * 1.41),
      verdict: grrPct <= 10 ? 'Acceptable (≤ 10%)' : grrPct <= 30 ? 'Conditional (10–30%)' : 'Not acceptable (> 30%)',
      ok: grrPct <= 10,
      // chart detail: per-appraiser grand means, per-part means, within-part ranges, R̄̄ and its UCL
      appMeans: appMeans.map(round),
      partMeans: partMeans.map(round),
      rangesByApp,
      rbar: round(mean(rangesByApp.flat())),
      uclR: round(mean(rangesByApp.flat()) * (nTrials === 2 ? 3.267 : 2.574)),
    };
  }

  // exact two-tailed t critical values, df = 1..30 (same table as the website MsaSuite)
  const T_CRIT = [12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228, 2.201, 2.179, 2.16, 2.145, 2.131, 2.12, 2.11, 2.101, 2.093, 2.086, 2.08, 2.074, 2.069, 2.064, 2.06, 2.056, 2.052, 2.048, 2.045, 2.042];
  function bias(measurements, refValue, tolerance) {
    const n = measurements.length;
    const m = mean(measurements), s = sampleStd(measurements);
    const b = m - refValue;
    const se = s / Math.sqrt(n);
    const t = se > 0 ? Math.abs(b) / se : 0;
    const tc = n - 1 >= 1 ? (n - 1 <= 30 ? T_CRIT[n - 2] : 1.96) : NaN;
    const sig = isFinite(tc) ? t > tc : null;
    const pctBias = tolerance > 0 ? round((100 * Math.abs(b)) / tolerance, 2) : NaN;
    return { bias: round(b), mean: round(m), std: round(s), t: round(t, 2), tc, sig, pctBias, note: sig ? 't > t-crit → significant bias' : 't ≤ t-crit → bias not significant' };
  }

  function linearity(points, tolerance) {
    // points: [{ref, meas}] — AIAG linearity: regress BIAS (meas − ref) on reference;
    // slope ≈ 0 expected. Linearity index = |slope| × process span.
    const n = points.length;
    if (n < 4) return { error: 'Need ≥ 4 reference levels' };
    const xs = points.map((p) => p.ref);
    const ys = points.map((p) => (p.meas != null ? p.meas - p.ref : p.bias));
    const mx = mean(xs), my = mean(ys);
    let num = 0, den = 0;
    for (let i = 0; i < n; i++) { num += (xs[i] - mx) * (ys[i] - my); den += (xs[i] - mx) ** 2; }
    const slope = den ? num / den : 0, intercept = my - slope * mx;
    let r2num = 0, r2den = 0;
    ys.forEach((y, i) => (r2num += (slope * xs[i] + intercept - my) ** 2));
    ys.forEach((y) => (r2den += (y - my) ** 2));
    const r2 = r2den ? r2num / r2den : 0;
    const refs = [...new Set(xs)];
    const perRef = refs.map((rf) => {
      const ms = ys.filter((_, i) => xs[i] === rf);
      return { ref: rf, bias: round(mean(ms), 4) };
    });
    const span = Math.max(...xs) - Math.min(...xs);
    const linIdx = Math.abs(slope) * span;
    const pctLin = tolerance > 0 ? round((100 * linIdx) / tolerance, 1) : NaN;
    return {
      slope: round(slope, 4), intercept: round(intercept, 4), r2: round(r2, 3),
      span: round(span, 4), linearity: round(linIdx, 4), pctLin, perRef,
      verdict: !isFinite(pctLin) ? null : pctLin <= 5 ? 'ok (≤ 5% of tol)' : pctLin <= 10 ? 'Conditional (5–10%)' : 'Not acceptable (> 10%)',
      note: 'Good linearity: slope ≈ 0, R² high, linearity ≤ 5% of tolerance',
    };
  }

  // Stability: one line per time point = repeated measurements at that time.
  // X̄-R limits across time-group means (same engine as the website: computeVariable('xr')).
  function stability(groups) {
    if (groups.length < 2) return { error: 'Need ≥ 2 time points' };
    const sizes = groups.map((g) => g.measurements.length);
    const n0 = sizes[0];
    const rows = groups.map((g, i) => ({ label: g.label || 't' + (i + 1), mean: round(mean(g.measurements)), bias: round(mean(g.measurements) - g.ref) }));
    if (n0 >= 2 && n0 <= 10 && sizes.every((x) => x === n0)) {
      const means = groups.map((g) => mean(g.measurements));
      const ranges = groups.map((g) => Math.max(...g.measurements) - Math.min(...g.measurements));
      const cl = mean(means), rbar = mean(ranges);
      const ucl = cl + A2[n0] * rbar, lcl = cl - A2[n0] * rbar;
      const viol = nelson14(means, cl, ucl, lcl);
      return {
        rows, means: means.map(round), cl: round(cl), ucl: round(ucl), lcl: round(lcl),
        rbar: round(rbar), violations: viol, note: `X̄ chart on time-group means (n=${n0} per point) — points beyond ±A₂R̄ limits or Western-Electric signals indicate drift`,
      };
    }
    return { rows, note: 'Unequal group sizes — no control limits; compare each time point against the reference bias' };
  }

  function attributeAgreement(rows) {
    // rows: [{app, actual, ref}] — percent agreement + Cohen's Kappa per appraiser
    // (Kappa corrects for chance agreement — same engine as the website MsaSuite attr tab)
    const byApp = {};
    let hit = 0, tot = 0;
    rows.forEach((r) => {
      byApp[r.app] = byApp[r.app] || { hit: 0, tot: 0 };
      byApp[r.app].tot++; tot++;
      if (String(r.actual).trim() === String(r.ref).trim()) { byApp[r.app].hit++; hit++; }
    });
    const apps = [...new Set(rows.map((r) => r.app))];
    const perApp = apps.map((app) => {
      const arows = rows.filter((r) => r.app === app);
      const p0 = arows.filter((r) => String(r.actual).trim() === String(r.ref).trim()).length / arows.length;
      const cats = [...new Set(arows.flatMap((r) => [String(r.actual).trim(), String(r.ref).trim()]))];
      let pe = 0;
      cats.forEach((c) => {
        const pa = arows.filter((r) => String(r.actual).trim() === c).length / arows.length;
        const pr = arows.filter((r) => String(r.ref).trim() === c).length / arows.length;
        pe += pa * pr;
      });
      const kappa = pe < 1 ? (p0 - pe) / (1 - pe) : NaN;
      return { app, pct: Math.round(p0 * 100), kappa: round(kappa, 2), n: arows.length };
    });
    const worstK = Math.min(...perApp.map((a) => a.kappa).filter((k) => isFinite(k)));
    return {
      overall: Math.round((100 * hit) / tot), perApp,
      worstKappa: isFinite(worstK) ? worstK : null,
      verdict: worstK >= 0.75 ? 'ok' : worstK >= 0.4 ? 'conditional' : 'not acceptable',
      note: 'Agreement ≥ 90% and Kappa ≥ 0.75 = acceptable (AIAG MSA)',
    };
  }

  // ---------- Cg / Cgk (VDA 5) ----------
  function cgCgk(measurements, refValue, tolerance) {
    const s = sampleStd(measurements), m = mean(measurements);
    if (s === 0) return { error: 'Zero variation — check measurement resolution' };
    const cg = (0.2 * tolerance) / (6 * s);
    const cgk = (0.1 * tolerance - Math.abs(m - refValue)) / (3 * s);
    return {
      cg: round(cg), cgk: round(cgk), mean: round(m), std: round(s),
      verdict: cg >= 1.33 && cgk >= 1.33 ? 'Accept (≥ 1.33)' : 'Not acceptable (< 1.33)',
      ok: cg >= 1.33 && cgk >= 1.33,
    };
  }

  const api = { mean, sampleStd, round, capability, spcLimits, SPC_CHARTS, gageRR, bias, linearity, stability, attributeAgreement, cgCgk, normCdf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LKPCalc = api;
})(typeof self !== 'undefined' ? self : this);
