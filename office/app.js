/* LabKitPro Excel add-in - v1.3 (Cpk + SPC + MSA + Cg/Cgk + Transition + import + charts + download).
   Shared offline kernel: calc.js -> window.LKPCalc.{capability, spcLimits, gageRR, cgCgk, ...}
   Chart engine: charts.js -> window.LKPCharts.* (canvas -> PNG dataURL, injected as <img>).
   Data flow: worksheet selection (matrix) -> parse -> kernel -> render (+ chart) -> download report.
   v1.2.2: insert-report overwrite guard (confirm) + empty-selection error.
   v1.3:    browser-parity charts on every analysis + self-contained HTML report download. */
(function () {
  'use strict';

  Office.onReady(function () {
    var $ = function (id) { return document.getElementById(id); };
    var optNum = function (s) { s = String(s).trim(); return s === '' ? null : (isFinite(+s) ? +s : null); };

    // ---------- chart + report state (for on-screen injection + download) ----------
    var reportCharts = { cpk: [], spc: [], msa: [], cg: [] };
    var cpkVals = null;       // raw measurement array (for capability charts)
    var msaShaped = null;     // [appraiser][trial][part] (for Gage R&R charts)
    var cgVals = null, cgRef = null;

    // ---------- generic error/flash helpers (per-panel id) ----------
    function showError(elId, m, ok) {
      var e = $(elId);
      e.textContent = m; e.classList.remove('hidden');
      if (ok) e.classList.add('ok'); else e.classList.remove('ok');
    }
    function hideError(elId) { $(elId).classList.add('hidden'); }
    function flash(elId, msg) {
      var e = $(elId); e.textContent = msg; e.classList.remove('hidden'); e.classList.add('ok');
      setTimeout(function () { e.classList.add('hidden'); }, 2500);
    }

    // ---------- tab switching + active panel tracking ----------
    var currentPanel = 'cpk';
    Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (t) {
      t.addEventListener('click', function () {
        var name = t.getAttribute('data-tab');
        currentPanel = name;
        Array.prototype.forEach.call(document.querySelectorAll('.tab'), function (x) { x.classList.remove('active'); });
        Array.prototype.forEach.call(document.querySelectorAll('.panel'), function (x) { x.classList.remove('active'); });
        t.classList.add('active');
        document.querySelector('.panel[data-panel="' + name + '"]').classList.add('active');
        $('importBar').classList.toggle('hidden', name === 'tr');   // import only applies to data tools
        if (name === 'tr') renderTransition();
      });
    });

    // ============================================================
    //  Cpk (preserved from v1.0)
    // ============================================================
    var cpkLast = null;
    $('analyze').addEventListener('click', analyzeCpk);
    $('loadSample').addEventListener('click', function (e) { e.preventDefault(); fillCpkSample(); });
    $('copyReport').addEventListener('click', function () { copyReport(cpkLast, cpkReportText, 'error'); });
    $('insertReport').addEventListener('click', function () { insertReport(cpkLast, cpkReportText, 'error'); });

    function fillCpkSample() {
      $('lsl').value = '9.80'; $('usl').value = '10.20'; $('target').value = '10.00';
      var data = [10.02, 9.97, 10.11, 10.05, 9.94, 10.08, 10.01, 9.99, 10.06, 10.03,
                  9.96, 10.09, 10.00, 10.04, 9.98, 10.07, 10.02, 9.95, 10.10, 10.01,
                  10.05, 9.99, 10.03, 10.06, 9.97];
      var rows = data.map(function (v) { return [[v]]; });
      Office.context.document.setSelectedDataAsync(rows, { coercionType: Office.CoercionType.Matrix }, function (ar) {
        if (ar.status === Office.AsyncResultStatus.Failed) showError('error', 'Could not write sample data: ' + ar.error.message);
      });
    }

    function analyzeCpk() {
      hideError('error');
      var lsl = optNum($('lsl').value), usl = optNum($('usl').value);
      if (lsl === null && usl === null) return showError('error', 'Enter at least one spec limit (LSL or USL).');
      getMatrix(runCpk, 'error');
    }

    function runCpk(matrix) {
      hideError('error');
      var lsl = optNum($('lsl').value), usl = optNum($('usl').value), target = optNum($('target').value);
      if (lsl === null && usl === null) return showError('error', 'Enter at least one spec limit (LSL or USL).');
      if (!matrix || !matrix.length) return showError('error', 'Select a range of measurement values first.');
      var cols = [], all = [];
      for (var c = 0; c < matrix[0].length; c++) cols.push([]);
      for (var r = 0; r < matrix.length; r++) for (var c2 = 0; c2 < matrix[r].length; c2++) {
        var v = typeof matrix[r][c2] === 'number' ? matrix[r][c2] : parseFloat(String(matrix[r][c2]).trim());
        if (isFinite(v)) { cols[c2].push(v); all.push(v); }
      }
      var numericCols = cols.filter(function (col) { return col.length >= 2; });
      var values, subgroups = null;
      if (numericCols.length >= 2) { values = [].concat.apply([], numericCols); subgroups = numericCols; }
      else values = all;
      if (values.length < 2) return showError('error', 'No numeric values found. Select the cells that contain your measurements.');
      var res = LKPCalc.capability(values, lsl, usl, target, subgroups);
      if (res.error) return showError('error', res.error);
      cpkVals = values; cpkLast = res; renderCpk(res);
    }

    function renderCpk(res) {
      var eff = (res.cpk != null ? res.cpk : res.ppk);
      $('verdict').textContent = res.verdict + '  (Cpk/Ppk = ' + eff + ')';
      $('verdict').className = 'verdict ' + (res.ok ? 'ok' : eff >= 1.0 ? 'marginal' : 'bad');
      var rows = [['N (values)', res.n], ['Mean', res.mean], ['Std dev (overall)', res.std]];
      if (res.sigW != null) rows.push(['Sigma within (R-bar/d2)', res.sigW]);
      var spec = [['Pp', res.pp], ['Ppk', res.ppk]];
      if (res.cp != null) spec = [['Cp', res.cp], ['Cpk', res.cpk], ['CPU', res.cpu], ['CPL', res.cpl]].concat(spec);
      if (res.cpm != null) spec.push(['Cpm', res.cpm]);
      rows = rows.concat(spec, [['Expected PPM', res.ppm]]);
      $('metrics').innerHTML = rows.map(function (r) {
        var dim = (r[1] === null || r[1] === undefined) ? ' class="dim"' : '';
        return '<tr' + dim + '><td>' + r[0] + '</td><td>' + (r[1] == null ? 'n/a' : r[1]) + '</td></tr>';
      }).join('');
      // browser-parity charts
      var lsl = optNum($('lsl').value), usl = optNum($('usl').value);
      var charts = [
        { title: 'Capability histogram — fitted normal curve with LSL / USL / mean', data: LKPCharts.capabilityChart(cpkVals, lsl, usl, res.mean, res.std) },
        { title: 'Run chart — individual values in data order vs spec limits', data: LKPCharts.capabilityRun(cpkVals, lsl, usl, res.mean) },
      ];
      renderCharts('cpkCharts', 'cpk', charts);
      $('result').classList.remove('hidden');
    }

    function cpkReportText() {
      var r = cpkLast; if (!r) return null;
      var L = [];
      L.push('LabKitPro - Process capability report');
      L.push('Date: ' + new Date().toISOString().slice(0, 10));
      L.push('LSL: ' + (optNum($('lsl').value) == null ? '-' : $('lsl').value) +
             '  USL: ' + (optNum($('usl').value) == null ? '-' : $('usl').value) +
             '  Target: ' + (optNum($('target').value) == null ? '-' : $('target').value));
      L.push('');
      L.push('N = ' + r.n + '   Mean = ' + r.mean + '   Std (overall) = ' + r.std);
      if (r.cp != null) L.push('Cp = ' + r.cp + '   Cpk = ' + r.cpk + '   CPU = ' + r.cpu + '   CPL = ' + r.cpl);
      if (r.sigW != null) L.push('Sigma within = ' + r.sigW);
      L.push('Pp = ' + (r.pp == null ? 'n/a' : r.pp) + '   Ppk = ' + r.ppk);
      if (r.cpm != null) L.push('Cpm = ' + r.cpm);
      L.push('Expected PPM = ' + r.ppm);
      L.push('Verdict: ' + r.verdict + ' (AIAG basis, capable >= 1.33)');
      L.push('');
      L.push('Full suite: https://labkitpro.com/tools/');
      return L.join('\n');
    }

    // ============================================================
    //  SPC
    // ============================================================
    var spcLast = null;
    var SPC_VAR = { 'xbar-r': 1, 'xbar-s': 1, 'xbar-r-median': 1, 'i-mr': 0 };
    var SPC_ATTR = { 'p': 1, 'np': 1, 'c': 1, 'u': 1 };

    function updateSpcHint() {
      var type = $('spcType').value;
      var h = $('spcHint');
      if (SPC_ATTR[type]) {
        h.innerHTML = '<b>Attribute data</b> &mdash; select a 2-column range: column 1 = sample size <b>n</b>, column 2 = defect count <b>d</b> (one row per subgroup).';
      } else if (type === 'i-mr') {
        h.innerHTML = '<b>Individuals</b> &mdash; select one column of individual readings (time-ordered). I-MR uses moving ranges.';
      } else {
        h.innerHTML = '<b>Subgroups</b> &mdash; select the data with <b>each column = one subgroup</b> (constant size 2&ndash;10). Need at least 2 subgroups.';
      }
    }
    $('spcType').addEventListener('change', updateSpcHint);
    updateSpcHint();
    $('spcAnalyze').addEventListener('click', analyzeSpc);
    $('spcSample').addEventListener('click', function (e) { e.preventDefault(); fillSpcSample(); });
    $('spcCopy').addEventListener('click', function () { copyReport(spcLast, spcReportText, 'spcError'); });
    $('spcInsert').addEventListener('click', function () { insertReport(spcLast, spcReportText, 'spcError'); });

    function fillSpcSample() {
      var type = $('spcType').value;
      var rows;
      if (SPC_ATTR[type]) {
        rows = [[50, 2], [50, 1], [50, 3], [50, 0], [50, 2], [50, 4], [50, 1], [50, 2], [50, 0], [50, 3], [50, 2], [50, 1]];
      } else if (type === 'i-mr') {
        rows = [[10.1], [10.3], [9.9], [10.2], [10.0], [10.4], [9.8], [10.1], [10.3], [9.7], [10.2], [10.0], [10.5], [9.9], [10.1]];
      } else {
        rows = [[10.02, 9.97, 10.11], [10.05, 9.94, 10.08], [10.01, 9.99, 10.06], [10.03, 9.96, 10.09],
                [10.00, 10.04, 10.07], [10.02, 9.95, 10.10], [10.01, 9.99, 10.03], [10.06, 9.97, 10.05]];
      }
      Office.context.document.setSelectedDataAsync(rows, { coercionType: Office.CoercionType.Matrix }, function (ar) {
        if (ar.status === Office.AsyncResultStatus.Failed) showError('spcError', 'Could not write sample data: ' + ar.error.message);
      });
    }

    function analyzeSpc() {
      hideError('spcError');
      getMatrix(runSpc, 'spcError');
    }

    function runSpc(matrix) {
      hideError('spcError');
      var type = $('spcType').value;
      if (!matrix || !matrix.length) return showError('spcError', 'Select the data range first.');
      var groups;
      if (SPC_ATTR[type]) {
        groups = matrix.map(function (row) {
          var n = optNum(row[0]), d = optNum(row[1]);
          if (n === null || d === null) return null;
          return { n: n, d: d };
        }).filter(Boolean);
        if (groups.length < 4) return showError('spcError', 'Need at least 4 subgroups (rows) of n/d pairs.');
      } else {
        var cols = [];
        for (var c = 0; c < matrix[0].length; c++) cols.push([]);
        for (var r = 0; r < matrix.length; r++) for (var c2 = 0; c2 < matrix[r].length; c2++) {
          var v = typeof matrix[r][c2] === 'number' ? matrix[r][c2] : parseFloat(String(matrix[r][c2]).trim());
          if (isFinite(v)) cols[c2].push(v);
        }
        if (type === 'i-mr') {
          groups = [].concat.apply([], cols).map(function (x) { return [x]; });
        } else {
          groups = cols.filter(function (col) { return col.length >= 2; });
          if (groups.length < 2) return showError('spcError', 'Select at least 2 subgroup columns (each column = one subgroup, size 2-10).');
        }
      }
      var res = LKPCalc.spcLimits(type, groups);
      if (res.error) return showError('spcError', res.error);
      spcLast = res; renderSpc(res, type);
    }

    function renderSpc(res, type) {
      var n = res.violations ? res.violations.length : 0;
      var ok = n === 0;
      $('spcVerdict').textContent = (ok ? 'In control' : ('Out of control - ' + n + ' signal' + (n > 1 ? 's' : ''))) +
        '  (' + type.toUpperCase() + (res.sizeNote ? ', ' + res.sizeNote : '') + ')';
      $('spcVerdict').className = 'verdict ' + (ok ? 'ok' : 'bad');
      var rows = [
        [type.toUpperCase() + ' center (CL)', res.cl],
        [type.toUpperCase() + ' UCL', Array.isArray(res.ucl) ? 'variable' : res.ucl],
        [type.toUpperCase() + ' LCL', Array.isArray(res.lcl) ? 'variable' : res.lcl],
      ];
      if (res.extra) {
        rows.push([res.extra.chart + ' center (CL)', res.extra.cl]);
        rows.push([res.extra.chart + ' UCL', res.extra.ucl]);
        rows.push([res.extra.chart + ' LCL', res.extra.lcl]);
      }
      rows.push(['Subgroups (k)', res.k]);
      $('spcMetrics').innerHTML = rows.map(function (r) {
        return '<tr><td>' + r[0] + '</td><td>' + (r[1] == null ? 'n/a' : r[1]) + '</td></tr>';
      }).join('');
      var v = res.violations || [];
      $('spcViol').innerHTML = v.length
        ? '<b>Signals:</b><ul>' + v.map(function (x) { return '<li>Pt ' + (x.i + 1) + ' &mdash; rule ' + x.rule + ': ' + x.txt + '</li>'; }).join('') + '</ul>'
        : '<span class="ok-note">No Western Electric / Nelson signals detected.</span>';
      // browser-parity charts: main chart + variation (R/S/MR) chart when present
      var charts = [{ title: type.toUpperCase() + ' control chart — values vs CL/UCL/LCL (red = rule signal)', data: LKPCharts.spcChart(res.points, res.cl, res.ucl, res.lcl, res.violations, 360) }];
      if (res.extra && res.extra.points.length > 1) {
        charts.push({ title: res.extra.chart + ' chart (variation) — values vs CL/UCL/LCL', data: LKPCharts.spcChart(res.extra.points, res.extra.cl, res.extra.ucl, res.extra.lcl, [], 300) });
      }
      renderCharts('spcCharts', 'spc', charts);
      $('spcResult').classList.remove('hidden');
    }

    function spcReportText() {
      var r = spcLast; if (!r) return null;
      var L = [];
      L.push('LabKitPro - SPC report (' + ($('spcType').value.toUpperCase()) + ')');
      L.push('Date: ' + new Date().toISOString().slice(0, 10));
      L.push('Subgroups (k) = ' + r.k + (r.sizeNote ? '  (' + r.sizeNote + ')' : ''));
      L.push('CL = ' + r.cl + '   UCL = ' + (Array.isArray(r.ucl) ? 'variable' : r.ucl) + '   LCL = ' + (Array.isArray(r.lcl) ? 'variable' : r.lcl));
      if (r.extra) L.push(r.extra.chart + ': CL = ' + r.extra.cl + '  UCL = ' + r.extra.ucl + '  LCL = ' + r.extra.lcl);
      var v = r.violations || [];
      L.push('Signals: ' + v.length);
      v.forEach(function (x) { L.push('  Pt ' + (x.i + 1) + ' rule ' + x.rule + ': ' + x.txt); });
      L.push('Verdict: ' + (v.length ? 'Out of control' : 'In control'));
      L.push('');
      L.push('Full suite: https://labkitpro.com/tools/');
      return L.join('\n');
    }

    // ============================================================
    //  MSA (Gage R&R)
    // ============================================================
    var msaLast = null;
    $('msaAnalyze').addEventListener('click', analyzeMsa);
    $('msaSample').addEventListener('click', function (e) { e.preventDefault(); fillMsaSample(); });
    $('msaCopy').addEventListener('click', function () { copyReport(msaLast, msaReportText, 'msaError'); });
    $('msaInsert').addEventListener('click', function () { insertReport(msaLast, msaReportText, 'msaError'); });

    function fillMsaSample() {
      // 2 appraisers x 3 trials x 5 parts
      var rows = [
        [0.10, 0.12, 0.09, 0.11, 0.10], [0.11, 0.13, 0.10, 0.12, 0.11], [0.10, 0.12, 0.09, 0.11, 0.10],
        [0.11, 0.12, 0.10, 0.12, 0.11], [0.12, 0.13, 0.11, 0.13, 0.12], [0.11, 0.12, 0.10, 0.12, 0.11]
      ];
      Office.context.document.setSelectedDataAsync(rows, { coercionType: Office.CoercionType.Matrix }, function (ar) {
        if (ar.status === Office.AsyncResultStatus.Failed) showError('msaError', 'Could not write sample data: ' + ar.error.message);
      });
    }

    function analyzeMsa() {
      hideError('msaError');
      var app = optNum($('msaApp').value), trials = optNum($('msaTrials').value), parts = optNum($('msaParts').value), tol = optNum($('msaTol').value);
      if (app === null || trials === null || parts === null) return showError('msaError', 'Appraisers, Trials and Parts must be 1-3 / 2-3 / 3-10.');
      if (app < 1 || app > 3 || trials < 2 || trials > 3 || parts < 3 || parts > 10) return showError('msaError', 'Appraisers 1-3, Trials 2-3, Parts 3-10.');
      getMatrix(runMsa, 'msaError');
    }

    function runMsa(matrix) {
      hideError('msaError');
      var app = optNum($('msaApp').value), trials = optNum($('msaTrials').value), parts = optNum($('msaParts').value), tol = optNum($('msaTol').value);
      if (!matrix || !matrix.length) return showError('msaError', 'Select the measurement table first.');
      var needRows = app * trials;
      if (matrix.length !== needRows) return showError('msaError', 'Need exactly ' + needRows + ' rows (appraisers x trials); found ' + matrix.length + '.');
      if (matrix[0].length !== parts) return showError('msaError', 'Need exactly ' + parts + ' columns (parts); found ' + matrix[0].length + '.');
      var trials3 = [];
      for (var a = 0; a < app; a++) {
        var appTrials = [];
        for (var t = 0; t < trials; t++) {
          var row = matrix[a * trials + t];
          appTrials.push(row.map(function (x) { return typeof x === 'number' ? x : parseFloat(String(x).trim()); }));
        }
        trials3.push(appTrials);
      }
      var res = LKPCalc.gageRR(trials3, tol == null ? 0 : tol);
      if (res.error) return showError('msaError', res.error);
      msaShaped = trials3; msaLast = res; renderMsa(res);
    }

    function renderMsa(res) {
      var ok = res.ok;
      $('msaVerdict').textContent = res.verdict + '  (%GRR = ' + res.pctGRR + ')';
      $('msaVerdict').className = 'verdict ' + (ok ? 'ok' : res.pctGRR <= 30 ? 'marginal' : 'bad');
      var rows = [
        ['Repeatability (EV)', res.ev], ['Reproducibility (AV)', res.av], ['Gage R&R (GRR)', res.grr],
        ['Part Variation (PV)', res.pv], ['Total Variation (TV)', res.tv],
        ['%EV', res.pctEV], ['%AV', res.pctAV], ['%GRR', res.pctGRR], ['%PV', res.pctPV],
        ['NDC', res.ndc]
      ];
      $('msaMetrics').innerHTML = rows.map(function (r) {
        return '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td></tr>';
      }).join('');
      // browser-parity charts (mirrors the Chrome pop-up): appraiser view when >1, else trial view
      var nApp = msaShaped.length;
      var appData = msaShaped.map(function (tr, i) {
        return { name: String.fromCharCode(65 + i), trials: tr, ranges: res.rangesByApp[i] };
      });
      var charts;
      if (nApp > 1) {
        charts = [
          { title: 'All measurements by part — colored by appraiser, line = appraiser mean (repeatability + reproducibility)', data: LKPCharts.grrAppChart(appData) },
          { title: 'Range within part by appraiser vs R̄̄ and UCL = D4·R̄̄ (repeatability check)', data: LKPCharts.grrRangeMulti(appData, res.rbar, res.uclR) },
          { title: 'Variance components — %EV / %AV / %GRR / %PV vs 10% / 30% thresholds', data: LKPCharts.varianceChart(res) },
        ];
      } else {
        var byPart = appData[0].trials.map(function (_, p) { return appData[0].trials.map(function (t) { return t[p]; }); });
        charts = [
          { title: 'Measurements by part — one line per trial (repeatability view)', data: LKPCharts.grrChart(byPart) },
          { title: 'Range within part vs R̄ and UCL = D4·R̄ (trial-to-trial spread check)', data: LKPCharts.grrRangeChart(byPart) },
          { title: 'Variance components — %EV / %GRR / %PV vs 10% / 30% thresholds', data: LKPCharts.varianceChart(res) },
        ];
      }
      renderCharts('msaCharts', 'msa', charts);
      $('msaResult').classList.remove('hidden');
    }

    function msaReportText() {
      var r = msaLast; if (!r) return null;
      var L = [];
      L.push('LabKitPro - Gage R&R (ANOVA-less, AIAG) report');
      L.push('Date: ' + new Date().toISOString().slice(0, 10));
      L.push('EV = ' + r.ev + '   AV = ' + r.av + '   GRR = ' + r.grr);
      L.push('PV = ' + r.pv + '   TV = ' + r.tv);
      L.push('%EV = ' + r.pctEV + '   %AV = ' + r.pctAV + '   %GRR = ' + r.pctGRR + '   %PV = ' + r.pctPV);
      L.push('NDC = ' + r.ndc);
      L.push('Verdict: ' + r.verdict);
      L.push('');
      L.push('Full suite: https://labkitpro.com/tools/');
      return L.join('\n');
    }

    // ============================================================
    //  Cg / Cgk
    // ============================================================
    var cgLast = null;
    $('cgAnalyze').addEventListener('click', analyzeCg);
    $('cgSample').addEventListener('click', function (e) { e.preventDefault(); fillCgSample(); });
    $('cgCopy').addEventListener('click', function () { copyReport(cgLast, cgReportText, 'cgError'); });
    $('cgInsert').addEventListener('click', function () { insertReport(cgLast, cgReportText, 'cgError'); });

    function fillCgSample() {
      $('cgRef').value = '10.000'; $('cgTol').value = '0.200';
      var data = [10.004, 9.998, 10.006, 10.002, 9.995, 10.008, 10.001, 9.997, 10.003, 10.000,
                  10.005, 9.999, 10.002, 10.006, 9.994, 10.001, 10.003, 9.996, 10.007, 10.000];
      Office.context.document.setSelectedDataAsync(data.map(function (v) { return [[v]]; }), { coercionType: Office.CoercionType.Matrix }, function (ar) {
        if (ar.status === Office.AsyncResultStatus.Failed) showError('cgError', 'Could not write sample data: ' + ar.error.message);
      });
    }

    function analyzeCg() {
      hideError('cgError');
      var ref = optNum($('cgRef').value), tol = optNum($('cgTol').value);
      if (ref === null) return showError('cgError', 'Enter the reference (true) value.');
      if (tol === null || tol <= 0) return showError('cgError', 'Enter a positive tolerance band (USL-LSL).');
      getMatrix(runCg, 'cgError');
    }

    function runCg(matrix) {
      hideError('cgError');
      var ref = optNum($('cgRef').value), tol = optNum($('cgTol').value);
      if (ref === null) return showError('cgError', 'Enter the reference (true) value.');
      if (tol === null || tol <= 0) return showError('cgError', 'Enter a positive tolerance band (USL-LSL).');
      if (!matrix || !matrix.length) return showError('cgError', 'Select the measurement column first.');
      var vals = [];
      for (var r = 0; r < matrix.length; r++) for (var c = 0; c < matrix[r].length; c++) {
        var v = typeof matrix[r][c] === 'number' ? matrix[r][c] : parseFloat(String(matrix[r][c]).trim());
        if (isFinite(v)) vals.push(v);
      }
      if (vals.length < 2) return showError('cgError', 'No numeric values found in the selection.');
      var res = LKPCalc.cgCgk(vals, ref, tol);
      if (res.error) return showError('cgError', res.error);
      cgVals = vals; cgRef = ref; cgLast = res; renderCg(res);
    }

    function renderCg(res) {
      $('cgVerdict').textContent = res.verdict;
      $('cgVerdict').className = 'verdict ' + (res.ok ? 'ok' : 'bad');
      var rows = [['N', res.n !== undefined ? res.n : ''], ['Mean', res.mean], ['Std dev', res.std], ['Cg', res.cg], ['Cgk', res.cgk]];
      $('cgMetrics').innerHTML = rows.map(function (r) {
        return '<tr><td>' + r[0] + '</td><td>' + (r[1] === '' ? '-' : r[1]) + '</td></tr>';
      }).join('');
      // browser-parity charts: run chart + deviation-from-reference histogram
      var bias = res.mean - cgRef;
      var charts = [
        { title: 'Run chart — repeated measurements vs reference (orange band = x̄ ± 3s)', data: LKPCharts.runChart(cgVals, cgRef, res.mean, res.std) },
        { title: 'Deviation from reference (measured − ref) — bias marker vs zero', data: LKPCharts.devHist(cgVals, cgRef, bias) },
      ];
      renderCharts('cgCharts', 'cg', charts);
      $('cgResult').classList.remove('hidden');
    }

    function cgReportText() {
      var r = cgLast; if (!r) return null;
      var L = [];
      L.push('LabKitPro - Measurement capability (Cg / Cgk, VDA 5) report');
      L.push('Date: ' + new Date().toISOString().slice(0, 10));
      L.push('Reference = ' + $('cgRef').value + '   Tolerance = ' + $('cgTol').value);
      L.push('Mean = ' + r.mean + '   Std dev = ' + r.std);
      L.push('Cg = ' + r.cg + '   Cgk = ' + r.cgk);
      L.push('Verdict: ' + r.verdict);
      L.push('');
      L.push('Full suite: https://labkitpro.com/tools/');
      return L.join('\n');
    }

    // ============================================================
    //  Charts (browser-parity) + self-contained HTML report download
    // ============================================================
    // inject a list of {title, data(PNG dataURL)} into a result container, and remember them for the report
    function renderCharts(containerId, tool, charts) {
      reportCharts[tool] = charts || [];
      var box = $(containerId);
      if (!box) return;
      if (!reportCharts[tool].length) { box.innerHTML = ''; return; }
      box.innerHTML = reportCharts[tool].map(function (c) {
        return '<div class="chart-block"><img class="chart" src="' + c.data + '" alt="' + (c.title || '') + '" />' +
          (c.title ? '<div class="chart-cap">' + c.title + '</div>' : '') + '</div>';
      }).join('');
    }

    var REPORT_META = {
      cpk: { title: 'Process Capability (Cpk / Ppk)', verdictId: 'verdict', tableId: 'metrics', extraId: null,
        method: 'Capability is assessed on the AIAG dual basis: within-subgroup Cpk (R̄/d₂) when explicit subgroup structure is present, otherwise long-term Ppk. Expected PPM uses the normal integral of the tails. Thresholds: capable ≥ 1.33, marginal 1.00–1.33, not capable < 1.00.' },
      spc: { title: 'Statistical Process Control', verdictId: 'spcVerdict', tableId: 'spcMetrics', extraId: 'spcViol',
        method: 'Control limits at ±3σ (Shewhart constants for subgroup sizes 2–10) with Western-Electric / Nelson signals: ① beyond limits ② 2-of-3 beyond 2σ ③ 4-of-5 beyond 1σ ④ 8 consecutive one side ⑤ 6 trending — per AIAG SPC manual.' },
      msa: { title: 'Gage R&R (Measurement Systems Analysis)', verdictId: 'msaVerdict', tableId: 'msaMetrics', extraId: null,
        method: 'Variable Gage R&R by the X̄-R method (AIAG): EV (repeatability), AV (reproducibility), GRR, PV and TV; %GRR vs tolerance or total variation. Thresholds: %GRR ≤ 10% acceptable, 10–30% conditional, > 30% not acceptable; NDC ≥ 5 required.' },
      cg: { title: 'Measurement Capability (Cg / Cgk, VDA 5)', verdictId: 'cgVerdict', tableId: 'cgMetrics', extraId: null,
        method: 'Cg uses the 20%-tolerance rule (0.2·Tol / 6σ) and Cgk the 10%-tolerance deviation rule (0.1·Tol − |mean−ref|) / 3σ. VDA 5 thresholds: Cg ≥ 1.33 and Cgk ≥ 1.33 = accept.' },
    };

    function buildReportHtml(tool) {
      var m = REPORT_META[tool];
      if (!m) return null;
      var vEl = $(m.verdictId);
      var verdictText = vEl ? vEl.textContent : '';
      var vClass = vEl ? (vEl.className || '').replace('verdict', '').trim() : '';
      var tableHtml = $(m.tableId) ? $(m.tableId).innerHTML : '';
      var extraHtml = '';
      if (m.extraId && $(m.extraId) && $(m.extraId).innerHTML.trim()) {
        extraHtml = '<h2>Signals</h2>' + $(m.extraId).innerHTML;
      }
      var title = m.title;
      if (tool === 'spc') title = 'Statistical Process Control (' + ($('spcType').value || '').toUpperCase() + ')';
      var charts = reportCharts[tool] || [];
      var chartHtml = '';
      if (charts.length) {
        chartHtml = '<h2>Charts</h2>' + charts.map(function (c) {
          return '<figure class="fig"><img src="' + c.data + '" alt="' + (c.title || '') + '" /><figcaption>' + (c.title || '') + '</figcaption></figure>';
        }).join('');
      }
      var date = new Date().toISOString().slice(0, 10);
      var safe = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
      return '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"/>' +
        '<meta name="viewport" content="width=device-width,initial-scale=1"/>' +
        '<title>LabKitPro — ' + safe(title) + ' report</title><style>' +
        'body{font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#1e293b;max-width:820px;margin:24px auto;padding:0 18px;line-height:1.55;}' +
        'h1{font-size:22px;color:#0f2f5c;margin:0 0 4px;}h2{font-size:15px;color:#0f2f5c;margin:22px 0 8px;border-bottom:1px solid #e2e8f0;padding-bottom:4px;}' +
        '.meta{color:#64748b;font-size:12px;margin:0 0 14px;}' +
        '.verdict{padding:11px 13px;border-radius:8px;font-weight:700;font-size:14px;margin:0 0 14px;}' +
        '.verdict.ok{background:#dcfce7;color:#166534;}.verdict.marginal{background:#fef3c7;color:#92400e;}.verdict.bad{background:#fee2e2;color:#991b1b;}' +
        'table.metrics{width:100%;border-collapse:collapse;margin:0 0 12px;}table.metrics td{padding:7px 6px;border-bottom:1px solid #e8edf3;}' +
        'table.metrics td:first-child{color:#64748b;}table.metrics td:last-child{text-align:right;font-variant-numeric:tabular-nums;font-weight:700;}' +
        '.fig{margin:0 0 16px;text-align:center;}figure.img{} .fig img{max-width:100%;height:auto;border:1px solid #e8edf3;border-radius:8px;background:#fff;}' +
        '.fig figcaption{font-size:12px;color:#64748b;margin-top:5px;}' +
        '.method{font-size:13px;color:#334155;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px 14px;}' +
        '.foot{margin-top:26px;padding-top:12px;border-top:1px dashed #cbd5e1;font-size:12px;color:#64748b;}' +
        '.foot a{color:#0f2f5c;font-weight:600;text-decoration:none;}' +
        '</style></head><body>' +
        '<h1>LabKitPro — ' + safe(title) + '</h1>' +
        '<p class="meta">Generated ' + date + ' · offline analysis · data stayed in Excel</p>' +
        '<div class="verdict ' + vClass + '">' + safe(verdictText) + '</div>' +
        '<table class="metrics">' + tableHtml + '</table>' +
        extraHtml +
        chartHtml +
        '<h2>Methodology</h2><p class="method">' + safe(m.method) + '</p>' +
        '<div class="foot">Full suite &amp; paid exports: <a href="https://labkitpro.com/tools/?utm_source=excel-addin&utm_medium=report&utm_campaign=v1" target="_blank">labkitpro.com/tools</a> · ' +
        'Questions: <a href="mailto:support@labkitpro.com">support@labkitpro.com</a></div>' +
        '</body></html>';
    }

    function downloadReport(tool) {
      var html = buildReportHtml(tool);
      if (!html) return;
      var blob = new Blob([html], { type: 'text/html;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = 'labkitpro-' + tool + '-' + new Date().toISOString().slice(0, 10) + '.html';
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    }

    $('cpkDownload').addEventListener('click', function () { if (cpkLast) downloadReport('cpk'); });
    $('spcDownload').addEventListener('click', function () { if (spcLast) downloadReport('spc'); });
    $('msaDownload').addEventListener('click', function () { if (msaLast) downloadReport('msa'); });
    $('cgDownload').addEventListener('click', function () { if (cgLast) downloadReport('cg'); });

    // ============================================================
    //  Transition countdown (rich, mirrors the browser hub)
    // ============================================================
    var UTM = '?utm_source=excel-addin&utm_medium=addin&utm_campaign=v1';
    var TRANSITION = [
      { id: 'iso9001-2026', std: 'ISO 9001:2026', status: 'Published 2026-09-16', hard: true,
        deadline: '2029-09-30', deadlineLabel: 'Transition deadline',
        milestones: [['2027-03-31', 'Accreditation bodies ready'], ['2027-06-30', 'CBs declare transition arrangements'],
                     ['2028-03-31', 'New certificates issued to new version only'], ['2029-09-30', 'All existing certificates expire']],
        url: 'https://labkitpro.com/iso-9001-2026/' },
      { id: 'iso14001-2026', std: 'ISO 14001:2026', status: 'Published 2026-04-15', hard: true,
        deadline: '2029-04-30', deadlineLabel: 'Transition deadline',
        milestones: [['2026-10-31', 'Accreditation bodies ready'], ['2027-01-31', 'CBs declare transition arrangements'],
                     ['2027-10-31', 'New certificates issued to new version only'], ['2029-04-30', 'All existing certificates expire']],
        url: 'https://labkitpro.com/iso-14001-2026/' },
      { id: 'iso45001-2027', std: 'ISO 45001:2027', status: 'DIS 40.60 — expected 2027', hard: false,
        deadline: '2030-06-30', deadlineLabel: 'Expected deadline (TBC)',
        milestones: [['2027-06-30', 'Publication expected (H1 2027)'], ['2030-06-30', 'Transition period expected ~3 years (estimate)']],
        url: 'https://labkitpro.com/iso-45001-2027/' },
    ];

    function fmtDays(d) { return d.toLocaleString('en-US'); }

    function renderTransition() {
      var grid = $('trGrid');
      if (!grid) return;
      var now = new Date();
      grid.innerHTML = TRANSITION.map(function (t) {
        var html = '<div class="tr-card">';
        html += '<div class="tr-top"><div class="tr-std">' + t.std + '</div>' +
          '<span class="tr-status ' + (t.hard ? 'hard' : 'soft') + '">' + t.status + '</span></div>';
        // countdown
        var dl = new Date(t.deadline + 'T00:00:00');
        var days = Math.ceil((dl - now) / 86400000);
        var cls = days < 0 ? 'past' : days < 365 ? 'hot' : 'ok';
        html += '<div class="tr-count"><span class="tr-num ' + cls + '">' + (days < 0 ? 'closed' : fmtDays(days)) + '</span>' +
          (days >= 0 ? '<span class="tr-unit">days</span>' : '') + '</div>';
        html += '<div class="tr-deadline">' + t.deadlineLabel + ': <b>' + t.deadline + '</b></div>';
        // next milestone banner
        var next = null, allDone = true;
        t.milestones.forEach(function (m) {
          var md = new Date(m[0] + 'T00:00:00');
          if (md.getTime() >= now.getTime()) { if (!next) next = { date: m[0], label: m[1], days: Math.ceil((md - now) / 86400000) }; }
          else allDone = false;
        });
        var bannerCls = cls;
        if (next) {
          var inTxt = next.days <= 0 ? 'today' : (next.days === 1 ? 'in 1 day' : 'in ' + fmtDays(next.days) + ' days');
          html += '<div class="tr-banner ' + bannerCls + '"><span class="lbl">Next milestone · ' + inTxt + '</span>' +
            next.label + ' &mdash; ' + next.date + '</div>';
        } else {
          html += '<div class="tr-banner ok"><span class="lbl">Milestones</span>' +
            (allDone ? 'All listed milestones reached &mdash; see certificate expiry above.' : 'No further scheduled milestones.') + '</div>';
        }
        // timeline
        html += '<ul class="tr-timeline">';
        t.milestones.forEach(function (m) {
          var md = new Date(m[0] + 'T00:00:00');
          var state = md.getTime() < now.getTime() ? 'done' : (next && m[0] === next.date ? 'next' : '');
          var tag = state === 'done' ? ' ✓' : '';
          html += '<li class="' + state + '"><span class="d">' + m[0] + '</span> &mdash; ' + m[1] + tag + '</li>';
        });
        html += '</ul>';
        // guide button
        html += '<a class="tr-guide" href="' + t.url + UTM + '" target="_blank">Transition guide &amp; checklist &#8599;</a>';
        html += '</div>';
        return html;
      }).join('');
    }

    // ============================================================
    //  Import external data (shared bar) - paste CSV/TSV -> analyze
    // ============================================================
    var IMPORT_ROUTE = { cpk: runCpk, spc: runSpc, msa: runMsa, cg: runCg };
    function parseMatrix(text) {
      var lines = String(text).replace(/\r/g, '').split('\n');
      var out = [];
      for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (line === '') continue;
        var parts;
        if (line.indexOf('\t') >= 0) parts = line.split('\t');
        else if (line.indexOf(',') >= 0) parts = line.split(',');
        else parts = line.split(/\s+/);
        var row = parts.map(function (p) { return p.trim(); }).filter(function (p) { return p !== ''; });
        if (row.length) out.push(row);
      }
      return out.length ? out : null;
    }
    function toNumeric(matrix) {
      return matrix.map(function (row) {
        return row.map(function (c) {
          if (c === '' || c == null) return '';
          var n = Number(c);
          return isFinite(n) ? n : c;
        });
      });
    }
    function doImport() {
      var txt = $('importText').value;
      var m = parseMatrix(txt);
      if (!m) return showError('importError', 'Paste some numeric data first (one value per cell, rows on separate lines).');
      hideError('importError');
      // write to the sheet for visibility (best effort; analysis runs on the pasted data regardless)
      Office.context.document.setSelectedDataAsync(toNumeric(m), { coercionType: Office.CoercionType.Matrix }, function (ar) {
        if (ar.status === Office.AsyncResultStatus.Failed)
          showError('importError', 'Could not write to sheet: ' + ar.error.message + ' (results shown from pasted data anyway).');
      });
      var fn = IMPORT_ROUTE[currentPanel];
      if (fn) fn(m);
    }
    $('importToggle').addEventListener('click', function () { $('importPanel').classList.toggle('hidden'); });
    $('importRun').addEventListener('click', doImport);
    $('importClear').addEventListener('click', function () { $('importText').value = ''; hideError('importError'); });

    // ============================================================
    //  shared: read selection, copy / insert report
    // ============================================================
    function getMatrix(cb, errId) {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
        { valueFormat: Office.ValueFormat.Unformatted }, function (ar) {
          if (ar.status === Office.AsyncResultStatus.Failed) {
            if (errId) showError(errId, 'Select a range of cells first, then click Analyze.');
            return; // no selection / read failed -> surface a friendly message
          }
          cb(ar.value);
        });
    }

    function copyReport(last, textFn, errId) {
      if (!last) return;
      var t = textFn();
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(t).then(function () { flash(errId, 'Report copied to clipboard.'); },
          function () { fallbackCopy(t, errId); });
      } else fallbackCopy(t, errId);
    }
    function fallbackCopy(t, errId) {
      var ta = document.createElement('textarea');
      ta.value = t; document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); flash(errId, 'Report copied to clipboard.'); }
      catch (e) { showError(errId, 'Copy failed - clipboard is blocked.'); }
      document.body.removeChild(ta);
    }
    function insertReport(last, textFn, errId) {
      if (!last) return;
      if (!confirm('The report will be written to the current selection and may overwrite existing data. Continue?'))
        return;
      var grid = textFn().split('\n').map(function (line) { return [[line]]; });
      Office.context.document.setSelectedDataAsync(grid, { coercionType: Office.CoercionType.Matrix },
        function (ar) {
          if (ar.status === Office.AsyncResultStatus.Failed) showError(errId, 'Insert failed: ' + ar.error.message);
          else flash(errId, 'Report inserted into the sheet.');
        });
    }
  });
})();
