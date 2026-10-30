/* LabKitPro Excel add-in - v1.0 Cpk slice.
   Uses the shared offline kernel (calc.js -> window.LKPCalc.capability).
   Data flow: worksheet selection (matrix) -> parse -> capability -> render. */
(function () {
  'use strict';

  Office.onReady(function () {
    var $ = function (id) { return document.getElementById(id); };

    $('analyze').addEventListener('click', analyze);
    $('loadSample').addEventListener('click', function (e) { e.preventDefault(); fillSample(); });
    $('copyReport').addEventListener('click', copyReport);
    $('insertReport').addEventListener('click', insertReport);

    // v1.1 tabs respond honestly instead of doing nothing
    Array.prototype.forEach.call(document.querySelectorAll('.tab.soon'), function (t) {
      t.addEventListener('click', function () {
        var name = t.textContent.replace('v1.1', '').trim();
        var e = $('error');
        e.innerHTML = '<b>' + name + '</b> ships in v1.1 &mdash; the engine is already live: '
          + '<a href="https://labkitpro.com/tools/" target="_blank">try it on labkitpro.com</a>. '
          + 'This Excel tab arrives with the next free update.';
        e.classList.remove('hidden'); e.classList.add('ok');
      });
    });

    function fillSample() {
      $('lsl').value = '9.80';
      $('usl').value = '10.20';
      $('target').value = '10.00';
      // write sample data into the selection area so the user sees a live demo
      var rows = [];
      var data = [10.02, 9.97, 10.11, 10.05, 9.94, 10.08, 10.01, 9.99, 10.06, 10.03,
                  9.96, 10.09, 10.00, 10.04, 9.98, 10.07, 10.02, 9.95, 10.10, 10.01,
                  10.05, 9.99, 10.03, 10.06, 9.97];
      for (var i = 0; i < data.length; i++) rows.push([[data[i]]]);
      Office.context.document.setSelectedDataAsync(
        rows, { coercionType: Office.CoercionType.Matrix }, function (ar) {
          if (ar.status === Office.AsyncResultStatus.Failed) showError('Could not write sample data: ' + ar.error.message);
        });
    }

    function analyze() {
      hideError();
      var lsl = optNum($('lsl').value);
      var usl = optNum($('usl').value);
      var target = optNum($('target').value);
      if (lsl === null && usl === null) return showError('Enter at least one spec limit (LSL or USL).');

      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Matrix,
        { valueFormat: Office.ValueFormat.Unformatted },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed)
            return showError('Could not read selection: ' + asyncResult.error.message);

          var matrix = asyncResult.value;
          if (!matrix || !matrix.length) return showError('Select a range of measurement values first.');

          // parse per column: each column = candidate subgroup (Excel-native flow)
          var cols = [];
          for (var c = 0; c < matrix[0].length; c++) cols.push([]);
          var all = [];
          for (var r = 0; r < matrix.length; r++) {
            var row = matrix[r];
            for (var c2 = 0; c2 < row.length; c2++) {
              var v = typeof row[c2] === 'number' ? row[c2] : parseFloat(String(row[c2]).trim());
              if (isFinite(v)) { cols[c2].push(v); all.push(v); }
            }
          }
          var numericCols = cols.filter(function (col) { return col.length >= 2; });
          var values, subgroups = null;
          if (numericCols.length >= 2) {
            // multi-column selection -> each column = one subgroup (enables Cp/Cpk)
            values = [].concat.apply([], numericCols);
            subgroups = numericCols;
          } else {
            values = all;
          }
          if (values.length < 2) return showError('No numeric values found in the selection. Select the cells that contain your measurements.');

          var res = LKPCalc.capability(values, lsl, usl, target, subgroups);
          if (res.error) return showError(res.error);
          render(res);
        });
    }

    function render(res) {
      var eff = (res.cpk != null ? res.cpk : res.ppk);
      $('verdict').textContent = res.verdict + '  (Cpk/Ppk = ' + eff + ')';
      $('verdict').className = 'verdict ' + (res.ok ? 'ok' : eff >= 1.0 ? 'marginal' : 'bad');

      var rows = [
        ['N (values)', res.n],
        ['Mean', res.mean],
        ['Std dev (overall)', res.std]
      ];
      if (res.sigW != null) rows.push(['Sigma within (R-bar/d2)', res.sigW]);
      var spec = [['Pp', res.pp], ['Ppk', res.ppk]];
      if (res.cp != null) spec = [['Cp', res.cp], ['Cpk', res.cpk],
        ['CPU', res.cpu], ['CPL', res.cpl]].concat(spec);
      if (res.cpm != null) spec.push(['Cpm', res.cpm]);
      rows = rows.concat(spec, [['Expected PPM', res.ppm]]);

      var html = '';
      rows.forEach(function (r) {
        var dim = (r[1] === null || r[1] === undefined) ? ' class="dim"' : '';
        html += '<tr' + dim + '><td>' + r[0] + '</td><td>' + (r[1] == null ? 'n/a' : r[1]) + '</td></tr>';
      });
      $('metrics').innerHTML = html;
      $('result').classList.remove('hidden');
      lastResult = res;
    }

    var lastResult = null;

    function reportText() {
      var r = lastResult;
      if (!r) return null;
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

    function copyReport() {
      if (!lastResult) return;
      var t = reportText();
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(t).then(function () { flash('Report copied to clipboard.'); },
          function () { fallbackCopy(t); });
      } else fallbackCopy(t);
    }

    function fallbackCopy(t) {
      var ta = document.createElement('textarea');
      ta.value = t; document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); flash('Report copied to clipboard.'); }
      catch (e) { showError('Copy failed - clipboard is blocked.'); }
      document.body.removeChild(ta);
    }

    function insertReport() {
      if (!lastResult) return;
      var grid = reportText().split('\n').map(function (line) { return [[line]]; });
      Office.context.document.setSelectedDataAsync(grid, { coercionType: Office.CoercionType.Matrix },
        function (ar) {
          if (ar.status === Office.AsyncResultStatus.Failed) showError('Insert failed: ' + ar.error.message);
        });
    }

    // ---------- helpers ----------
    function optNum(s) { s = String(s).trim(); return s === '' ? null : (isFinite(+s) ? +s : null); }
    function showError(m) { var e = $('error'); e.textContent = m; e.classList.remove('hidden'); e.classList.remove('ok'); }
    function hideError() { $('error').classList.add('hidden'); }
    function flash(msg) { var e = $('error'); e.textContent = msg; e.classList.remove('hidden'); e.classList.add('ok'); setTimeout(function () { e.classList.add('hidden'); }, 2500); }
  });
})();
