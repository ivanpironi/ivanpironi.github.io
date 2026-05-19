/* =================================================================
   pipeline-render.js — Static chart rendering for /pipeline/
   Depends on: Chart.js 4.4.0 (global Chart), pipeline-data.js
   ================================================================= */
(function () {
  'use strict';

  /* ── Helpers ──────────────────────────────────────────────────── */

  function safeNum(v, fallback) {
    var n = Number(v);
    return isNaN(n) ? (fallback !== undefined ? fallback : 0) : n;
  }

  function fmtM(k) {
    if (k == null || isNaN(k)) return '—';
    var m = k / 1000;
    if (Math.abs(m) >= 1) return '$' + m.toFixed(1) + 'M';
    return '$' + Math.round(k) + 'K';
  }

  function toM(k) { return Math.round(safeNum(k) / 100) / 10; }

  function yearOf(month) {
    var m = safeNum(month, 1);
    if (m <= 12) return 'y1';
    if (m <= 24) return 'y2';
    return 'y3';
  }

  /* ── Aggregations ─────────────────────────────────────────────── */

  function aggregateByYear(engagements, nrfRatio) {
    var r = {
      y1: { sales: 0, revenue: 0, nrf: 0 },
      y2: { sales: 0, revenue: 0, nrf: 0 },
      y3: { sales: 0, revenue: 0, nrf: 0 },
      total: { sales: 0, revenue: 0, nrf: 0 }
    };
    var ratio = safeNum(nrfRatio, 0.75);

    (engagements || []).forEach(function (e) {
      var sm  = safeNum(e.saleMonth, 1);
      var dur = Math.max(1, safeNum(e.duration, 1));
      var wk  = safeNum(e.weightedK);
      var yr  = yearOf(sm);

      r[yr].sales    += wk;
      r.total.sales  += wk;

      var perMonth = wk / dur;
      for (var m = sm; m < sm + dur; m++) {
        var ry = yearOf(m);
        if (r[ry]) { r[ry].revenue += perMonth; r.total.revenue += perMonth; }
      }
    });

    ['y1', 'y2', 'y3', 'total'].forEach(function (yr) {
      r[yr].nrf = r[yr].revenue * ratio;
    });
    return r;
  }

  function monthlyTrajectory(engagements, nrfRatio) {
    var ratio  = safeNum(nrfRatio, 0.75);
    var months = Array.from({ length: 36 }, function (_, i) {
      return { month: i + 1, sales: 0, revenue: 0, nrf: 0,
               cumSales: 0, cumRevenue: 0, cumNrf: 0 };
    });

    (engagements || []).forEach(function (e) {
      var sm  = safeNum(e.saleMonth, 1);
      var dur = Math.max(1, safeNum(e.duration, 1));
      var wk  = safeNum(e.weightedK);

      if (sm >= 1 && sm <= 36) months[sm - 1].sales += wk;

      var perMonth = wk / dur;
      for (var m = sm; m < sm + dur; m++) {
        if (m >= 1 && m <= 36) months[m - 1].revenue += perMonth;
      }
    });

    var cumS = 0, cumR = 0, cumN = 0;
    months.forEach(function (mo) {
      mo.nrf       = mo.revenue * ratio;
      cumS        += mo.sales;    mo.cumSales    = cumS;
      cumR        += mo.revenue;  mo.cumRevenue  = cumR;
      cumN        += mo.nrf;      mo.cumNrf      = cumN;
    });
    return months;
  }

  function accountRanking(engagements) {
    var map = {}, totalWk = 0;

    (engagements || []).forEach(function (e) {
      var name   = String(e.account || e.client || e.name || 'Unknown');
      var region = String(e.region || e.geo || '').toUpperCase();
      if (['KSA', 'SAUDI', 'SA'].indexOf(region) !== -1) region = 'KSA';
      else if (['UAE', 'DUBAI', 'AUH', 'ABU DHABI'].indexOf(region) !== -1) region = 'UAE';
      else if (['QATAR', 'QAT', 'DOH'].indexOf(region) !== -1) region = 'Qatar';
      var wk = safeNum(e.weightedK);
      if (!map[name]) map[name] = { name: name, region: region, weightedK: 0 };
      map[name].weightedK += wk;
      totalWk += wk;
    });

    return Object.values(map)
      .sort(function (a, b) { return b.weightedK - a.weightedK; })
      .slice(0, 10)
      .map(function (a) {
        return { name: a.name, region: a.region, weightedK: a.weightedK,
                 pctOfTotal: totalWk > 0 ? a.weightedK / totalWk : 0 };
      });
  }

  function typeBreakdown(engagements) {
    var r = { A: 0, S: 0, R: 0, RR: 0, CP: 0 };
    (engagements || []).forEach(function (e) {
      var t = String(e.type || '').toUpperCase();
      if (Object.prototype.hasOwnProperty.call(r, t)) r[t] += safeNum(e.weightedK);
      else r.A += safeNum(e.weightedK);
    });
    return r;
  }

  function probabilityBands(engagements) {
    var r = { high: 0, midHigh: 0, midLow: 0, low: 0 };
    (engagements || []).forEach(function (e) {
      var p  = safeNum(e.prob);
      if (p > 1) p = p / 100;
      var wk = safeNum(e.weightedK);
      if      (p >= 0.50) r.high    += wk;
      else if (p >= 0.35) r.midHigh += wk;
      else if (p >= 0.25) r.midLow  += wk;
      else                r.low     += wk;
    });
    return r;
  }

  function funnelCascade(engagements, nrfRatio) {
    var ratio = safeNum(nrfRatio, 0.75);
    var gross = 0, weighted = 0;
    (engagements || []).forEach(function (e) {
      var wk = safeNum(e.weightedK);
      var p  = safeNum(e.prob);
      if (p > 1) p = p / 100;
      gross    += p > 0 ? wk / p : wk;
      weighted += wk;
    });
    return { gross: gross, weighted: weighted,
             revenue: weighted, nrf: weighted * ratio };
  }

  /* ── Chart shared config ──────────────────────────────────────── */

  var GRID  = '#F3F4F6';
  var MUTED = '#6B7280';

  var TOOLTIP = {
    backgroundColor: '#111827', titleColor: '#F9FAFB',
    bodyColor: '#D1D5DB', padding: 10, cornerRadius: 4
  };

  function scalesXY(xCb, yCb) {
    return {
      x: { grid: { color: GRID }, ticks: { color: MUTED, callback: xCb || undefined } },
      y: { grid: { color: GRID }, ticks: { color: MUTED, callback: yCb || undefined } }
    };
  }

  function mCb(v) { return '$' + v + 'M'; }

  /* ── DOM: KPI cards ───────────────────────────────────────────── */

  function renderKPIs(agg) {
    [['kpi-sales', 'sales'], ['kpi-revenue', 'revenue'], ['kpi-nrf', 'nrf']]
      .forEach(function (pair) {
        var el = document.getElementById(pair[0]);
        if (!el) return;
        var metric = pair[1];
        var v = el.querySelector('.value');
        var b = el.querySelector('.breakdown');
        if (v) v.textContent = fmtM(agg.total[metric]);
        if (b) b.innerHTML =
          'Y1 ' + fmtM(agg.y1[metric]) +
          ' · Y2 ' + fmtM(agg.y2[metric]) +
          ' · Y3 ' + fmtM(agg.y3[metric]);
      });
  }

  /* ── Chart: By year ───────────────────────────────────────────── */

  function renderChartByYear(canvasId, agg) {
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: ['Year 1', 'Year 2', 'Year 3'],
        datasets: [
          { label: 'Sales',   data: ['y1','y2','y3'].map(function(y){return toM(agg[y].sales);}),   backgroundColor: '#3B82F6' },
          { label: 'Revenue', data: ['y1','y2','y3'].map(function(y){return toM(agg[y].revenue);}), backgroundColor: '#10B981' },
          { label: 'NRF',     data: ['y1','y2','y3'].map(function(y){return toM(agg[y].nrf);}),     backgroundColor: '#059669' }
        ]
      },
      options: {
        animation: false, responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: true, position: 'top',
          labels: { color: MUTED, boxWidth: 12, font: { size: 11 } } }, tooltip: TOOLTIP },
        scales: scalesXY(null, mCb)
      }
    });
  }

  /* ── Chart: 36-month cumulative trajectory ────────────────────── */

  function renderChartTrajectory(canvasId, monthly) {
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    var labels = monthly.map(function (m) {
      return m.month % 6 === 0 ? 'M' + m.month : '';
    });
    new Chart(ctx, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          { label: 'Cum. Sales',   data: monthly.map(function(m){return toM(m.cumSales);}),   borderColor: '#3B82F6', backgroundColor: 'transparent', pointRadius: 0, tension: 0.3 },
          { label: 'Cum. Revenue', data: monthly.map(function(m){return toM(m.cumRevenue);}), borderColor: '#10B981', backgroundColor: 'transparent', pointRadius: 0, tension: 0.3 },
          { label: 'Cum. NRF',     data: monthly.map(function(m){return toM(m.cumNrf);}),     borderColor: '#059669', backgroundColor: 'transparent', pointRadius: 0, tension: 0.3, borderDash: [4, 3] }
        ]
      },
      options: {
        animation: false, responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: true, position: 'top',
          labels: { color: MUTED, boxWidth: 12, font: { size: 11 } } }, tooltip: TOOLTIP },
        scales: scalesXY(null, mCb)
      }
    });
  }

  /* ── Chart: Accounts ranked ───────────────────────────────────── */

  var REGION_COLOR = { KSA: '#3B82F6', UAE: '#14B8A6', Qatar: '#F59E0B' };

  function renderChartAccounts(canvasId, ranking) {
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: ranking.map(function (a) { return a.name; }),
        datasets: [{
          label: 'Weighted $M',
          data: ranking.map(function (a) { return toM(a.weightedK); }),
          backgroundColor: ranking.map(function (a) { return REGION_COLOR[a.region] || '#6B7280'; })
        }]
      },
      options: {
        animation: false, responsive: true, maintainAspectRatio: false,
        indexAxis: 'y',
        plugins: { legend: { display: false }, tooltip: TOOLTIP },
        scales: {
          x: { grid: { color: GRID }, ticks: { color: MUTED, callback: mCb } },
          y: { grid: { color: GRID }, ticks: { color: MUTED, font: { size: 11 } } }
        }
      }
    });
  }

  /* ── Chart: Engagement type doughnut ──────────────────────────── */

  function renderChartTypes(canvasId, types) {
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    var LABELS = ['Anchor', 'Secondary', 'Retainer', 'Renewal', 'CP'];
    var KEYS   = ['A',      'S',         'R',        'RR',      'CP'];
    var COLORS = ['#3B82F6','#10B981',   '#F59E0B',  '#8B5CF6', '#6B7280'];
    new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: LABELS,
        datasets: [{
          data: KEYS.map(function (k) { return toM(types[k] || 0); }),
          backgroundColor: COLORS, borderWidth: 1, borderColor: '#FFFFFF'
        }]
      },
      options: {
        animation: false, responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { display: true, position: 'bottom',
            labels: { color: MUTED, boxWidth: 12, font: { size: 11 } } },
          tooltip: Object.assign({}, TOOLTIP, {
            callbacks: { label: function (c) { return c.label + ': $' + c.raw + 'M'; } }
          })
        }
      }
    });
  }

  /* ── Chart: Probability bands ─────────────────────────────────── */

  function renderChartProb(canvasId, bands) {
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: ['High\n(≥50%)', 'Mid-High\n(35–49%)', 'Mid-Low\n(25–34%)', 'Low\n(<25%)'],
        datasets: [{
          label: 'Weighted $M',
          data: [toM(bands.high), toM(bands.midHigh), toM(bands.midLow), toM(bands.low)],
          backgroundColor: ['#059669', '#10B981', '#F59E0B', '#EF4444']
        }]
      },
      options: {
        animation: false, responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: TOOLTIP },
        scales: scalesXY(null, mCb)
      }
    });
  }

  /* ── Chart: Gross → NRF funnel ────────────────────────────────── */

  function renderChartFunnel(canvasId, cascade) {
    var ctx = document.getElementById(canvasId);
    if (!ctx) return;
    new Chart(ctx, {
      type: 'bar',
      data: {
        labels: ['Gross pipeline', 'Weighted sales', 'Revenue', 'NRF to EY'],
        datasets: [{
          label: '$M',
          data: [toM(cascade.gross), toM(cascade.weighted), toM(cascade.revenue), toM(cascade.nrf)],
          backgroundColor: ['#BFDBFE', '#3B82F6', '#10B981', '#059669']
        }]
      },
      options: {
        animation: false, responsive: true, maintainAspectRatio: false,
        indexAxis: 'y',
        plugins: { legend: { display: false }, tooltip: TOOLTIP },
        scales: {
          x: { grid: { color: GRID }, ticks: { color: MUTED, callback: mCb } },
          y: { grid: { color: GRID }, ticks: { color: MUTED } }
        }
      }
    });
  }

  /* ── main ─────────────────────────────────────────────────────── */

  async function main() {
    if (typeof PipelineData === 'undefined') {
      console.warn('[pipeline-render] PipelineData not found — pipeline-data.js missing or not yet loaded');
      return;
    }
    try { await PipelineData.load(); }
    catch (err) { console.error('[pipeline-render] JSON load failed:', err); return; }

    var data = PipelineData.data;
    if (!data) { console.error('[pipeline-render] PipelineData.data is empty'); return; }

    var eng   = data.engagements || [];
    var ratio = safeNum((data.assumptions || {}).nrfRatio, 0.75);

    var agg    = aggregateByYear(eng, ratio);
    var mo     = monthlyTrajectory(eng, ratio);
    var ranked = accountRanking(eng);
    var types  = typeBreakdown(eng);
    var bands  = probabilityBands(eng);
    var funnel = funnelCascade(eng, ratio);

    renderKPIs(agg);
    renderChartByYear('chart-by-year', agg);
    renderChartTrajectory('chart-trajectory', mo);
    renderChartAccounts('chart-accounts', ranked);
    renderChartTypes('chart-types', types);
    renderChartProb('chart-prob', bands);
    renderChartFunnel('chart-funnel', funnel);
  }

  document.addEventListener('DOMContentLoaded', main);
})();
