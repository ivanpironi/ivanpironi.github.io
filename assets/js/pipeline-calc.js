/* =================================================================
   pipeline-calc.js — Pure recalculation engine
   No DOM access. Exports window.PipelineCalc.
   ================================================================= */
(function () {
  'use strict';

  function safeNum(v, fallback) {
    var n = Number(v);
    return isNaN(n) ? (fallback !== undefined ? fallback : 0) : n;
  }

  function yearOf(month) {
    var m = safeNum(month, 1);
    return m <= 12 ? 1 : m <= 24 ? 2 : 3;
  }

  /* ── applyOverrides ─────────────────────────────────────────────
     overrides = { qiddiyaProb (25-75), retainerConvRate (25-75),
                   y3OrigProb (0-50), nrfHaircut (15-35) }
     Returns a new array — input is never mutated.
  ─────────────────────────────────────────────────────────────── */
  function applyOverrides(engagements, overrides) {
    var ov = overrides || {};
    return (engagements || []).map(function (e) {
      var origProb = safeNum(e.prob);
      if (origProb > 1) origProb = origProb / 100; // normalise 0-100 → 0-1

      /* Reconstruct gross value (prefer explicit grossK field) */
      var grossK = e.grossK != null
        ? safeNum(e.grossK)
        : (origProb > 0 ? safeNum(e.weightedK) / origProb : safeNum(e.weightedK));

      var newProb = origProb; /* default: unchanged */
      var account = String(e.account || e.client || e.name || '').toLowerCase();

      if (account === 'qiddiya') {
        if (e.type === 'A')      newProb = safeNum(ov.qiddiyaProb, 50) / 100;
        else if (e.type === 'S') newProb = (safeNum(ov.qiddiyaProb, 50) / 100) * 0.6;
      } else if (e.type === 'R' || e.type === 'RR') {
        /* Scale relative to default 50: slider 25 → ×0.5, slider 75 → ×1.5 */
        newProb = Math.min(1, origProb * (safeNum(ov.retainerConvRate, 50) / 50));
      } else if (e.isY3Origination === true) {
        newProb = safeNum(ov.y3OrigProb, 30) / 100;
      }

      return Object.assign({}, e, {
        prob:      newProb,
        weightedK: Math.round(grossK * newProb * 10) / 10, /* round to 1dp */
        _grossK:   grossK
      });
    });
  }

  /* ── applyFilters ───────────────────────────────────────────────
     filters = { years: [1,2,3], types: ['A','S',...] }
     Empty array → no filter applied for that dimension.
  ─────────────────────────────────────────────────────────────── */
  function applyFilters(engagements, filters) {
    var f = filters || {};
    var years = f.years && f.years.length ? f.years : null;
    var types = f.types && f.types.length ? f.types : null;

    return (engagements || []).filter(function (e) {
      if (years) {
        var yr = yearOf(safeNum(e.saleMonth, 1));
        if (years.indexOf(yr) === -1) return false;
      }
      if (types) {
        var t = String(e.type || '').toUpperCase();
        if (types.indexOf(t) === -1) return false;
      }
      return true;
    });
  }

  window.PipelineCalc = { applyOverrides: applyOverrides, applyFilters: applyFilters };
})();
