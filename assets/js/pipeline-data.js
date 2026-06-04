/* =================================================================
   pipeline-data.js — Loads from /data/pipeline-data.json
   Canonical 15 accounts · KSA 11 · UAE 3 · Qatar 1
   Exposes window.PipelineData = { load(), data }
   ================================================================= */
(function () {
  'use strict';

  function populateSpans(meta) {
    document.querySelectorAll('[data-pipeline]').forEach(function (el) {
      var key = el.getAttribute('data-pipeline');
      if (key === 'lastUpdated')      el.textContent = meta.lastUpdated;
      if (key === 'totalEngagements') el.textContent = String(meta.totalEngagements);
    });
  }

  window.PipelineData = {
    data: null,
    load: function () {
      return fetch('/data/pipeline-data.json')
        .then(function (r) {
          if (!r.ok) throw new Error('HTTP ' + r.status);
          return r.json();
        })
        .then(function (json) {
          var engagements = [];
          (json.accounts || []).forEach(function (acc) {
            (acc.engagements || []).forEach(function (eng) {
              engagements.push(Object.assign({}, eng, {
                account:   acc.name,
                region:    acc.region,
                weightedK: Math.round(eng.grossK * eng.prob * 10) / 10
              }));
            });
          });
          var meta = {
            lastUpdated:      (json.meta || {}).lastUpdated || '',
            accounts:         (json.meta || {}).accounts    || 0,
            totalEngagements: engagements.length
          };
          window.PipelineData.data = { engagements: engagements, meta: meta };
          if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', function () { populateSpans(meta); });
          } else {
            populateSpans(meta);
          }
        });
    }
  };
})();
