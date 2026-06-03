/* =================================================================
   pipeline-data.js — Engagement data + PipelineData API
   Canonical 15 accounts · KSA 11 · UAE 3 · Qatar 1
   Year 1 (4): Qiddiya, SAR, Expo 2030, FIFA 2034
   Year 2 (5): Diriyah Gate, ROSHN, KSPF, KSIA, Khazna
   Year 3 (6): Saudi Aramco, New Murabba, Red Sea Global, ADMD, TAQA, TASMU
   Exposes window.PipelineData = { load(), data }
   ================================================================= */
(function () {
  'use strict';

  function mk(account, region, type, sm, dur, gK, prob, extra) {
    return Object.assign({
      account: account, region: region, type: type,
      saleMonth: sm, duration: dur, grossK: gK, prob: prob,
      weightedK: Math.round(gK * prob * 10) / 10
    }, extra || {});
  }

  var ENGAGEMENTS = [
    /* ── Qiddiya · KSA · Year 1 · 6 engagements ────────────── */
    mk('Qiddiya','KSA','A', 3,18,1400,0.50),
    mk('Qiddiya','KSA','A',16,12, 700,0.50),
    mk('Qiddiya','KSA','S', 6,12, 500,0.30),
    mk('Qiddiya','KSA','S',20, 9, 333,0.30),
    mk('Qiddiya','KSA','R',13,24, 500,0.50),
    mk('Qiddiya','KSA','RR',25,12, 100,0.50),

    /* ── SAR (Saudi Railways) · KSA · Year 1 · 4 engagements ─ */
    mk('SAR','KSA','CP', 4, 4,  60,0.75),
    mk('SAR','KSA','A',  6, 9, 350,0.45),
    mk('SAR','KSA','S', 10, 6, 200,0.30),
    mk('SAR','KSA','R', 13,24, 100,0.50),

    /* ── Expo 2030 Riyadh · KSA · Year 1 · 5 engagements ───── */
    mk('Expo 2030','KSA','CP', 5, 6,  80,0.80),
    mk('Expo 2030','KSA','A',  8,12, 900,0.50),
    mk('Expo 2030','KSA','S', 12, 9, 400,0.30),
    mk('Expo 2030','KSA','R', 15,24, 160,0.50),
    mk('Expo 2030','KSA','RR',27,12, 120,0.50),

    /* ── FIFA 2034 · KSA · Year 1 · 5 engagements ──────────── */
    mk('FIFA 2034','KSA','A',10,12,1200,0.50),
    mk('FIFA 2034','KSA','A',23,15, 900,0.40),
    mk('FIFA 2034','KSA','S',14, 9, 400,0.30),
    mk('FIFA 2034','KSA','CP', 8, 6, 100,0.80),
    mk('FIFA 2034','KSA','R',18,12, 160,0.50),

    /* ── Diriyah Gate · KSA · Year 2 · 5 engagements ────────── */
    mk('Diriyah Gate','KSA','A',12,12,1000,0.40),
    mk('Diriyah Gate','KSA','A',26,12, 550,0.35),
    mk('Diriyah Gate','KSA','S',16, 9, 400,0.30),
    mk('Diriyah Gate','KSA','S',28, 6, 200,0.25),
    mk('Diriyah Gate','KSA','R',20,24, 140,0.50),

    /* ── New Murabba · KSA · Year 3 · 4 engagements ─────────── */
    mk('New Murabba','KSA','CP',25, 6, 100,0.70),
    mk('New Murabba','KSA','A', 27,12, 700,0.40),
    mk('New Murabba','KSA','S', 29, 9, 200,0.25),
    mk('New Murabba','KSA','A', 32,12, 550,0.35),

    /* ── ROSHN · KSA · Year 2 · 4 engagements ───────────────── */
    mk('ROSHN','KSA','CP',17, 6,  50,0.80),
    mk('ROSHN','KSA','A', 19,12, 700,0.35),
    mk('ROSHN','KSA','R', 22,24, 100,0.40),
    mk('ROSHN','KSA','S', 23, 9, 300,0.25),

    /* ── KSPF (King Salman Park) · KSA · Year 2 · 5 engag. ─── */
    mk('KSPF','KSA','CP',13, 6,  71,0.75),
    mk('KSPF','KSA','A', 15,12, 900,0.40),
    mk('KSPF','KSA','S', 20, 9, 400,0.25),
    mk('KSPF','KSA','R', 19,24, 140,0.50),
    mk('KSPF','KSA','RR',31,12, 100,0.50),

    /* ── KSIA (King Salman Airport) · KSA · Year 2 · 6 engag. ─ */
    mk('KSIA','KSA','CP',15, 6,  71,0.70),
    mk('KSIA','KSA','A', 17,12,1000,0.35),
    mk('KSIA','KSA','R', 21,24, 160,0.50),
    mk('KSIA','KSA','S', 22, 9, 400,0.25),
    mk('KSIA','KSA','A', 30,15, 700,0.30),
    mk('KSIA','KSA','RR',33,12, 120,0.50),

    /* ── Khazna Data Centres · UAE · Year 2 · 5 engagements ─── */
    mk('Khazna','UAE','CP',13, 6,  71,0.70),
    mk('Khazna','UAE','A', 14,12, 600,0.45),
    mk('Khazna','UAE','R', 16,24, 100,0.50),
    mk('Khazna','UAE','S', 18, 9, 280,0.25),
    mk('Khazna','UAE','A', 22,12, 400,0.35),

    /* ── Saudi Aramco · KSA · Year 3 · 4 engagements ────────── */
    mk('Saudi Aramco','KSA','CP',26, 6, 100,0.60),
    mk('Saudi Aramco','KSA','A', 25,12,1200,0.30),
    mk('Saudi Aramco','KSA','R', 28,24, 160,0.30),
    mk('Saudi Aramco','KSA','S', 29, 9, 500,0.20),

    /* ── Red Sea Global · KSA · Year 3 · 4 engagements ──────── */
    mk('Red Sea Global','KSA','CP',27, 6,  71,0.60),
    mk('Red Sea Global','KSA','A', 26,12, 700,0.35),
    mk('Red Sea Global','KSA','R', 29,24, 120,0.35),
    mk('Red Sea Global','KSA','S', 30, 9, 300,0.25),

    /* ── ADMD (Aldar · Miral · MODON) · UAE · Year 3 · 8 engag. */
    mk('ADMD','UAE','CP',27, 6,  71,0.70),
    mk('ADMD','UAE','CP',30, 6,  57,0.70),
    mk('ADMD','UAE','A', 26,15,1000,0.40),
    mk('ADMD','UAE','S', 28, 9, 400,0.30),
    mk('ADMD','UAE','R', 27,24, 160,0.50),
    mk('ADMD','UAE','A', 32,12, 700,0.35),
    mk('ADMD','UAE','S', 33, 9, 300,0.25),
    mk('ADMD','UAE','RR',34,24, 140,0.45),

    /* ── TAQA · UAE · Year 3 · 6 engagements ────────────────── */
    mk('TAQA','UAE','CP',26, 6,  71,0.65),
    mk('TAQA','UAE','A', 27,18, 800,0.40),
    mk('TAQA','UAE','R', 28,24, 100,0.45),
    mk('TAQA','UAE','S', 29,12, 350,0.30),
    mk('TAQA','UAE','A', 33,12, 500,0.35),
    mk('TAQA','UAE','RR',34,12,  80,0.40),

    /* ── TASMU · Qatar · Year 3 · 7 engagements ─────────────── */
    mk('TASMU','Qatar','CP',26, 6,  57,0.65),
    mk('TASMU','Qatar','A', 25,18, 700,0.45),
    mk('TASMU','Qatar','R', 28,24, 120,0.45),
    mk('TASMU','Qatar','S', 27,12, 350,0.30),
    mk('TASMU','Qatar','A', 32,12, 500,0.40),
    mk('TASMU','Qatar','S', 33, 9, 250,0.25),
    mk('TASMU','Qatar','RR',34,12,  80,0.40)
  ];

  var META = {
    lastUpdated:       '2026-06-03',
    accounts:          15,
    totalEngagements:  ENGAGEMENTS.length
  };

  function populateSpans() {
    document.querySelectorAll('[data-pipeline]').forEach(function (el) {
      var key = el.getAttribute('data-pipeline');
      if (key === 'lastUpdated')       el.textContent = META.lastUpdated;
      if (key === 'totalEngagements')  el.textContent = String(META.totalEngagements);
    });
  }

  window.PipelineData = {
    data: null,
    load: function () {
      return Promise.resolve().then(function () {
        window.PipelineData.data = { engagements: ENGAGEMENTS, meta: META };
        if (document.readyState === 'loading') {
          document.addEventListener('DOMContentLoaded', populateSpans);
        } else {
          populateSpans();
        }
      });
    }
  };
})();
