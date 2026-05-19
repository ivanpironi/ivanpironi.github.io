/* =================================================================
   pipeline-data.js — Engagement data + PipelineData API
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
    /* ── Qiddiya · KSA · 6 engagements ─────────────────────── */
    mk('Qiddiya','KSA','A', 3,18,1400,0.50),
    mk('Qiddiya','KSA','A',16,12, 700,0.50),
    mk('Qiddiya','KSA','S', 6,12, 500,0.30),
    mk('Qiddiya','KSA','S',20, 9, 333,0.30),
    mk('Qiddiya','KSA','R',13,24, 500,0.50),
    mk('Qiddiya','KSA','RR',25,12, 100,0.50),

    /* ── FIFA World Cup 2034 · KSA · 5 ─────────────────────── */
    mk('FIFA 2034','KSA','A',10,12,1200,0.50),
    mk('FIFA 2034','KSA','A',23,15, 900,0.40),
    mk('FIFA 2034','KSA','S',14, 9, 400,0.30),
    mk('FIFA 2034','KSA','CP', 8, 6, 100,0.80),
    mk('FIFA 2034','KSA','R',18,12, 160,0.50),

    /* ── Diriyah Gate · KSA · 5 ────────────────────────────── */
    mk('Diriyah Gate','KSA','A',12,12,1000,0.40),
    mk('Diriyah Gate','KSA','A',26,12, 550,0.35),
    mk('Diriyah Gate','KSA','S',16, 9, 400,0.30),
    mk('Diriyah Gate','KSA','S',28, 6, 200,0.25),
    mk('Diriyah Gate','KSA','R',20,24, 140,0.50),

    /* ── New Murabba · KSA · 4 ──────────────────────────────── */
    mk('New Murabba','KSA','A',15,12, 700,0.40),
    mk('New Murabba','KSA','A',28,12, 550,0.35),
    mk('New Murabba','KSA','S',18, 9, 200,0.25),
    mk('New Murabba','KSA','CP',14, 6, 100,0.70),

    /* ── NEOM · KSA · 6 ─────────────────────────────────────── */
    mk('NEOM','KSA','A', 7,15,1200,0.50),
    mk('NEOM','KSA','A',20,18, 900,0.45),
    mk('NEOM','KSA','S',10,12, 300,0.35),
    mk('NEOM','KSA','S',22, 9, 280,0.30),
    mk('NEOM','KSA','R',16,24, 240,0.50),
    mk('NEOM','KSA','RR',30,12, 180,0.50),

    /* ── ROSHN · KSA · 3 ────────────────────────────────────── */
    mk('ROSHN','KSA','A',19,12, 700,0.35),
    mk('ROSHN','KSA','S',23, 9, 300,0.25),
    mk('ROSHN','KSA','CP',17, 6,  50,0.80),

    /* ── Aldar · UAE · 7 ────────────────────────────────────── */
    mk('Aldar','UAE','A', 2,15,1000,0.50),
    mk('Aldar','UAE','A',17,12, 650,0.45),
    mk('Aldar','UAE','S', 5, 9, 400,0.35),
    mk('Aldar','UAE','S',19, 9, 300,0.30),
    mk('Aldar','UAE','R',11,24, 200,0.50),
    mk('Aldar','UAE','RR',23,24, 200,0.50),
    mk('Aldar','UAE','CP', 9, 6,  79,0.70),

    /* ── Miral · UAE · 5 ────────────────────────────────────── */
    mk('Miral','UAE','A', 5,12, 800,0.45),
    mk('Miral','UAE','A',21,12, 500,0.35),
    mk('Miral','UAE','S',10, 9, 350,0.30),
    mk('Miral','UAE','R',15,24, 160,0.50),
    mk('Miral','UAE','CP', 8, 6,  71,0.70),

    /* ── MODON · UAE · 4 ────────────────────────────────────── */
    mk('MODON','UAE','A',14,12, 700,0.35),
    mk('MODON','UAE','S',20, 9, 400,0.25),
    mk('MODON','UAE','S',27, 9, 280,0.25),
    mk('MODON','UAE','CP',12, 6,  50,0.70),

    /* ── TAQA · UAE · 6 ─────────────────────────────────────── */
    mk('TAQA','UAE','A', 4,18, 800,0.50),
    mk('TAQA','UAE','A',22,12, 600,0.40),
    mk('TAQA','UAE','S', 8,12, 350,0.30),
    mk('TAQA','UAE','R',14,24, 100,0.50),
    mk('TAQA','UAE','RR',26,12, 100,0.50),
    mk('TAQA','UAE','CP', 7, 6,  71,0.70),

    /* ── Khazna · UAE · 5 ───────────────────────────────────── */
    mk('Khazna','UAE','A', 6,12, 800,0.50),
    mk('Khazna','UAE','A',21,12, 500,0.35),
    mk('Khazna','UAE','S',10, 9, 280,0.25),
    mk('Khazna','UAE','R',15,24, 100,0.50),
    mk('Khazna','UAE','CP', 9, 6,  71,0.70),

    /* ── Emaar · UAE · 3 ────────────────────────────────────── */
    mk('Emaar','UAE','A',10,12, 700,0.35),
    mk('Emaar','UAE','S',17, 9, 300,0.25),
    mk('Emaar','UAE','CP',14, 6,  57,0.70),

    /* ── TASMU · Qatar · 7 ──────────────────────────────────── */
    mk('TASMU','Qatar','A', 2,18, 700,0.55),
    mk('TASMU','Qatar','A',20,12, 500,0.45),
    mk('TASMU','Qatar','S', 5,12, 400,0.35),
    mk('TASMU','Qatar','S',22, 9, 300,0.30),
    mk('TASMU','Qatar','R',12,24, 160,0.50),
    mk('TASMU','Qatar','RR',24,12, 100,0.50),
    mk('TASMU','Qatar','CP', 8, 6,  57,0.70),

    /* ── Ministry of Municipality Qatar · 4 ────────────────── */
    mk('MoM Qatar','Qatar','A',15,12, 750,0.35),
    mk('MoM Qatar','Qatar','S',21, 9, 350,0.25),
    mk('MoM Qatar','Qatar','R',18,24, 120,0.50),
    mk('MoM Qatar','Qatar','CP',13, 6,  57,0.70),

    /* ── Ashghal · Qatar · 8 (6 Y3 origination) ────────────── */
    mk('Ashghal','Qatar','A',13, 9, 200,0.40),
    mk('Ashghal','Qatar','S',17, 6, 200,0.30),
    mk('Ashghal','Qatar','A',26, 9, 167,0.30,{isY3Origination:true}),
    mk('Ashghal','Qatar','A',28, 9, 167,0.30,{isY3Origination:true}),
    mk('Ashghal','Qatar','A',29, 9, 167,0.30,{isY3Origination:true}),
    mk('Ashghal','Qatar','A',30, 9, 167,0.30,{isY3Origination:true}),
    mk('Ashghal','Qatar','A',32, 9, 167,0.30,{isY3Origination:true}),
    mk('Ashghal','Qatar','A',34, 6, 167,0.30,{isY3Origination:true})
  ];

  var META = {
    lastUpdated:       '2026-05-19',
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
