/* =================================================================
   pipeline-render.js — Rendering + interactivity for /pipeline/
   Depends on: Chart.js 4.4.0, pipeline-data.js, pipeline-calc.js
   ================================================================= */
(function () {
  'use strict';

  /* ── Module-level state ─────────────────────────────────────── */
  var _charts      = {};
  var _initialized = false;
  var _rafPending  = null;

  var DEFAULTS = { qiddiyaProb: 50, retainerConvRate: 50, y3OrigProb: 30, nrfHaircut: 25 };

  /* ── Helpers ─────────────────────────────────────────────────── */
  function safeNum(v, fallback) {
    var n = Number(v);
    return isNaN(n) ? (fallback !== undefined ? fallback : 0) : n;
  }
  function fmtM(k) {
    if (k == null || isNaN(k)) return '—';
    var m = k / 1000;
    return Math.abs(m) >= 1 ? '$' + m.toFixed(1) + 'M' : '$' + Math.round(k) + 'K';
  }
  function toM(k) { return Math.round(safeNum(k) / 100) / 10; }
  function yearOf(month) {
    var m = safeNum(month, 1);
    return m <= 12 ? 'y1' : m <= 24 ? 'y2' : 'y3';
  }

  /* ── Aggregations ────────────────────────────────────────────── */
  function aggregateByYear(engagements, nrfRatio) {
    var r = {
      y1: { sales:0, revenue:0, nrf:0 }, y2: { sales:0, revenue:0, nrf:0 },
      y3: { sales:0, revenue:0, nrf:0 }, total: { sales:0, revenue:0, nrf:0 }
    };
    var ratio = safeNum(nrfRatio, 0.75);
    (engagements || []).forEach(function (e) {
      var sm = safeNum(e.saleMonth, 1), dur = Math.max(1, safeNum(e.duration, 1)), wk = safeNum(e.weightedK);
      var yr = yearOf(sm);
      r[yr].sales += wk; r.total.sales += wk;
      var pm = wk / dur;
      for (var m = sm; m < sm + dur; m++) {
        var ry = yearOf(m);
        if (r[ry]) { r[ry].revenue += pm; r.total.revenue += pm; }
      }
    });
    ['y1','y2','y3','total'].forEach(function (yr) { r[yr].nrf = r[yr].revenue * ratio; });
    return r;
  }

  function monthlyTrajectory(engagements, nrfRatio) {
    var ratio = safeNum(nrfRatio, 0.75);
    var months = Array.from({ length: 36 }, function (_, i) {
      return { month: i+1, sales:0, revenue:0, nrf:0, cumSales:0, cumRevenue:0, cumNrf:0 };
    });
    (engagements || []).forEach(function (e) {
      var sm = safeNum(e.saleMonth,1), dur = Math.max(1,safeNum(e.duration,1)), wk = safeNum(e.weightedK);
      if (sm>=1&&sm<=36) months[sm-1].sales += wk;
      var pm = wk/dur;
      for (var m=sm; m<sm+dur; m++) { if(m>=1&&m<=36) months[m-1].revenue += pm; }
    });
    var cs=0,cr=0,cn=0;
    months.forEach(function(mo){
      mo.nrf=mo.revenue*ratio; cs+=mo.sales; cr+=mo.revenue; cn+=mo.nrf;
      mo.cumSales=cs; mo.cumRevenue=cr; mo.cumNrf=cn;
    });
    return months;
  }

  function accountRanking(engagements) {
    var map={}, total=0;
    (engagements||[]).forEach(function(e){
      var name=String(e.account||e.client||e.name||'Unknown');
      var reg=String(e.region||e.geo||'').toUpperCase();
      if(['KSA','SAUDI','SA'].indexOf(reg)!==-1) reg='KSA';
      else if(['UAE','DUBAI','AUH'].indexOf(reg)!==-1) reg='UAE';
      else if(['QATAR','QAT','DOH'].indexOf(reg)!==-1) reg='Qatar';
      var wk=safeNum(e.weightedK);
      if(!map[name]) map[name]={name:name,region:reg,weightedK:0};
      map[name].weightedK+=wk; total+=wk;
    });
    return Object.values(map).sort(function(a,b){return b.weightedK-a.weightedK;}).slice(0,10)
      .map(function(a){return {name:a.name,region:a.region,weightedK:a.weightedK,
        pctOfTotal:total>0?a.weightedK/total:0};});
  }

  function typeBreakdown(engagements) {
    var r={A:0,S:0,R:0,RR:0,CP:0};
    (engagements||[]).forEach(function(e){
      var t=String(e.type||'').toUpperCase();
      if(Object.prototype.hasOwnProperty.call(r,t)) r[t]+=safeNum(e.weightedK);
      else r.A+=safeNum(e.weightedK);
    });
    return r;
  }

  function probabilityBands(engagements) {
    var r={high:0,midHigh:0,midLow:0,low:0};
    (engagements||[]).forEach(function(e){
      var p=safeNum(e.prob); if(p>1)p=p/100;
      var wk=safeNum(e.weightedK);
      if(p>=0.50) r.high+=wk; else if(p>=0.35) r.midHigh+=wk;
      else if(p>=0.25) r.midLow+=wk; else r.low+=wk;
    });
    return r;
  }

  function funnelCascade(engagements, nrfRatio) {
    var ratio=safeNum(nrfRatio,0.75), gross=0, weighted=0;
    (engagements||[]).forEach(function(e){
      var wk=safeNum(e.weightedK), p=safeNum(e.prob); if(p>1)p=p/100;
      gross+=(p>0?wk/p:wk); weighted+=wk;
    });
    return {gross:gross,weighted:weighted,revenue:weighted,nrf:weighted*ratio};
  }

  /* ── Chart shared config ─────────────────────────────────────── */
  var GRID='rgba(245,244,240,0.08)', MUTED='rgba(245,244,240,0.55)';
  var REGION_COLOR={KSA:'#3B82F6',UAE:'#14B8A6',Qatar:'#F59E0B'};
  var TIP={backgroundColor:'#111827',titleColor:'#F9FAFB',bodyColor:'#D1D5DB',padding:10,cornerRadius:4};
  function mCb(v){return '$'+v+'M';}
  function sc(xCb,yCb){return{x:{grid:{color:GRID},ticks:{color:MUTED,callback:xCb||undefined}},y:{grid:{color:GRID},ticks:{color:MUTED,callback:yCb||undefined}}};}

  /* ── KPI DOM ─────────────────────────────────────────────────── */
  function _kpiRow(agg, metric) {
    return 'Y1 '+fmtM(agg.y1[metric])+' · Y2 '+fmtM(agg.y2[metric])+' · Y3 '+fmtM(agg.y3[metric]);
  }
  function renderKPIs(agg) {
    [['kpi-sales','sales'],['kpi-revenue','revenue'],['kpi-nrf','nrf']].forEach(function(p){
      var el=document.getElementById(p[0]); if(!el) return;
      var v=el.querySelector('.value'), b=el.querySelector('.breakdown');
      if(v) v.textContent=fmtM(agg.total[p[1]]);
      if(b) b.innerHTML=_kpiRow(agg,p[1]);
    });
  }
  function updateKPIs(agg) {
    [['kpi-sales','sales'],['kpi-revenue','revenue'],['kpi-nrf','nrf']].forEach(function(p){
      var el=document.getElementById(p[0]); if(!el) return;
      var v=el.querySelector('.value'), b=el.querySelector('.breakdown');
      if(v){v.style.opacity='0.5'; v.textContent=fmtM(agg.total[p[1]]);
        requestAnimationFrame(function(){v.style.opacity='1';});}
      if(b) b.innerHTML=_kpiRow(agg,p[1]);
    });
  }

  /* ── Charts: CREATE (initial render with animation) ─────────────── */
  function renderChartByYear(id, agg) {
    var ctx=document.getElementById(id); if(!ctx) return;
    _charts[id]=new Chart(ctx,{type:'bar',data:{labels:['Year 1','Year 2','Year 3'],datasets:[
      {label:'TCV',    data:['y1','y2','y3'].map(function(y){return toM(agg[y].sales);}),  backgroundColor:'#3B82F6'},
      {label:'Revenue',data:['y1','y2','y3'].map(function(y){return toM(agg[y].revenue);}),backgroundColor:'#10B981'},
      {label:'NRF',    data:['y1','y2','y3'].map(function(y){return toM(agg[y].nrf);}),    backgroundColor:'#059669'}
    ]},options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,position:'top',labels:{color:MUTED,boxWidth:12,font:{size:11}}},tooltip:TIP},
      scales:sc(null,mCb)}});
  }

  function renderChartTrajectory(id, mo) {
    var ctx=document.getElementById(id); if(!ctx) return;
    var labels=mo.map(function(m){return m.month%6===0?'M'+m.month:'';});
    _charts[id]=new Chart(ctx,{type:'line',data:{labels:labels,datasets:[
      {label:'Cum. TCV',    data:mo.map(function(m){return toM(m.cumSales);}),  borderColor:'#3B82F6',backgroundColor:'transparent',pointRadius:0,tension:0.3},
      {label:'Cum. Revenue',data:mo.map(function(m){return toM(m.cumRevenue);}),borderColor:'#10B981',backgroundColor:'transparent',pointRadius:0,tension:0.3},
      {label:'Cum. NRF',   data:mo.map(function(m){return toM(m.cumNrf);}),    borderColor:'#059669',backgroundColor:'transparent',pointRadius:0,tension:0.3,borderDash:[4,3]}
    ]},options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,position:'top',labels:{color:MUTED,boxWidth:12,font:{size:11}}},tooltip:TIP},
      scales:sc(null,mCb)}});
  }

  function renderChartAccounts(id, ranking) {
    var ctx=document.getElementById(id); if(!ctx) return;
    _charts[id]=new Chart(ctx,{type:'bar',
      data:{labels:ranking.map(function(a){return a.name;}),datasets:[{label:'Weighted $M',
        data:ranking.map(function(a){return toM(a.weightedK);}),
        backgroundColor:ranking.map(function(a){return REGION_COLOR[a.region]||'#6B7280';})}]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{display:false},tooltip:TIP},
        scales:{x:{grid:{color:GRID},ticks:{color:MUTED,callback:mCb}},y:{grid:{color:GRID},ticks:{color:MUTED,font:{size:11}}}},
        onClick:function(evt,els){
          if(!els.length) return;
          var name=_charts[id].data.labels[els[0].index];
          openDrillPanel(name);
        }}});
  }

  function renderChartTypes(id, types) {
    var ctx=document.getElementById(id); if(!ctx) return;
    var LBL=['Anchor','Secondary','Retainer','Renewal','CP'],KEYS=['A','S','R','RR','CP'];
    var CLR=['#3B82F6','#10B981','#F59E0B','#8B5CF6','#6B7280'];
    _charts[id]=new Chart(ctx,{type:'doughnut',
      data:{labels:LBL,datasets:[{data:KEYS.map(function(k){return toM(types[k]||0);}),
        backgroundColor:CLR,borderWidth:1,borderColor:'#111110'}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{
        legend:{display:true,position:'bottom',labels:{color:MUTED,boxWidth:12,font:{size:11}}},
        tooltip:Object.assign({},TIP,{callbacks:{label:function(c){return c.label+': $'+c.raw+'M';}}})}}});
  }

  function renderChartProb(id, bands) {
    var ctx=document.getElementById(id); if(!ctx) return;
    _charts[id]=new Chart(ctx,{type:'bar',data:{
      labels:['High\n(≥50%)','Mid-High\n(35–49%)','Mid-Low\n(25–34%)','Low\n(<25%)'],
      datasets:[{label:'Weighted $M',
        data:[toM(bands.high),toM(bands.midHigh),toM(bands.midLow),toM(bands.low)],
        backgroundColor:['#059669','#10B981','#F59E0B','#EF4444']}]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{display:false},tooltip:TIP},scales:sc(null,mCb)}});
  }

  function renderChartFunnel(id, cascade) {
    var ctx=document.getElementById(id); if(!ctx) return;
    _charts[id]=new Chart(ctx,{type:'bar',data:{
      labels:['Gross pipeline','Weighted TCV','Revenue','NRF to EY'],
      datasets:[{label:'$M',
        data:[toM(cascade.gross),toM(cascade.weighted),toM(cascade.revenue),toM(cascade.nrf)],
        backgroundColor:['#BFDBFE','#3B82F6','#10B981','#059669']}]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
        plugins:{legend:{display:false},tooltip:TIP},
        scales:{x:{grid:{color:GRID},ticks:{color:MUTED,callback:mCb}},y:{grid:{color:GRID},ticks:{color:MUTED}}}}});
  }

  /* ── Charts: UPDATE (live, no animation) ─────────────────────── */
  function updateChartByYear(agg) {
    var c=_charts['chart-by-year']; if(!c) return;
    var ks=['y1','y2','y3'];
    c.data.datasets[0].data=ks.map(function(k){return toM(agg[k].sales);});
    c.data.datasets[1].data=ks.map(function(k){return toM(agg[k].revenue);});
    c.data.datasets[2].data=ks.map(function(k){return toM(agg[k].nrf);});
    c.update('none');
  }
  function updateChartTrajectory(mo) {
    var c=_charts['chart-trajectory']; if(!c) return;
    c.data.datasets[0].data=mo.map(function(m){return toM(m.cumSales);});
    c.data.datasets[1].data=mo.map(function(m){return toM(m.cumRevenue);});
    c.data.datasets[2].data=mo.map(function(m){return toM(m.cumNrf);});
    c.update('none');
  }
  function updateChartAccounts(ranking) {
    var c=_charts['chart-accounts']; if(!c) return;
    c.data.labels=ranking.map(function(a){return a.name;});
    c.data.datasets[0].data=ranking.map(function(a){return toM(a.weightedK);});
    c.data.datasets[0].backgroundColor=ranking.map(function(a){return REGION_COLOR[a.region]||'#6B7280';});
    c.update('none');
  }
  function updateChartTypes(types) {
    var c=_charts['chart-types']; if(!c) return;
    c.data.datasets[0].data=['A','S','R','RR','CP'].map(function(k){return toM(types[k]||0);});
    c.update('none');
  }
  function updateChartProb(bands) {
    var c=_charts['chart-prob']; if(!c) return;
    c.data.datasets[0].data=[toM(bands.high),toM(bands.midHigh),toM(bands.midLow),toM(bands.low)];
    c.update('none');
  }
  function updateChartFunnel(cascade) {
    var c=_charts['chart-funnel']; if(!c) return;
    c.data.datasets[0].data=[toM(cascade.gross),toM(cascade.weighted),toM(cascade.revenue),toM(cascade.nrf)];
    c.update('none');
  }

  /* ── Recalculation ───────────────────────────────────────────── */
  function recalc() {
    var state = window.appState;
    if (!state) return;

    var calc = typeof PipelineCalc !== 'undefined' ? PipelineCalc : null;
    var overridden = calc
      ? calc.applyOverrides(state.originalEngagements, state.overrides)
      : state.originalEngagements;
    var filtered = calc
      ? calc.applyFilters(overridden, state.filters)
      : overridden;

    state.currentEngagements = filtered;

    var nrfRatio = (100 - safeNum(state.overrides.nrfHaircut, 25)) / 100;
    var agg    = aggregateByYear(filtered, nrfRatio);
    var mo     = monthlyTrajectory(filtered, nrfRatio);
    var ranked = accountRanking(filtered);
    var types  = typeBreakdown(filtered);
    var bands  = probabilityBands(filtered);
    var funnel = funnelCascade(filtered, nrfRatio);

    if (!_initialized) {
      renderKPIs(agg);
      renderChartByYear('chart-by-year', agg);
      renderChartTrajectory('chart-trajectory', mo);
      renderChartAccounts('chart-accounts', ranked);
      renderChartTypes('chart-types', types);
      renderChartProb('chart-prob', bands);
      renderChartFunnel('chart-funnel', funnel);
      _initialized = true;
    } else {
      updateKPIs(agg);
      updateChartByYear(agg);
      updateChartTrajectory(mo);
      updateChartAccounts(ranked);
      updateChartTypes(types);
      updateChartProb(bands);
      updateChartFunnel(funnel);
    }

    updateURL(state.overrides);
  }

  function scheduleRecalc() {
    if (_rafPending) cancelAnimationFrame(_rafPending);
    _rafPending = requestAnimationFrame(function () {
      _rafPending = null;
      recalc();
    });
  }

  /* ── URL state ───────────────────────────────────────────────── */
  function parseURLState() {
    var p = new URLSearchParams(location.search);
    return {
      qiddiyaProb:      safeNum(p.get('q'), DEFAULTS.qiddiyaProb),
      retainerConvRate: safeNum(p.get('r'), DEFAULTS.retainerConvRate),
      y3OrigProb:       safeNum(p.get('y'), DEFAULTS.y3OrigProb),
      nrfHaircut:       safeNum(p.get('n'), DEFAULTS.nrfHaircut)
    };
  }

  function updateURL(ov) {
    var p = new URLSearchParams();
    if (ov.qiddiyaProb      !== DEFAULTS.qiddiyaProb)      p.set('q', ov.qiddiyaProb);
    if (ov.retainerConvRate !== DEFAULTS.retainerConvRate) p.set('r', ov.retainerConvRate);
    if (ov.y3OrigProb       !== DEFAULTS.y3OrigProb)       p.set('y', ov.y3OrigProb);
    if (ov.nrfHaircut       !== DEFAULTS.nrfHaircut)       p.set('n', ov.nrfHaircut);
    var qs = p.toString();
    history.replaceState(null, '', location.pathname + (qs ? '?' + qs : ''));
  }

  /* ── Slider helpers ──────────────────────────────────────────── */
  var SLIDER_MAP = {
    'slider-qiddiya':  { key:'qiddiyaProb',      val:'val-qiddiya'  },
    'slider-retainer': { key:'retainerConvRate',  val:'val-retainer' },
    'slider-y3orig':   { key:'y3OrigProb',        val:'val-y3orig'   },
    'slider-nrf':      { key:'nrfHaircut',        val:'val-nrf'      }
  };

  function setSliders(ov) {
    Object.keys(SLIDER_MAP).forEach(function (id) {
      var cfg = SLIDER_MAP[id];
      var el  = document.getElementById(id);
      var vel = document.getElementById(cfg.val);
      if (el)  el.value = ov[cfg.key];
      if (vel) vel.textContent = ov[cfg.key] + '%';
    });
  }

  function wireSliders() {
    Object.keys(SLIDER_MAP).forEach(function (id) {
      var el = document.getElementById(id);
      if (!el) return;
      var cfg = SLIDER_MAP[id];
      el.addEventListener('input', function () {
        var val = safeNum(el.value);
        var vel = document.getElementById(cfg.val);
        if (vel) vel.textContent = val + '%';
        window.appState.overrides[cfg.key] = val;
        scheduleRecalc();
      });
    });
  }

  /* ── Buttons ─────────────────────────────────────────────────── */
  function wireButtons() {
    var resetBtn = document.getElementById('btn-reset');
    var copyBtn  = document.getElementById('btn-copy');

    if (resetBtn) {
      resetBtn.addEventListener('click', function () {
        window.appState.overrides = Object.assign({}, DEFAULTS);
        setSliders(DEFAULTS);
        recalc();
        history.replaceState(null, '', location.pathname);
      });
    }

    if (copyBtn) {
      copyBtn.addEventListener('click', function () {
        var ov = window.appState.overrides;
        var p  = new URLSearchParams({
          q: ov.qiddiyaProb, r: ov.retainerConvRate,
          y: ov.y3OrigProb,  n: ov.nrfHaircut
        });
        var url = location.origin + location.pathname + '?' + p.toString();
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(url).then(function () {
            var orig = copyBtn.textContent;
            copyBtn.textContent = 'Copied ✓';
            setTimeout(function () { copyBtn.textContent = orig; }, 1500);
          }).catch(function () { window.prompt('Copy this URL:', url); });
        } else {
          window.prompt('Copy this URL:', url);
        }
      });
    }
  }

  /* ── Filter chips ────────────────────────────────────────────── */
  function wireChips() {
    document.querySelectorAll('.chip[data-filter]').forEach(function (chip) {
      chip.addEventListener('click', function () {
        var group = chip.dataset.filter;
        var value = chip.dataset.value;
        var siblings = Array.from(document.querySelectorAll('.chip[data-filter="'+group+'"]'));
        var allChip  = document.querySelector('.chip[data-filter="'+group+'"][data-value="all"]');

        if (value === 'all') {
          siblings.forEach(function (s) { s.classList.remove('active'); });
          chip.classList.add('active');
        } else {
          chip.classList.toggle('active');
          if (allChip) allChip.classList.remove('active');
          var anyActive = siblings.some(function (s) {
            return s.dataset.value !== 'all' && s.classList.contains('active');
          });
          if (!anyActive && allChip) allChip.classList.add('active');
        }

        var years = Array.from(document.querySelectorAll('.chip[data-filter="year"].active'))
          .filter(function(c){return c.dataset.value!=='all';})
          .map(function(c){return parseInt(c.dataset.value,10);});
        var types = Array.from(document.querySelectorAll('.chip[data-filter="type"].active'))
          .filter(function(c){return c.dataset.value!=='all';})
          .map(function(c){return c.dataset.value;});

        window.appState.filters = { years: years, types: types };
        scheduleRecalc();
      });
    });
  }

  /* ── Drill-down panel ────────────────────────────────────────── */
  var TYPE_LABELS = { A:'Anchor', S:'Secondary', R:'Retainer', RR:'Renewal', CP:'CP Advisory' };

  function openDrillPanel(accountName) {
    var state = window.appState;
    var eng = (state.currentEngagements || state.originalEngagements || [])
      .filter(function (e) {
        return String(e.account || e.client || e.name || '') === accountName;
      });
    if (!eng.length) return;

    var region  = eng[0].region || eng[0].geo || '';
    var totalWk = eng.reduce(function (s, e) { return s + safeNum(e.weightedK); }, 0);

    var titleEl   = document.getElementById('drill-title');
    var summaryEl = document.getElementById('drill-summary');
    var tbodyEl   = document.getElementById('drill-tbody');

    if (titleEl) titleEl.textContent = accountName;

    if (summaryEl) summaryEl.innerHTML =
      '<div class="kpi-card"><div class="label">Total weighted</div>' +
        '<div class="value" style="font-size:1.5rem">'+fmtM(totalWk)+'</div></div>' +
      '<div class="kpi-card"><div class="label">Engagements</div>' +
        '<div class="value" style="font-size:1.5rem">'+eng.length+'</div>' +
        '<div class="breakdown">'+region+'</div></div>';

    if (tbodyEl) tbodyEl.innerHTML = eng.map(function (e) {
      var p  = safeNum(e.prob); if (p > 1) p = p / 100;
      var gk = safeNum(e._grossK || (p > 0 ? safeNum(e.weightedK) / p : safeNum(e.weightedK)));
      return '<tr>' +
        '<td>'+(TYPE_LABELS[e.type] || e.type || '—')+'</td>' +
        '<td>'+(e.type||'—')+'</td>' +
        '<td>M'+(e.saleMonth||'—')+'</td>' +
        '<td>'+fmtM(gk)+'</td>' +
        '<td>'+Math.round(p*100)+'%</td>' +
        '<td>'+fmtM(safeNum(e.weightedK))+'</td>' +
        '</tr>';
    }).join('');

    var panel    = document.getElementById('drill-panel');
    var backdrop = document.getElementById('drill-backdrop');
    if (panel)    { panel.classList.add('open'); panel.setAttribute('aria-hidden','false'); }
    if (backdrop) backdrop.classList.add('open');
  }

  function closeDrillPanel() {
    var panel    = document.getElementById('drill-panel');
    var backdrop = document.getElementById('drill-backdrop');
    if (panel)    { panel.classList.remove('open'); panel.setAttribute('aria-hidden','true'); }
    if (backdrop) backdrop.classList.remove('open');
  }

  function wireDrillPanel() {
    var closeBtn = document.getElementById('drill-close');
    var backdrop = document.getElementById('drill-backdrop');
    if (closeBtn) closeBtn.addEventListener('click', closeDrillPanel);
    if (backdrop) backdrop.addEventListener('click', closeDrillPanel);
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') closeDrillPanel();
    });
  }

  /* ── main ────────────────────────────────────────────────────── */
  async function main() {
    if (typeof PipelineData === 'undefined') {
      console.warn('[pipeline-render] PipelineData not found — is pipeline-data.js loaded?'); return;
    }
    try { await PipelineData.load(); }
    catch (err) { console.error('[pipeline-render] JSON load failed:', err); return; }

    var data = PipelineData.data;
    if (!data) { console.error('[pipeline-render] PipelineData.data is empty'); return; }

    var urlOverrides = parseURLState();
    window.appState = {
      originalEngagements: data.engagements || [],
      overrides:           urlOverrides,
      filters:             { years: [], types: [] },
      currentEngagements:  data.engagements || [],
      charts:              _charts
    };

    setSliders(urlOverrides);
    wireSliders();
    wireButtons();
    wireChips();
    wireDrillPanel();

    recalc(); /* initial render — _initialized is false, so charts are created */
  }

  document.addEventListener('DOMContentLoaded', main);
})();
