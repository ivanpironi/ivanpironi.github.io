// =============================================================
//  EY Pipeline — Google Apps Script
//  Generated from pipeline-data.json — 2026-06-09
//
//  STEP 1: Open a new Google Sheet
//  STEP 2: Extensions → Apps Script → paste this file
//  STEP 3: Run setupSheet() once  (populates the three tabs)
//  STEP 4: Deploy → New deployment → Web app
//           Execute as: Me  |  Access: Anyone
//  STEP 5: Copy the web app URL → paste in data/pipeline-config.json
//
//  To update numbers: edit the Engagements or Accounts tab.
//  Changes go live on the next page load — no redeploy needed.
//
//  Claude Code write access: POST with token from data/.pipeline-secret
//  Actions: setRow · appendRow · deleteRow · setCell
// =============================================================


/* ── setupSheet ──────────────────────────────────────────────
   Run once to create and populate Meta, Accounts, Engagements.
   Safe to re-run: clears and repopulates every time.
────────────────────────────────────────────────────────────*/
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  /* Meta */
  _tab(ss, 'Meta', [
    ['lastUpdated','retainerConvPct','retainerTcvPct','retainerLagMonths','retainerInitMonths','y3OrigProbPct','nrfHaircutPct'],
    ["2026-06-09", 50, 60, 3, 12, 30, 25]
  ]);

  /* Accounts — one row per account; retainer columns blank if no retainer */
  _tab(ss, 'Accounts', [
    ['key','name','phase','type','region','tip',
     'g_sale','g_saleAmt','g_consPct','g_techPct','g_delStart','g_delEnd',
     'ret_signed','ret_rate','ret_visStart','ret_visEnd','ret_consPct','ret_techPct'],
    ["qiddiya", "Qiddiya", "Phase 1", "ivan", "KSA", "Incumbent PMIS advisor. PIF entertainment giga-programme $40B+. Oracle/SAP/P6/Aconex deployed. Programme Governance & Digital Twin Phase 2.", 1, 700, 60, 40, 2, 7, 9, 35, 10, 33, 60, 40],
    ["sar", "SAR", "Phase 1", "ey", "KSA", "Saudi Arabia Railways. Schedule Health & EPMO design. 18-month bid familiarity. Landbridge $7B PPP.", 2, 275, 65, 35, 3, 6, "", "", "", "", "", ""],
    ["expo", "Expo 2030 Riyadh", "Phase 1", "ey", "KSA", "RFP active. Bechtel PMC. Owner-side assurance above Bechtel. 6km², 42M visitors, Oct 2030 hard deadline. Budget protected from PIF cuts.", 5, 900, 75, 25, 7, 15, 17, 45, 18, 36, 75, 25],
    ["fifa", "FIFA 2034 · WCA34", "Phase 1", "ey", "KSA", "RFP active. 15 stadiums, 5 cities, chaired by Crown Prince MBS. Governance crisis live. Qiddiya & New Murabba contain FIFA 2034 stadiums.", 5, 1150, 80, 20, 7, 15, "", "", "", "", "", ""],
    ["roshn", "ROSHN Group", "Phase 2", "ivan", "KSA", "PIF housing developer. $120B+ portfolio. PMWeb award-winning but ERP integration gap structural at multi-region scale.", 8, 700, 55, 45, 10, 15, 17, 35, 18, 36, 55, 45],
    ["kspf", "King Salman Park", "Phase 2", "ivan", "KSA", "$38B · 17.2km² · 2027 opening. Seven warm contacts. Parsons PMO. Digital Twin & iCCC mandates open.", 9, 700, 50, 50, 11, 18, 20, 35, 21, 36, 50, 50],
    ["diriyah", "Diriyah (DGDA)", "Phase 2", "ey", "KSA", "$63B UNESCO heritage programme. Mace Consult sitewide PMC. $1.5B Arena Block awarded June 2025.", 13, 700, 70, 30, 14, 21, 23, 35, 24, 36, 70, 30],
    ["ksia", "King Salman Airport", "Phase 2", "ey", "KSA", "$30B · 57km² · 3 PMCs simultaneously: Mace, Bechtel, Parsons. Hard deadlines: Expo 2030 + FIFA 2034.", 13, 1100, 55, 45, 15, 24, "", "", "", "", "", ""],
    ["aramco", "Saudi Aramco", "Phase 3", "ey", "KSA", "$40B+ annual capex. Worley GES+ contract. SAP↔PMIS cost reconciliation structural problem. Highest single fee in portfolio.", 15, 900, 35, 65, 18, 26, "", "", "", "", "", ""],
    ["newmurabba", "New Murabba", "Phase 3", "ivan", "KSA", "$38B · 3 PMCs (Bechtel/Turner/AECOM). Contains FIFA 2034 stadium. Naver Cloud MoU.", 16, 700, 55, 45, 18, 24, "", "", "", "", "", ""],
    ["khazna", "Khazna Data Centers", "Phase 2", "ivan", "UAE", "30+ facilities. $26.2B financing. Stargate UAE partner. NexOps/Presight AI launched Feb 2026.", 18, 450, 50, 50, 19, 24, "", "", "", "", "", ""],
    ["rsg", "Red Sea Global", "Phase 3", "ivan", "KSA", "IOC Operating Model delivered — warm relationships. PIF spending cuts heaviest here. Phase 1 resorts operational.", 20, 550, 70, 30, 22, 26, "", "", "", "", "", ""],
    ["taqa", "TAQA", "Phase 3", "ivan", "UAE", "Abu Dhabi integrated utilities. AED 75B investment pipeline to 2030. ~27GW new capacity. Multi-segment capital portfolio across generation, water, T&D.", 25, 750, 40, 60, 27, 33, "", "", "", "", "", ""],
    ["admd", "Abu Dhabi Master Developers", "Phase 3", "ey", "UAE", "Aldar · Miral · MODON. EY audit relationships across UAE master developer ecosystem. Multi-asset Oracle PMIS governance.", 26, 700, 55, 45, 28, 36, "", "", "", "", "", ""],
    ["tasmu", "TASMU", "Phase 3", "ivan", "Qatar", "Qatar Smart City Programme. Incumbent delivery context — post-deployment expansion and iCCC advisory re-engagement.", 27, 550, 50, 50, 29, 36, "", "", "", "", "", ""]
  ]);

  /* Engagements — one row per engagement; edit prob / grossK to update model */
  _tab(ss, 'Engagements', [
    ['account_key','type','saleMonth','duration','grossK','prob','isY3Origination'],
    ["qiddiya", "A", 3, 18, 1400, 0.5, false],
    ["qiddiya", "A", 16, 12, 700, 0.5, false],
    ["qiddiya", "S", 6, 12, 500, 0.3, false],
    ["qiddiya", "S", 20, 9, 333, 0.3, false],
    ["qiddiya", "R", 13, 24, 500, 0.5, false],
    ["qiddiya", "RR", 25, 12, 100, 0.5, false],
    ["sar", "CP", 4, 4, 60, 0.75, false],
    ["sar", "A", 6, 9, 350, 0.45, false],
    ["sar", "S", 10, 6, 200, 0.3, false],
    ["sar", "R", 13, 24, 100, 0.5, false],
    ["expo", "CP", 5, 6, 80, 0.8, false],
    ["expo", "A", 8, 12, 900, 0.5, false],
    ["expo", "S", 12, 9, 400, 0.3, false],
    ["expo", "R", 15, 24, 160, 0.5, false],
    ["expo", "RR", 27, 12, 120, 0.5, false],
    ["fifa", "A", 10, 12, 1200, 0.5, false],
    ["fifa", "A", 23, 15, 900, 0.4, false],
    ["fifa", "S", 14, 9, 400, 0.3, false],
    ["fifa", "CP", 8, 6, 100, 0.8, false],
    ["fifa", "R", 18, 12, 160, 0.5, false],
    ["roshn", "CP", 17, 6, 50, 0.8, false],
    ["roshn", "A", 19, 12, 700, 0.35, false],
    ["roshn", "R", 22, 24, 100, 0.4, false],
    ["roshn", "S", 23, 9, 300, 0.25, false],
    ["kspf", "CP", 13, 6, 71, 0.75, false],
    ["kspf", "A", 15, 12, 900, 0.4, false],
    ["kspf", "S", 20, 9, 400, 0.25, false],
    ["kspf", "R", 19, 24, 140, 0.5, false],
    ["kspf", "RR", 31, 12, 100, 0.5, false],
    ["diriyah", "A", 13, 12, 1000, 0.4, false],
    ["diriyah", "A", 26, 12, 550, 0.35, true],
    ["diriyah", "S", 16, 9, 400, 0.3, false],
    ["diriyah", "S", 28, 6, 200, 0.25, false],
    ["diriyah", "R", 20, 24, 140, 0.5, false],
    ["ksia", "CP", 15, 6, 71, 0.7, false],
    ["ksia", "A", 17, 12, 1000, 0.35, false],
    ["ksia", "R", 21, 24, 160, 0.5, false],
    ["ksia", "S", 22, 9, 400, 0.25, false],
    ["ksia", "A", 30, 15, 700, 0.3, true],
    ["ksia", "RR", 33, 12, 120, 0.5, false],
    ["aramco", "CP", 26, 6, 100, 0.6, false],
    ["aramco", "A", 25, 12, 1200, 0.3, true],
    ["aramco", "R", 28, 24, 160, 0.3, false],
    ["aramco", "S", 29, 9, 500, 0.2, false],
    ["newmurabba", "CP", 25, 6, 100, 0.7, false],
    ["newmurabba", "A", 27, 12, 700, 0.4, true],
    ["newmurabba", "S", 29, 9, 200, 0.25, false],
    ["newmurabba", "A", 32, 12, 550, 0.35, true],
    ["khazna", "CP", 13, 6, 71, 0.7, false],
    ["khazna", "A", 14, 12, 600, 0.45, false],
    ["khazna", "R", 16, 24, 100, 0.5, false],
    ["khazna", "S", 18, 9, 280, 0.25, false],
    ["khazna", "A", 22, 12, 400, 0.35, false],
    ["rsg", "CP", 27, 6, 71, 0.6, false],
    ["rsg", "A", 26, 12, 700, 0.35, true],
    ["rsg", "R", 29, 24, 120, 0.35, false],
    ["rsg", "S", 30, 9, 300, 0.25, false],
    ["taqa", "CP", 26, 6, 71, 0.65, false],
    ["taqa", "A", 27, 18, 800, 0.4, true],
    ["taqa", "R", 28, 24, 100, 0.45, false],
    ["taqa", "S", 29, 12, 350, 0.3, false],
    ["taqa", "A", 33, 12, 500, 0.35, true],
    ["taqa", "RR", 34, 12, 80, 0.4, false],
    ["admd", "CP", 27, 6, 71, 0.7, false],
    ["admd", "CP", 30, 6, 57, 0.7, false],
    ["admd", "A", 26, 15, 1000, 0.4, true],
    ["admd", "S", 28, 9, 400, 0.3, false],
    ["admd", "R", 27, 24, 160, 0.5, false],
    ["admd", "A", 32, 12, 700, 0.35, true],
    ["admd", "S", 33, 9, 300, 0.25, false],
    ["admd", "RR", 34, 24, 140, 0.45, false],
    ["tasmu", "CP", 26, 6, 57, 0.65, false],
    ["tasmu", "A", 25, 18, 700, 0.45, true],
    ["tasmu", "R", 28, 24, 120, 0.45, false],
    ["tasmu", "S", 27, 12, 350, 0.3, false],
    ["tasmu", "A", 32, 12, 500, 0.4, true],
    ["tasmu", "S", 33, 9, 250, 0.25, false],
    ["tasmu", "RR", 34, 12, 80, 0.4, false]
  ]);

  SpreadsheetApp.flush();
  Logger.log('✓ Setup complete — 3 tabs created.');
}

function _tab(ss, name, data) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  sh.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff');
}


/* ── doGet ───────────────────────────────────────────────────
   Web app entry point. Returns pipeline JSON.
   Deploy: Execute as Me · Access: Anyone
────────────────────────────────────────────────────────────*/
function doGet(e) {
  try {
    var json = _buildJSON(SpreadsheetApp.getActiveSpreadsheet());
    return ContentService
      .createTextOutput(JSON.stringify(json))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/* ── doPost ──────────────────────────────────────────────────
   Write endpoint for Claude Code. Protected by shared secret.
   Body (JSON): { token, action, tab, ...params }
   Actions:
     setCell   — { row, col, value }           (1-indexed)
     setRow    — { match: {col:val}, data: {col:val} }
     appendRow — { data: {col:val} }
     deleteRow — { match: {col:val} }
────────────────────────────────────────────────────────────*/
function doPost(e) {
  try {
    var SECRET = 'ey-pipeline-2026-ivan';
    var body = JSON.parse(e.postData.contents);
    if (body.token !== SECRET) {
      return ContentService
        .createTextOutput(JSON.stringify({error: 'unauthorized'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = _handleWrite(ss, body);
    return ContentService
      .createTextOutput(JSON.stringify({ok: true, result: result}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function _handleWrite(ss, body) {
  var sh = ss.getSheetByName(body.tab);
  if (!sh) throw new Error('Tab not found: ' + body.tab);
  var action = body.action;

  if (action === 'setCell') {
    sh.getRange(body.row, body.col).setValue(body.value);
    return 'setCell OK r' + body.row + ' c' + body.col;
  }

  if (action === 'setRow') {
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0];
    for (var i = 1; i < vals.length; i++) {
      var matched = Object.keys(body.match).every(function(k) {
        var ci = hdrs.indexOf(k);
        return ci >= 0 && String(vals[i][ci]) === String(body.match[k]);
      });
      if (matched) {
        Object.keys(body.data).forEach(function(k) {
          var ci = hdrs.indexOf(k);
          if (ci >= 0) sh.getRange(i + 1, ci + 1).setValue(body.data[k]);
        });
        return 'setRow OK row ' + (i + 1);
      }
    }
    throw new Error('No matching row found');
  }

  if (action === 'appendRow') {
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var newRow = hdrs.map(function(h) {
      return body.data[h] !== undefined ? body.data[h] : '';
    });
    sh.appendRow(newRow);
    return 'appendRow OK';
  }

  if (action === 'deleteRow') {
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0];
    for (var i = vals.length - 1; i >= 1; i--) {
      var matched = Object.keys(body.match).every(function(k) {
        var ci = hdrs.indexOf(k);
        return ci >= 0 && String(vals[i][ci]) === String(body.match[k]);
      });
      if (matched) {
        sh.deleteRow(i + 1);
        return 'deleteRow OK row ' + (i + 1);
      }
    }
    throw new Error('No matching row found');
  }

  throw new Error('Unknown action: ' + action);
}


/* ── _buildJSON ──────────────────────────────────────────────
   Reads three tabs → assembles pipeline-data.json structure.
   All derived fields (consAmt, techAmt, cpm, tpm …) are
   computed here so the sheet stays clean and editable.
────────────────────────────────────────────────────────────*/
function _buildJSON(ss) {

  function rows(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    var vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var hdrs = vals[0];
    return vals.slice(1).map(function(r) {
      var o = {};
      hdrs.forEach(function(h, i) { o[h] = r[i]; });
      return o;
    });
  }

  /* ── Meta ── */
  var m = rows('Meta')[0];
  var meta = {
    lastUpdated: String(m.lastUpdated),
    accounts: 15,
    assumptions: {
      retainerConvPct:    Number(m.retainerConvPct),
      retainerTcvPct:     Number(m.retainerTcvPct),
      retainerLagMonths:  Number(m.retainerLagMonths),
      retainerInitMonths: Number(m.retainerInitMonths),
      y3OrigProbPct:      Number(m.y3OrigProbPct),
      nrfHaircutPct:      Number(m.nrfHaircutPct)
    }
  };

  /* ── Engagements (grouped by account key) ── */
  var engByKey = {};
  rows('Engagements').forEach(function(e) {
    var k = e.account_key;
    if (!engByKey[k]) engByKey[k] = [];
    var eng = {
      type:      String(e.type),
      saleMonth: Number(e.saleMonth),
      duration:  Number(e.duration),
      grossK:    Number(e.grossK),
      prob:      Number(e.prob)
    };
    if (e.isY3Origination === true || e.isY3Origination === 'TRUE')
      eng.isY3Origination = true;
    engByKey[k].push(eng);
  });

  /* ── Accounts ── */
  var accounts = rows('Accounts').map(function(a) {
    var saleAmt = Number(a.g_saleAmt);
    var cPct    = Number(a.g_consPct);
    var dS      = Number(a.g_delStart);
    var dE      = Number(a.g_delEnd);
    var consAmt = Math.round(saleAmt * cPct / 100);

    var gantt = {
      sale:     Number(a.g_sale),
      saleAmt:  saleAmt,
      consPct:  cPct,
      techPct:  Number(a.g_techPct),
      consAmt:  consAmt,
      techAmt:  saleAmt - consAmt,
      delStart: dS,
      delEnd:   dE,
      dur:      dE - dS + 1
    };

    var rSigned = a.ret_signed;
    if (rSigned !== '' && rSigned !== null && Number(rSigned) > 0) {
      var rRate  = Number(a.ret_rate);
      var rCPct  = Number(a.ret_consPct);
      var rVisS  = Number(a.ret_visStart);
      var rVisE  = Number(a.ret_visEnd);
      var months = rVisE - rVisS + 1;
      var cpm    = Math.round(rRate * rCPct / 100);
      var tpm    = rRate - cpm;
      gantt.ret = {
        signed:   Number(rSigned),
        rate:     rRate,
        visStart: rVisS,
        visEnd:   rVisE,
        consPct:  rCPct,
        techPct:  Number(a.ret_techPct),
        cpm:      cpm,
        tpm:      tpm,
        saleAmt:  rRate * months,
        consSale: cpm * months,
        techSale: tpm * months
      };
    }

    return {
      key:         String(a.key),
      name:        String(a.name),
      phase:       String(a.phase),
      type:        String(a.type),
      region:      String(a.region),
      tip:         String(a.tip),
      gantt:       gantt,
      engagements: engByKey[String(a.key)] || []
    };
  });

  return { meta: meta, accounts: accounts };
}
