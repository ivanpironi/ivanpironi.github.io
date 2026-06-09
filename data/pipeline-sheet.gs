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
    ['key','name','phase','type','region','tip','overview','contacts',
     'g_sale','g_saleAmt','g_consPct','g_techPct','g_delStart','g_delEnd',
     'ret_signed','ret_rate','ret_visStart','ret_visEnd','ret_consPct','ret_techPct'],
    ["qiddiya", "Qiddiya", "Phase 1", "ivan", "KSA",
     "Incumbent PMIS advisor. PIF entertainment giga-programme $40B+. Oracle/SAP/P6/Aconex deployed. Programme Governance & Digital Twin Phase 2.",
     "PIF-owned entertainment and tourism giga-programme. Phase 3 city establishment now mobilising — 11,000 housing units, full entertainment city. Oracle Unifier + SAP + P6 + Aconex deployed. September 2025: Yardi adopted for real estate operations — creating a three-way integration challenge between Yardi, Oracle Unifier, and SAP that no current advisor can address without a conflict of interest. Ivan designed the Oracle layer. Ivan is the incumbent PMIS advisor.",
     "Manuel Chaure — Executive Director, Digital & IT Strategy & Governance | Abdullah Barakat — Executive Director, IT Business Relationship Management | Sreedhar Reddy Kadiveti — Senior Director, Engineering & Development Business Solutions",
     1, 700, 60, 40, 2, 7, 9, 35, 10, 33, 60, 40],
    ["sar", "SAR", "Phase 1", "ey", "KSA",
     "Saudi Arabia Railways. Schedule Health & EPMO design. 18-month bid familiarity. Landbridge $7B PPP.",
     "National rail operator managing Landbridge ($7B PPP/BOT), the Qiddiya rail link (PPP concession), and 8,000km+ national network expansion simultaneously. Oracle Analytics Cloud deployed for finance and procurement — capital project delivery data remains EPC-managed. 18-month PMIS/EPMO bid engagement provides direct familiarity with SAR's governance structure, decision-makers, and the specific data trust gaps the advisory addresses. The EPMO mandate is confirmed open.",
     "Dr. Mishari Al Mishari — Deputy Director, NIC | Mohaimeed Bagabas — EPMO Manager | Mohammed AlShaia — Director, Strategic Management Office | Fahad Almutairi — SAR (role TBC) | Salamh Alenazi — SAR (role TBC)",
     2, 275, 65, 35, 3, 6, "", "", "", "", "", ""],
    ["expo", "Expo 2030 Riyadh", "Phase 1", "ey", "KSA",
     "RFP active. Bechtel PMC. Owner-side assurance above Bechtel. 6km², 42M visitors, Oct 2030 hard deadline. Budget protected from PIF cuts.",
     "PIF launched Expo 2030 Riyadh Company (ERC) in June 2025. CEO: Talal H. AlMarri. Bechtel appointed as PMC (John Doyle, Programme Director). 6km² site adjacent to King Salman International Airport. 226 pavilions, 246 exhibiting nations, 42M visitors, October 2030 hard opening deadline. Budget explicitly protected from PIF spending cuts. No advisory firm confirmed as independent programme governance lead above Bechtel. Advisory RFP expected imminently.",
     "Talal H. AlMarri — CEO, Expo 2030 Riyadh Company | Ibrahim Al Sultan — Chairman, Expo 2030 · Minister of State | John Doyle — Programme Director, Expo 2030 Riyadh (Bechtel PMC) | Investigate: EY MENA relationships with PIF and ERC board members",
     5, 900, 75, 25, 7, 15, 17, 45, 18, 36, 75, 25],
    ["fifa", "FIFA 2034 · WCA34", "Phase 1", "ey", "KSA",
     "RFP active. 15 stadiums, 5 cities, chaired by Crown Prince MBS. Governance crisis live. Qiddiya & New Murabba contain FIFA 2034 stadiums.",
     "World Cup 2034 Hosting Higher Authority (WCA34) — chaired by Crown Prince MBS. Board includes Ministers of Finance, Transport, Sport, Tourism, Health, Communications, and the PIF Governor. 15 stadiums (11 new builds, 4 renovations) across Riyadh, Jeddah, Al Khobar, Abha, and NEOM. Stadium costs being reassessed December 2025 — construction starts delayed. The governance crisis creating the advisory mandate is live now. Qiddiya and New Murabba both contain FIFA 2034 stadiums.",
     "World Cup 2034 Hosting Higher Authority (WCA34) | HRH Crown Prince MBS — Chairman, Board of Directors | Yasir Al-Rumayyan — Governor, PIF · Board member | Prince Abdulaziz bin Turki — Minister of Sport · Board member | Investigate: EY government advisory relationships + Mace global stadium credentials",
     5, 1150, 80, 20, 7, 15, "", "", "", "", "", ""],
    ["roshn", "ROSHN Group", "Phase 2", "ivan", "KSA",
     "PIF housing developer. $120B+ portfolio. PMWeb award-winning but ERP integration gap structural at multi-region scale.",
     "PIF-owned national housing developer scaling from 2 to 7+ active regions simultaneously. PMWeb deployed across the EPMO — Project of the Year 2025. The governance question at multi-region scale: is PMWeb data trustworthy across all regions? SAP vs PMWeb cost reconciliation gap structural at portfolio scale. PIF spending cut pressure (20%+ mandated 2025) requires defensible ROI data for PIF leadership reporting.",
     "Hazem Alaily — Chief Program & Project Management Officer | Hadi E. — Executive Director, EPMO | Roger Fatovic — Group Chief Projects Delivery Officer (Acting) | Mohammad Ajlouni — Director, EPMO & Project Controls",
     8, 700, 55, 45, 10, 15, 17, 35, 18, 36, 55, 45],
    ["kspf", "King Salman Park", "Phase 2", "ivan", "KSA",
     "$38B · 17.2km² · 2027 opening. Seven warm contacts. Parsons PMO. Digital Twin & iCCC mandates open.",
     "World's largest urban park — 17.2km², former Riyadh air base, $38B investment, 2027 opening deadline. 13,000 workers on site. Parsons holds the 25-month PMO contract. CEO publicly confirmed Digital Twin in active development. KSPF leadership receives all programme intelligence through Parsons. No independent owner-side view of whether Parsons' reporting reflects delivery reality. Seven named contacts from a delivered Smart City Strategy engagement.",
     "Majed Alhudaib — CTO | Dale Chadwick — Chief Development Officer | Tom Jarvis — Head of Park Operations | Douglas Stagner — Operational Readiness Director | Ben Edmonds — Smart City Director | Abdulmalik Alsaleh — Chief Data Officer",
     9, 700, 50, 50, 11, 18, 20, 35, 21, 36, 50, 50],
    ["diriyah", "Diriyah (DGDA)", "Phase 2", "ey", "KSA",
     "$63B UNESCO heritage programme. Mace Consult sitewide PMC. $1.5B Arena Block awarded June 2025.",
     "$63B UNESCO World Heritage Site being developed into cultural and hospitality destination. Mace Consult holds the sitewide PMC mandate. $1.5B Arena Block awarded June 2025 — construction accelerating. Heritage asset information management overlays traditional capital delivery governance in ways no standard PMC addresses. EY advises DGDA at the owner-side layer above Mace.",
     "Entry via: Christopher Seymour — Managing Director MEA, Mace Consult | Antony Brania — Director, Mace Consult",
     13, 700, 70, 30, 14, 21, 23, 35, 24, 36, 70, 30],
    ["ksia", "King Salman Airport", "Phase 2", "ey", "KSA",
     "$30B · 57km² · 3 PMCs simultaneously: Mace, Bechtel, Parsons. Hard deadlines: Expo 2030 + FIFA 2034.",
     "Largest airport infrastructure programme in the world. Three major delivery firms simultaneously: Mace Consult (Delivery Partner — Christopher Seymour signed personally), Bechtel (three passenger terminals), Parsons (runways, taxiways, ATC towers). Three firms, three reporting chains, one owner who needs to trust the data. Hard deadlines: Expo 2030 operational opening, FIFA 2034 tournament. Construction commenced September 2025.",
     "Marco Mejia — Acting CEO, KSIADC | Linda Schucroft — VP Innovation & Guest Experience | Entry via: Christopher Seymour — Managing Director MEA, Mace Consult (signed personally)",
     13, 1100, 55, 45, 15, 24, "", "", "", "", "", ""],
    ["aramco", "Saudi Aramco", "Phase 3", "ey", "KSA",
     "$40B+ annual capex. Worley GES+ contract. SAP vs PMIS cost reconciliation structural problem. Highest single fee in portfolio.",
     "Largest capital spender in KSA — $40B+ annual capex. Worley holds the GES+ 5-year engineering and project management contract. SAP deeply embedded as ERP across the capital portfolio. Aramco cut $40B dividends in 2025 due to lower oil prices — capital efficiency under the highest scrutiny in recent years. SAP vs PMIS cost reconciliation confirmed as a structural failure across the capital portfolio. Ivan has resolved this exact problem at Qiddiya scale. Highest single-engagement fee potential in the GCC market.",
     "Entry via: Alex Silvermaster — Plan Practice Lead EMEA, Worley (+44 7585 984 918) | Worley GES+ contract provides programme-level access to Aramco capital leadership",
     15, 900, 35, 65, 18, 26, "", "", "", "", "", ""],
    ["newmurabba", "New Murabba", "Phase 3", "ivan", "KSA",
     "$38B · 3 PMCs (Bechtel/Turner/AECOM). Contains FIFA 2034 stadium. Naver Cloud MoU.",
     "$38B mixed-use district in central Riyadh including The Mukaab (one of the world's largest built structures). Three PMC firms simultaneously: Bechtel (masterplan + infrastructure), Turner Arabia (The Mukaab), AECOM (buildings). No single owner-side PMIS layer above three separate reporting chains. Naver Cloud MoU confirmed for smart city monitoring. FIFA 2034 New Murabba Stadium (46,010 capacity) on site.",
     "Mohammed Alqahtani — Head of Smart City & AI (MIT) | Wail Balkhair — Executive Director, Asset Management | Faisal AlSayegh — Manager, Smart City & AI | Abdullah Bafakih — Development Director | Michael Dyke — CEO",
     16, 700, 55, 45, 18, 24, "", "", "", "", "", ""],
    ["khazna", "Khazna Data Centers", "Phase 2", "ivan", "UAE",
     "30+ facilities. $26.2B financing. Stargate UAE partner. NexOps/Presight AI launched Feb 2026.",
     "UAE data centre platform — 30+ facilities, $26.2B financing secured, Stargate UAE partner (February 2026), G42/Mubadala ecosystem. Turner & Townsend holds PMC role. February 2026: Khazna launched NexOps — in-house operations model with Presight AI providing AI-powered command and control across 30+ facilities. NexOps insourced 230+ specialists, 5,000 operational documents built in under 12 months. The data governance gap between construction delivery and NexOps operational requirements is completely unaddressed.",
     "Abdulmajeed Harmoodi — CTO | Kevin Ayling — Senior Director, Technical Program Management | Hassan Al Naqbi — CEO",
     18, 450, 50, 50, 19, 24, "", "", "", "", "", ""],
    ["rsg", "Red Sea Global", "Phase 3", "ivan", "KSA",
     "IOC Operating Model delivered — warm relationships. PIF spending cuts heaviest here. Phase 1 resorts operational.",
     "PIF-owned luxury tourism giga-programme — 100+ operators, 10+ integrated city functions, geographic footprint comparable to Belgium. Among the PIF projects bearing the heaviest 2025 spending cut pressure. Phase 1 resorts are operational while Phase 2+ construction remains active. IOC Operating Model previously delivered — warm relationships with RSG operations leadership already established.",
     "Via team delivery relationships with RSG operations leadership | Established through IOC Operating Model engagement | Warm account — no cold approach required",
     20, 550, 70, 30, 22, 26, "", "", "", "", "", ""],
    ["taqa", "TAQA", "Phase 3", "ivan", "UAE",
     "Abu Dhabi integrated utilities. AED 75B investment pipeline to 2030. ~27GW new capacity. Multi-segment capital portfolio across generation, water, T&D.",
     "Abu Dhabi's integrated utilities champion and one of EMEA's largest listed utilities (~AED 200bn total assets). Committed to an AED 75bn investment pipeline to 2030, including ~AED 40bn on UAE transmission & distribution networks, plus major generation and water capital — Taweelah (world's largest RO desalination), Mirfa 2, Shuweihat 4 — and ~27GW of new capacity. Capex is accelerating sharply. A multi-segment regulated capital portfolio (generation · water · T&D) of exactly the scale where owner-side PMIS governance, capital controls, and energy/water asset-performance intelligence are in demand — with no single owner-side programme-controls layer across the segment portfolio.",
     "Entry via: Ivan's network — internal relationships from time in TAQA's Customer Centre business unit. Warm organisational entry; bridge toward T&D and generation capital-programme leadership | Programme-side contacts — to be confirmed",
     25, 750, 40, 60, 27, 33, "", "", "", "", "", ""],
    ["admd", "Abu Dhabi Master Developers", "Phase 3", "ey", "UAE",
     "Aldar · Miral · MODON. EY audit relationships across UAE master developer ecosystem. Multi-asset Oracle PMIS governance.",
     "Three of Abu Dhabi's largest master developers — Aldar Properties, Miral, and MODON — each operating large-scale capital programmes across residential, tourism, and industrial sectors. EY holds existing audit and advisory relationships across the UAE developer ecosystem. Multi-asset portfolio scale creates exactly the class of PMIS governance, capital controls, and Digital Twin lifecycle challenges this practice is built to address. Combined capital deployment creates programme oversight complexity that no single developer has yet resolved at portfolio level.",
     "Entry via: EY existing audit and advisory relationships · UAE | Programme-side contacts — to be confirmed",
     26, 700, 55, 45, 28, 36, "", "", "", "", "", ""],
    ["tasmu", "TASMU", "Phase 3", "ivan", "Qatar",
     "Qatar Smart City Programme. Incumbent delivery context — post-deployment expansion and iCCC advisory re-engagement.",
     "Qatar's national smart city programme — the integrated command and control platform built for the FIFA 2022 World Cup and now transitioning to long-term smart city operations. Ivan led the iCCC design and delivery. The advisory opportunity lies in post-deployment platform evolution: extending the iCCC into new operational domains, integrating with ongoing capital delivery programmes, and building the digital twin operational intelligence layer that the original build mandate did not include. Incumbent delivery context — the relationship exists; this is re-engagement, not cold outreach.",
     "Entry via: Incumbent delivery context · direct programme relationship · no introduction required | Programme-side contacts — to be confirmed",
     27, 550, 50, 50, 29, 36, "", "", "", "", "", ""]
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
      overview:    String(a.overview),
      contacts:    String(a.contacts),
      gantt:       gantt,
      engagements: engByKey[String(a.key)] || []
    };
  });

  return { meta: meta, accounts: accounts };
}
