// ═══════════════════════════════════════════════════════
// SECTOR ROTATION TERMINAL v3 — Google Apps Script Backend
// ═══════════════════════════════════════════════════════
// WHAT'S NEW IN V3:
//   • SECTOR MAP extended to col O (rotation, breadth, verdict)
//   • EMERGING extended to col O (percentile, quadrant, rotation, note)
//   • ROTATION now parses all 3 sections (risk pairs, sector vs SPY, credit)
//   • CONVICTION label parsed as text (not number)
//   • Drill-down uses INDUSTRY SCANNER PARENT column (explicit mapping)
//   • Server-side regime verdict for snapshot consistency
//
// SETUP:
//   1. Paste this into Extensions → Apps Script → Code.gs
//   2. Run snapshotSignals() manually once to create _SNAPSHOTS tab
//   3. Add a daily trigger: Edit → Triggers → snapshotSignals
//      - Event source: Time-driven
//      - Type: Day timer → 4pm to 5pm (after market close)
// ═══════════════════════════════════════════════════════
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('◈ Sector Rotation Terminal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// ═══════════════════════════════════════════════════════
// MAIN DATA ENDPOINT
// ═══════════════════════════════════════════════════════
function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};
  try { result.regime    = parseRegime(ss); }      catch(e) { result.regime    = {signals:[]}; }
  try { result.sectorMap = parseSectorMap(ss); }    catch(e) { result.sectorMap = {context:'',sectors:[]}; }
  try { result.universe  = parseUniverse(ss); }     catch(e) { result.universe  = []; }
  try { result.rotation  = parseRotation(ss); }     catch(e) { result.rotation  = {pairs:[],sectorSignals:[],creditHealth:[],creditVerdict:'',movers:[]}; }
  try { result.emerging  = parseEmerging(ss); }      catch(e) { result.emerging  = []; }
  try { result.ratios    = parseRatios(ss); }        catch(e) { result.ratios    = {}; }
  try { result.credit    = parseCredit(ss); }        catch(e) { result.credit    = {spreads:[]}; }
  try { result.equalWeight = parseEqualWeight(ss); } catch(e) { result.equalWeight = {gauges:[],sectorBreadth:[]}; }
  // ── Conviction picks (formula-driven sheet) ──
  try { result.conviction = parseConviction(ss); } catch(e) { result.conviction = {top:[],bottom:[]}; }
  // ── Drill-down groups (INDUSTRY SCANNER with PARENT mapping) ──
  try { result.drillDown = buildDrillDown(ss, result.sectorMap); } catch(e) { result.drillDown = {}; }
  // ── Signal changes vs last snapshot ──
  try { result.changes = computeChanges(ss, result); } catch(e) { result.changes = {items:[],hasPrevious:false}; }
  result.ts = new Date().toISOString();
  return JSON.stringify(result);
}
// ═══════════════════════════════════════════════════════
// SNAPSHOT SYSTEM
// ═══════════════════════════════════════════════════════
function computeVerdict(signals) {
  if (!signals || !signals.length) return 'LOADING';
  var bull = 0, bear = 0;
  for (var i = 0; i < signals.length; i++) {
    var v = String(signals[i].verdict || '').toUpperCase();
    if (v.indexOf('BULL') !== -1 || v.indexOf('RISK-ON') !== -1 || v.indexOf('CALM') !== -1 || v.indexOf('GLOBAL ON') !== -1) bull++;
    if (v.indexOf('BEAR') !== -1 || v.indexOf('RISK-OFF') !== -1 || v.indexOf('FEAR') !== -1) bear++;
  }
  if (bull >= 4) return 'RISK-ON';
  if (bull >= 3) return 'CONSTRUCTIVE';
  if (bear >= 3) return 'CAUTIOUS';
  if (bear >= 4) return 'RISK-OFF';
  return 'MIXED';
}
function buildSnapshot(data) {
  return {
    regimeVerdict: computeVerdict(data.regime.signals),
    signals: (data.regime.signals || []).map(function(s) {
      return { signal: String(s.signal), verdict: String(s.verdict) };
    }),
    sectors: (data.sectorMap.sectors || []).map(function(s) {
      return { ticker: String(s.ticker), name: String(s.name), quadrant: String(s.quadrant), rank: s.rank, conviction: s.conviction };
    }),
    top10: (((data.conviction && data.conviction.top && data.conviction.top.length) ? data.conviction.top : data.universe) || []).slice(0, 10).map(function(u) {
      return String(u.ticker);
    }),
    top25: (((data.conviction && data.conviction.top && data.conviction.top.length) ? data.conviction.top : data.universe) || []).slice(0, 25).map(function(u) {
      return { ticker: String(u.ticker), rank: u.rank };
    }),
    rotations: (data.rotation.pairs || []).map(function(p) {
      return { pair: String(p.pair), signal: String(p.signal) };
    })
  };
}
function snapshotSignals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {};
  try { data.regime    = parseRegime(ss); }      catch(e) { data.regime    = {signals:[]}; }
  try { data.sectorMap = parseSectorMap(ss); }    catch(e) { data.sectorMap = {context:'',sectors:[]}; }
  try { data.universe  = parseUniverse(ss); }     catch(e) { data.universe  = []; }
  try { data.rotation  = parseRotation(ss); }     catch(e) { data.rotation  = {pairs:[],sectorSignals:[],creditHealth:[],creditVerdict:'',movers:[]}; }
  try { data.conviction = parseConviction(ss); }  catch(e) { data.conviction = {top:[],bottom:[]}; }
  var snapshot = buildSnapshot(data);
  var sheet = ss.getSheetByName('_SNAPSHOTS');
  if (!sheet) {
    sheet = ss.insertSheet('_SNAPSHOTS');
    sheet.getRange('A1:D1').setValues([['DATE', 'TIMESTAMP', 'REGIME', 'STATE_JSON']]);
    sheet.setColumnWidth(4, 600);
    sheet.getRange('A1:D1').setFontWeight('bold');
    sheet.hideSheet();
  }
  var today = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
  var ts = new Date().toISOString();
  var jsonStr = JSON.stringify(snapshot);
  var rows = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < rows.length; i++) {
    var rowDate = '';
    try { rowDate = Utilities.formatDate(new Date(rows[i][0]), 'America/New_York', 'yyyy-MM-dd'); } catch(e) { rowDate = String(rows[i][0]); }
    if (rowDate === today) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([[today, ts, snapshot.regimeVerdict, jsonStr]]);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([today, ts, snapshot.regimeVerdict, jsonStr]);
  }
  return 'Snapshot saved: ' + today + ' — ' + snapshot.regimeVerdict;
}
function computeChanges(ss, currentData) {
  var sheet = ss.getSheetByName('_SNAPSHOTS');
  if (!sheet) return { items: [], hasPrevious: false, prevDate: '' };
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { items: [], hasPrevious: false, prevDate: '' };
  var today = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
  var prevRow = null;
  for (var i = rows.length - 1; i >= 1; i--) {
    var rowDate = '';
    try { rowDate = Utilities.formatDate(new Date(rows[i][0]), 'America/New_York', 'yyyy-MM-dd'); } catch(e) { rowDate = String(rows[i][0]); }
    if (rowDate !== today) {
      prevRow = rows[i];
      break;
    }
  }
  if (!prevRow) return { items: [], hasPrevious: false, prevDate: '' };
  var prev;
  try { prev = JSON.parse(prevRow[3]); } catch(e) { return { items: [], hasPrevious: false, prevDate: '' }; }
  var prevDate = '';
  try { prevDate = Utilities.formatDate(new Date(prevRow[0]), 'America/New_York', 'MMM d'); } catch(e) { prevDate = String(prevRow[0]); }
  var now = buildSnapshot(currentData);
  var changes = [];
  // 1. Regime verdict flip
  if (now.regimeVerdict !== prev.regimeVerdict) {
    changes.push({
      type: 'REGIME', severity: 'high',
      icon: '◈', color: '#ff8800',
      msg: 'Regime shifted: ' + prev.regimeVerdict + ' → ' + now.regimeVerdict
    });
  }
  // 2. Individual signal flips
  var prevSigMap = {};
  (prev.signals || []).forEach(function(s) { prevSigMap[s.signal] = s.verdict; });
  (now.signals || []).forEach(function(s) {
    var old = prevSigMap[s.signal];
    if (old && old !== s.verdict) {
      var isBullish = s.verdict.indexOf('BULL') !== -1 || s.verdict.indexOf('RISK-ON') !== -1 || s.verdict.indexOf('CALM') !== -1;
      changes.push({
        type: 'SIGNAL', severity: 'medium',
        icon: isBullish ? '▲' : '▼',
        color: isBullish ? '#00dd88' : '#ff5555',
        msg: s.signal + ': ' + old + ' → ' + s.verdict
      });
    }
  });
  // 3. Sector quadrant transitions
  var prevSecMap = {};
  (prev.sectors || []).forEach(function(s) { prevSecMap[s.ticker] = s.quadrant; });
  (now.sectors || []).forEach(function(s) {
    var old = prevSecMap[s.ticker];
    if (old && old !== s.quadrant && old !== '' && s.quadrant !== '') {
      var isUpgrade = (s.quadrant.indexOf('LEADING') !== -1) || (s.quadrant.indexOf('IMPROVING') !== -1 && old.indexOf('LAGGING') !== -1);
      changes.push({
        type: 'QUADRANT', severity: 'medium',
        icon: isUpgrade ? '↗' : '↘',
        color: isUpgrade ? '#44aaff' : '#ff8844',
        msg: s.ticker + ' quadrant: ' + old.replace('RRG ','') + ' → ' + s.quadrant.replace('RRG ','')
      });
    }
  });
  // 4. Rank surges / drops (top 25 movers)
  var prevRankMap = {};
  (prev.top25 || []).forEach(function(u) { prevRankMap[u.ticker] = u.rank; });
  (now.top25 || []).forEach(function(u) {
    var old = prevRankMap[u.ticker];
    if (old != null) {
      var delta = old - u.rank;
      if (delta >= 15) {
        changes.push({
          type: 'RANK', severity: 'low',
          icon: '⚡', color: '#00dd88',
          msg: u.ticker + ' surged +' + delta + ' ranks → #' + u.rank
        });
      }
    } else {
      changes.push({
        type: 'RANK', severity: 'low',
        icon: '★', color: '#44aaff',
        msg: u.ticker + ' entered Top 25 at #' + u.rank
      });
    }
  });
  var nowTop25Set = {};
  (now.top25 || []).forEach(function(u) { nowTop25Set[u.ticker] = true; });
  (prev.top25 || []).slice(0, 15).forEach(function(u) {
    if (!nowTop25Set[u.ticker]) {
      changes.push({
        type: 'RANK', severity: 'low',
        icon: '↓', color: '#ff8844',
        msg: u.ticker + ' dropped out of Top 25 (was #' + u.rank + ')'
      });
    }
  });
  // 5. Rotation signal state changes
  var prevRotMap = {};
  (prev.rotations || []).forEach(function(r) { prevRotMap[r.pair] = r.signal; });
  (now.rotations || []).forEach(function(r) {
    var old = prevRotMap[r.pair];
    if (old && old !== r.signal) {
      var isConfirm = r.signal.indexOf('CONFIRMED') !== -1 && r.signal.indexOf('OFF') === -1;
      changes.push({
        type: 'ROTATION', severity: 'medium',
        icon: isConfirm ? '✦' : '↻',
        color: isConfirm ? '#00dd88' : '#ff8844',
        msg: r.pair + ': ' + old + ' → ' + r.signal
      });
    }
  });
  var sevOrder = { high: 0, medium: 1, low: 2 };
  changes.sort(function(a, b) { return (sevOrder[a.severity] || 9) - (sevOrder[b.severity] || 9); });
  return { items: changes, hasPrevious: true, prevDate: prevDate };
}
// ═══════════════════════════════════════════════════════
// DRILL-DOWN — Uses INDUSTRY SCANNER PARENT column
// ═══════════════════════════════════════════════════════
function buildDrillDown(ss, sectorMap) {
  var d = readRange(ss, 'INDUSTRY SCANNER', 'A1:Q120');
  if (!d || d.length < 2) return {};
  // Find header row
  var hdr = findRow(d, ['ETF', 'PARENT']);
  if (hdr < 0) hdr = findRow(d, ['ETF', 'NAME']);
  if (hdr < 0) return {};
  var drill = {};
  for (var i = hdr + 1; i < d.length; i++) {
    var etf = String(d[i][0] || '').trim();
    var parent = String(d[i][4] || '').trim(); // PARENT is col E (index 4)
    if (!etf || !parent) continue;
    if (!drill[parent]) drill[parent] = [];
    drill[parent].push({
      ticker: etf,
      name: String(d[i][1] || ''),
      category: String(d[i][2] || ''),
      sub: String(d[i][3] || ''),
      rs1w: num(d[i][5]),
      rs1m: num(d[i][6]),
      rs3m: num(d[i][7]),
      rs6m: num(d[i][8]),
      rs12m: num(d[i][9]),
      rsYtd: num(d[i][10]),
      rsVsSpy: num(d[i][11]),
      rank: d[i][12],
      rsVsSector: num(d[i][13]),
      sectorRank: d[i][14],
      momentum: num(d[i][15]),
      quadrant: String(d[i][16] || ''),
      type: 'INDUSTRY',
      composite: num(d[i][11]),        // rsVsSpy serves as composite for display
      percentile: null                 // not directly in INDUSTRY SCANNER
    });
  }
  // Sort each group by rank ascending
  for (var key in drill) {
    drill[key].sort(function(a, b) { return (a.rank || 999) - (b.rank || 999); });
  }
  return drill;
}
// ═══════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════
function readRange(ss, sheetName, range) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getRange(range).getValues();
}
function findRow(data, keywords) {
  for (var i = 0; i < data.length; i++) {
    var joined = data[i].map(function(c){return String(c).toUpperCase();}).join('|');
    var allFound = true;
    for (var k = 0; k < keywords.length; k++) {
      if (joined.indexOf(keywords[k].toUpperCase()) === -1) { allFound = false; break; }
    }
    if (allFound) return i;
  }
  return -1;
}
function findSection(data, sectionText) {
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase().indexOf(sectionText.toUpperCase()) !== -1) return i;
  }
  return -1;
}
function num(v) { return (typeof v === 'number' && !isNaN(v)) ? v : null; }
// ═══════════════════════════════════════════════════════
// PARSERS
// ═══════════════════════════════════════════════════════
// ── REGIME ───────────────────────────────────────────
function parseRegime(ss) {
  var d = readRange(ss, 'REGIME', 'A1:J25');
  var hdr = findRow(d, ['SIGNAL','PROXY','VERDICT']);
  var signals = [];
  if (hdr >= 0) {
    for (var i = hdr+1; i < d.length && i < hdr+8; i++) {
      if (!d[i][1] || String(d[i][1]).indexOf('SIGNAL') !== -1) continue;
      if (String(d[i][0]).indexOf('▶') !== -1 || String(d[i][0]).indexOf('▸') !== -1) break;
      signals.push({
        signal: d[i][1], proxy: d[i][2], reading: num(d[i][3]),
        w1: num(d[i][4]), m1: num(d[i][5]), m3: num(d[i][6]),
        trend: d[i][7], verdict: String(d[i][8]), weight: d[i][9]
      });
    }
  }
  // D12 = row index 11, col D = index 3 (composite regime score, e.g. 0.047 = 4.7%)
  var score = (d.length > 11) ? num(d[11][3]) : null;
  // I12 = row index 11, col I = index 8 (full verdict string)
  var verdictFull = (d.length > 11) ? String(d[11][8] || '') : '';
  // Implications from rows 14-18 (0-indexed 13-17)
  var implications = [];
  for (var r = 13; r <= 17 && r < d.length; r++) {
    var lbl = String(d[r][0] || '').trim();
    var txt = String(d[r][1] || '').trim();
    if (!txt) txt = String(d[r][2] || '').trim();
    if (lbl) implications.push({ label: lbl, text: txt });
  }
  return { signals: signals, score: score, verdictFull: verdictFull, implications: implications };
}
// ── SECTOR MAP (EXTENDED to col O) ───────────────────
function parseSectorMap(ss) {
  var d = readRange(ss, 'SECTOR MAP', 'A1:O25');  // ← CHANGED: was A1:K25
  var context = '';
  for (var i = 0; i < Math.min(5, d.length); i++) {
    if (String(d[i][0]).indexOf('REGIME') !== -1) { context = String(d[i][2] || ''); break; }
  }
  var hdr = findRow(d, ['CONVICTION','QUADRANT']);
  var sectors = [];
  if (hdr >= 0) {
    for (var i = hdr+1; i < d.length; i++) {
      if (!d[i][1]) break;
      sectors.push({
        rank: d[i][0], ticker: d[i][1], name: d[i][2], conviction: num(d[i][3]),
        rs_spy: num(d[i][4]), rs_1w: num(d[i][5]), rs_1m: num(d[i][6]),
        rs_3m: num(d[i][7]), momentum: num(d[i][8]), accel: num(d[i][9]),
        quadrant: String(d[i][10] || ''),
        rotation: String(d[i][11] || ''),          // NEW: col L
        breadth3m: num(d[i][12]),                   // NEW: col M
        breadthStatus: String(d[i][13] || ''),      // NEW: col N
        verdict: String(d[i][14] || '')             // NEW: col O
      });
    }
  }
  return { context: context, sectors: sectors };
}
// ── UNIVERSE RANK ────────────────────────────────────
function parseUniverse(ss) {
  var d = readRange(ss, 'UNIVERSE RANK', 'A1:P260');
  var hdr = findRow(d, ['TICKER','COMPOSITE']);
  var items = [];
  if (hdr >= 0) {
    for (var i = hdr+1; i < d.length; i++) {
      if (!d[i][1]) break;
      items.push({
        rank: d[i][0], ticker: String(d[i][1]), name: String(d[i][2]),
        type: String(d[i][3]), category: String(d[i][4]),
        rs1w: num(d[i][5]), rs1m: num(d[i][6]), rs3m: num(d[i][7]),
        rs6m: num(d[i][8]), rs12m: num(d[i][9]), rsYtd: num(d[i][10]),
        composite: num(d[i][11]), percentile: num(d[i][12]),
        momentum: num(d[i][13]), quadrant: String(d[i][14] || '')
      });
    }
  }
  return items;
}
// ── ROTATION (ALL 3 SECTIONS) ────────────────────────
function parseRotation(ss) {
  var d = readRange(ss, 'ROTATION', 'A1:I45');  // ← CHANGED: was A1:I40
  var pairs = [];
  var sectorSignals = [];
  var creditHealth = [];
  var creditVerdict = '';
  // Section 1: Risk Regime Pairs
  var sec1 = findSection(d, 'RISK REGIME PAIRS');
  if (sec1 >= 0) {
    for (var i = sec1 + 1; i < d.length; i++) {
      // Skip header row
      if (String(d[i][1]).toUpperCase() === 'PAIR') continue;
      // Stop at next section
      if (!d[i][1] || String(d[i][0]).indexOf('▸') !== -1) break;
      pairs.push({
        pair: String(d[i][1]), desc: String(d[i][2]),
        spread1w: num(d[i][3]), spread1m: num(d[i][4]),
        spread3m: num(d[i][5]), spread12m: num(d[i][6]),
        trend: String(d[i][7] || ''), signal: String(d[i][8] || '')
      });
    }
  }
  // Section 2: Sector vs SPY
  var sec2 = findSection(d, 'SECTOR vs SPY');
  if (sec2 >= 0) {
    for (var i = sec2 + 1; i < d.length; i++) {
      if (String(d[i][1]).toUpperCase() === 'SECTOR') continue; // skip header
      if (!d[i][1] || String(d[i][0]).indexOf('▸') !== -1) break;
      sectorSignals.push({
        sector: String(d[i][1]), desc: String(d[i][2]),
        spread1w: num(d[i][3]), spread1m: num(d[i][4]),
        spread3m: num(d[i][5]), spread12m: num(d[i][6]),
        trend: String(d[i][7] || ''), signal: String(d[i][8] || '')
      });
    }
  }
  // Section 3: Credit Health
  var sec3 = findSection(d, 'CREDIT HEALTH');
  if (sec3 >= 0) {
    for (var i = sec3 + 1; i < d.length; i++) {
      // Stop at credit verdict row or next section
      if (String(d[i][0]).indexOf('CREDIT VERDICT') !== -1) {
        creditVerdict = String(d[i][8] || '');
        break;
      }
      // Skip header row
      if (String(d[i][1]).toUpperCase() === 'RATIO' || String(d[i][2]).toUpperCase() === 'MEASURES') continue;
      if (!d[i][1]) break;
      creditHealth.push({
        ratio: String(d[i][1]), measures: String(d[i][2]),
        delta1w: num(d[i][3]), delta1m: num(d[i][4]),
        delta3m: num(d[i][5]), composite: num(d[i][6]),
        trend: String(d[i][7] || ''), status: String(d[i][8] || '')
      });
    }
  }
  return {
    pairs: pairs,
    sectorSignals: sectorSignals,
    creditHealth: creditHealth,
    creditVerdict: creditVerdict,
    movers: []                        // backward compat
  };
}
// ── EMERGING (EXTENDED to col O) ─────────────────────
function parseEmerging(ss) {
  var d = readRange(ss, 'EMERGING', 'A1:O30');  // ← CHANGED: was A1:J30
  var hdr = findRow(d, ['TICKER','RS 1W']);
  var items = [];
  if (hdr >= 0) {
    for (var i = hdr+1; i < d.length && items.length < 20; i++) {
      if (!d[i][1]) break;
      items.push({
        rank: d[i][0], ticker: String(d[i][1]), name: String(d[i][2]),
        type: String(d[i][3]), sector: String(d[i][4]),
        rs1w: num(d[i][5]), rs1m: num(d[i][6]), rs3m: num(d[i][7]),
        rsComposite: num(d[i][8]), compRank: d[i][9],
        percentile: num(d[i][10]),         // NEW: col K
        quadrant: String(d[i][11] || ''),  // NEW: col L
        rotation: String(d[i][12] || ''),  // NEW: col M
        rankDelta: d[i][13],               // NEW: col N (Δ W→M)
        note: String(d[i][14] || '')       // NEW: col O
      });
    }
  }
  return items;
}
// ── RATIO COCKPIT ────────────────────────────────────
function parseRatios(ss) {
  var d = readRange(ss, 'RATIO COCKPIT', 'A1:K65');
  var sections = {riskOnOff:[], credit:[], sectorLeadership:[], inflation:[], globalMacro:[]};
  var currentKey = null;
  for (var i = 0; i < d.length; i++) {
    var cell = String(d[i][0]).toUpperCase();
    if (cell.indexOf('▸') !== -1 || cell.indexOf('►') !== -1) {
      if (cell.indexOf('RISK') !== -1) currentKey = 'riskOnOff';
      else if (cell.indexOf('CREDIT') !== -1) currentKey = 'credit';
      else if (cell.indexOf('EQUITY') !== -1 || cell.indexOf('SECTOR') !== -1) currentKey = 'sectorLeadership';
      else if (cell.indexOf('INFLATION') !== -1 || cell.indexOf('COMMOD') !== -1) currentKey = 'inflation';
      else if (cell.indexOf('GLOBAL') !== -1 || cell.indexOf('MACRO') !== -1) currentKey = 'globalMacro';
      continue;
    }
    if (currentKey && d[i][1] && d[i][2] && String(d[i][1]) !== 'NUM') {
      var numStr = String(d[i][1]);
      var denStr = String(d[i][2]);
      if (numStr.length > 0 && numStr.length < 8 && denStr.length > 0 && denStr.length < 8) {
        sections[currentKey].push({
          num: numStr, den: denStr, desc: String(d[i][3] || ''),
          spread1w: num(d[i][4]), spread1m: num(d[i][5]),
          spread3m: num(d[i][6]), spread12m: num(d[i][7]),
          delta: num(d[i][8]), trend: String(d[i][9] || ''),
          signal: String(d[i][10] || '')
        });
      }
    }
  }
  return sections;
}
// ── CREDIT MONITOR ───────────────────────────────────
function parseCredit(ss) {
  var d = readRange(ss, 'CREDIT MONITOR', 'A1:J25');
  var hdr = findRow(d, ['RATIO','MEASURES']);
  var spreads = [];
  if (hdr >= 0) {
    for (var i = hdr+1; i < d.length; i++) {
      if (!d[i][0] || String(d[i][0]).indexOf('▸') !== -1 || String(d[i][0]).indexOf('COMPOSITE') !== -1) break;
      spreads.push({
        ratio: String(d[i][0]), measures: String(d[i][1]),
        d1w: num(d[i][2]), d1m: num(d[i][3]), d3m: num(d[i][4]),
        d6m: num(d[i][5]), d12m: num(d[i][6]),
        composite: num(d[i][7]), rank: d[i][8], delta: num(d[i][9])
      });
    }
  }
  return { spreads: spreads };
}
// ── EQUAL WEIGHT ─────────────────────────────────────
function parseEqualWeight(ss) {
  var d = readRange(ss, 'EQUAL WEIGHT', 'A1:L40');
  var g1 = findRow(d, ['RATIO','DESCRIPTION']);
  var gauges = [];
  if (g1 >= 0) {
    for (var i = g1+1; i < d.length; i++) {
      if (!d[i][0] || String(d[i][0]).indexOf('▸') !== -1 || String(d[i][0]).indexOf('COMPOSITE') !== -1) break;
      gauges.push({
        ratio: String(d[i][0]), desc: String(d[i][1]),
        spread1w: num(d[i][5]), spread1m: num(d[i][8]), spread3m: num(d[i][11])
      });
    }
  }
  var g2 = findSection(d, 'SECTOR BREADTH');
  var sectorBreadth = [];
  if (g2 >= 0) {
    for (var i = g2+2; i < d.length; i++) {
      if (!d[i][0] || String(d[i][0]).indexOf('▸') !== -1 || String(d[i][0]).indexOf('GOOGLE') !== -1) break;
      if (String(d[i][0]).indexOf('RATIO') !== -1) continue;
      sectorBreadth.push({
        sector: String(d[i][1] || d[i][0]),
        spread1w: num(d[i][5]), spread1m: num(d[i][8]), spread3m: num(d[i][11])
      });
    }
  }
  return { gauges: gauges, sectorBreadth: sectorBreadth };
}
// ── CONVICTION (with text label) ─────────────────────
function parseConviction(ss) {
  var d = readRange(ss, 'CONVICTION', 'A1:O260');  // ← CHANGED: was 'Conviction'
  var hdr = findRow(d, ['TICKER', 'CONVICTION']);
  var items = [];
  if (hdr >= 0) {
    for (var i = hdr+1; i < d.length; i++) {
      if (!d[i][1]) break;
      // Stop if we hit the "BOTTOM 10" section marker
      if (String(d[i][0]).indexOf('▸') !== -1 || String(d[i][0]).indexOf('BOTTOM') !== -1) break;
      items.push({
        rank:            d[i][0],
        ticker:          String(d[i][1]),
        name:            String(d[i][2]),
        type:            String(d[i][3]),
        category:        String(d[i][4]),
        composite:       num(d[i][5]),
        percentile:      num(d[i][6]),
        rs1w:            num(d[i][7]),
        rs1m:            num(d[i][8]),
        rs3m:            num(d[i][9]),
        momentum:        num(d[i][10]),
        accel:           num(d[i][11]),
        quadrant:        String(d[i][12] || ''),
        rotation:        String(d[i][13] || ''),
        convictionLabel: String(d[i][14] || '')  // ← CHANGED: was num(d[i][14])
      });
    }
  }
  // ── Bottom 10 (weakest / short candidates) ──
  var bottom = [];
  var bot = findSection(d, 'BOTTOM 10');
  if (bot >= 0) {
    for (var i = bot + 1; i < d.length; i++) {
      if (String(d[i][1]).toUpperCase() === 'TICKER') continue;
      if (!d[i][1]) break;
      bottom.push({
        rank:            d[i][0],
        ticker:          String(d[i][1]),
        name:            String(d[i][2]),
        type:            String(d[i][3]),
        category:        String(d[i][4]),
        composite:       num(d[i][5]),
        percentile:      num(d[i][6]),
        rs1w:            num(d[i][7]),
        rs1m:            num(d[i][8]),
        rs3m:            num(d[i][9]),
        momentum:        num(d[i][10]),
        accel:           num(d[i][11]),
        quadrant:        String(d[i][12] || ''),
        rotation:        String(d[i][13] || ''),
        convictionLabel: String(d[i][14] || '')
      });
    }
  }
  return { top: items, bottom: bottom };
}
