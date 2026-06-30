// ─────────────────────────────────────────────────────────────────────────────
// TXG Weekly FactSet Consensus — Dashboard Builder
//
// SETUP INSTRUCTIONS:
//  1. Open the dashboard Google Sheet
//  2. Extensions → Apps Script → replace all content with this file → Save
//  3. Run updateDashboard() once manually to authorize
//  4. Set up trigger: Triggers → Add Trigger → updateDashboard → Time-driven
//     → Week timer → Monday → 3pm–4pm
//
// HOW IT WORKS:
//  Each time it runs, it finds the most recently modified Excel file in your
//  Drive folder, converts it to a temp Google Sheet, reads the data, writes
//  a formatted dashboard, then deletes the temp sheet.
// ─────────────────────────────────────────────────────────────────────────────

var DRIVE_FOLDER_ID = '1f9Wx7BAg5mhaeCpfM0R-hqrRfCn761q9';

// Brand colors
var C = {
  navy:        '#1B3A6B',
  navyText:    '#FFFFFF',
  sectionBg:   '#D6E4F0',
  sectionText: '#1B3A6B',
  totalBg:     '#EBF3FB',
  altRow:      '#F7FAFD',
  positive:    '#C6EFCE',
  negative:    '#FFC7CE',
  neutral:     '#FFFFFF',
  headerCol:   '#BDD7EE',
  analystBuy:  '#C6EFCE',
  analystHold: '#FFEB9C',
  analystSell: '#FFC7CE',
  border:      '#ADB5BD',
};

// ── Entry point ───────────────────────────────────────────────────────────────

function updateDashboard() {
  var excelFile = findLatestExcelFile();
  if (!excelFile) {
    SpreadsheetApp.getUi().alert('No Excel file found in the Drive folder.');
    return;
  }

  Logger.log('Reading: ' + excelFile.getName() + ' (' + excelFile.getLastUpdated() + ')');

  var tempSS = convertExcelToGoogleSheet(excelFile);
  try {
    var data = extractData(tempSS);
    writeDashboard(data);
    Logger.log('Dashboard updated successfully.');
  } finally {
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  }
}

// ── Drive helpers ─────────────────────────────────────────────────────────────

function findLatestExcelFile() {
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  var latest = null;
  var latestDate = new Date(0);
  while (files.hasNext()) {
    var f = files.next();
    if (f.getLastUpdated() > latestDate) {
      latestDate = f.getLastUpdated();
      latest = f;
    }
  }
  return latest;
}

function convertExcelToGoogleSheet(excelFile) {
  var resource = { title: '_TEMP_factset_dashboard_', mimeType: MimeType.GOOGLE_SHEETS };
  var converted = Drive.Files.insert(resource, excelFile.getBlob(), { convert: true });
  return SpreadsheetApp.openById(converted.id);
}

// ── Data extraction ───────────────────────────────────────────────────────────

function extractData(ss) {
  var cs = ss.getSheetByName('Hardcode Weekly FactSet Consensus') ||
           ss.getSheetByName('FactSet Consensus') ||
           ss.getSheetByName('Weekly FactSet Consensus');
  var as = ss.getSheetByName('Detailed Analyst Consensus');
  if (!cs) throw new Error('Could not find consensus sheet. Available: ' + ss.getSheets().map(function(s){return s.getName();}).join(', '));

  var cData = cs.getDataRange().getValues();
  var aData = as ? as.getDataRange().getValues() : [];

  // Row indices (0-based)
  var dateRow     = cData[4];   // Row 5
  var currentDate = dateRow[3];
  var priorDate   = dateRow[9];

  var periods = ['Q1', 'Q2', 'Q3', 'Q4', 'FY26', 'FY27'];

  // Extract a data row: cols D-I (current) and J-O (prior), then P-U (change)
  function extractRow(rowIdx) {
    var row = cData[rowIdx];
    var curr  = [row[3], row[4], row[5], row[6], row[7], row[8]];
    var prior = [row[9], row[10], row[11], row[12], row[13], row[14]];
    var chg   = [row[16], row[17], row[18], row[19], row[20], row[21]];
    return { curr: curr, prior: prior, chg: chg };
  }

  var consensus = {
    currentDate: currentDate,
    priorDate:   priorDate,
    periods:     periods,
    rows: {
      instChromium:  { label: '  Chromium',              data: extractRow(13) },
      instSpatial:   { label: '  Spatial',               data: extractRow(14) },
      instVisium:    { label: '  Visium',                data: extractRow(15) },
      instXenium:    { label: '  Xenium',                data: extractRow(16) },
      instTotal:     { label: 'Instrument Revenue',      data: extractRow(17), isTotal: true },
      consChromium:  { label: '  Chromium',              data: extractRow(19) },
      consSpatial:   { label: '  Spatial',               data: extractRow(20) },
      consVisium:    { label: '  Visium',                data: extractRow(21) },
      consXenium:    { label: '  Xenium',                data: extractRow(22) },
      consTotal:     { label: 'Consumables Revenue',     data: extractRow(23), isTotal: true },
      services:      { label: 'Services Revenue',        data: extractRow(24) },
      totalRev:      { label: 'Total Revenue',           data: extractRow(25), isTotal: true },
      cogs:          { label: 'COGS',                    data: extractRow(27) },
      grossProfit:   { label: 'Gross Profit',            data: extractRow(28), isTotal: true },
      grossMargin:   { label: 'Gross Margin %',          data: extractRow(29), isMargin: true },
      rd:            { label: 'R&D',                     data: extractRow(31) },
      sga:           { label: 'SG&A',                    data: extractRow(32) },
      totalOpex:     { label: 'Total Opex',              data: extractRow(33), isTotal: true },
      ebit:          { label: 'EBIT',                    data: extractRow(34), isTotal: true },
      netIncome:     { label: 'Net Income',              data: extractRow(36), isTotal: true },
    }
  };

  // Extract analyst data
  var analysts = [];
  if (aData.length > 0) {
    var lastFirm = '';
    for (var i = 9; i < aData.length; i++) {
      var r = aData[i];
      var firm     = String(r[2] || '').trim();
      var analyst  = String(r[5] || '').trim();
      var colLabel = String(r[4] || '').trim();

      if (!firm && !analyst && !colLabel) continue;
      if (!analyst && colLabel.indexOf('Current') >= 0) continue;

      if (colLabel === 'Mean' || colLabel === 'Median' || colLabel === 'FactSet Consensus') {
        analysts.push({ isSummary: true, label: colLabel,
          pt: r[8], q1: r[9], q2: r[10], q3: r[11], q4: r[12], fy26: r[13], fy27: r[14] });
        continue;
      }

      if (firm) lastFirm = firm;
      if (!analyst) continue;

      analysts.push({
        firm: lastFirm, analyst: analyst,
        date: r[6], rating: r[7], pt: r[8],
        q1: r[9], q2: r[10], q3: r[11], q4: r[12],
        fy26: r[13], fy27: r[14],
        growth26: r[15], growth27: r[18]
      });
    }
  }
  consensus.analysts = analysts;
  return consensus;
}

// ── Dashboard writer ──────────────────────────────────────────────────────────

function writeDashboard(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  writeConsensusTab(ss, data);
  writeAnalystTab(ss, data);
}

function fmt(v, isMargin) {
  if (v === null || v === undefined || v === '' || v === 'na' || v === '#N/A') return 'na';
  var n = parseFloat(v);
  if (isNaN(n)) return String(v);
  if (isMargin) return (n * 100).toFixed(1) + '%';
  return n.toFixed(1);
}

function fmtChg(v, isMargin) {
  if (v === null || v === undefined || v === '' || v === 'na' || v === '#N/A') return '';
  var n = parseFloat(v);
  if (isNaN(n) || Math.abs(n) < 0.0001) return '—';
  if (isMargin) {
    var pct = (n * 100).toFixed(1);
    return (n > 0 ? '+' : '') + pct + 'pp';
  }
  return (n > 0 ? '+' : '') + n.toFixed(1);
}

function writeConsensusTab(ss, data) {
  var tabName = 'Consensus';
  var ws = ss.getSheetByName(tabName);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(tabName, 0);

  var dateStr = data.currentDate instanceof Date
    ? Utilities.formatDate(data.currentDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')
    : String(data.currentDate);
  var priorStr = data.priorDate instanceof Date
    ? Utilities.formatDate(data.priorDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')
    : String(data.priorDate);

  var headers = ['', 'Q1\'26', 'Q2\'26', 'Q3\'26', 'Q4\'26', 'FY\'26', 'FY\'27',
                 '', 'WoW Q1', 'WoW Q2', 'WoW Q3', 'WoW Q4', 'WoW FY\'26', 'WoW FY\'27'];

  // ── Title rows ──
  ws.getRange(1, 1, 1, 14).merge()
    .setValue('TXG — Weekly FactSet Consensus')
    .setBackground(C.navy).setFontColor(C.navyText)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 36);

  ws.getRange(2, 1).setValue('As of ' + dateStr).setFontWeight('bold').setFontSize(10);
  ws.getRange(2, 8).setValue('WoW vs ' + priorStr).setFontStyle('italic').setFontSize(9)
    .setFontColor('#555555');

  // ── Column headers ──
  var hRow = ws.getRange(3, 1, 1, 14);
  hRow.setValues([headers]);
  hRow.setBackground(C.navy).setFontColor(C.navyText).setFontWeight('bold')
    .setHorizontalAlignment('center');
  ws.setRowHeight(3, 24);

  // ── Data sections ──
  var rowNum = 4;

  function writeSection(title, keys) {
    // Section header
    ws.getRange(rowNum, 1, 1, 14).merge()
      .setValue(title)
      .setBackground(C.sectionBg).setFontColor(C.sectionText)
      .setFontWeight('bold').setFontSize(10);
    ws.setRowHeight(rowNum, 20);
    rowNum++;

    keys.forEach(function(key, idx) {
      var item = data.rows[key];
      var d = item.data;
      var isMargin = !!item.isMargin;
      var isTotal = !!item.isTotal;
      var bg = isTotal ? C.totalBg : (idx % 2 === 0 ? C.neutral : C.altRow);

      var rowValues = [item.label];
      // Current values
      for (var i = 0; i < 6; i++) rowValues.push(fmt(d.curr[i], isMargin));
      rowValues.push(''); // spacer
      // WoW changes
      for (var i = 0; i < 6; i++) rowValues.push(fmtChg(d.chg[i], isMargin));

      var r = ws.getRange(rowNum, 1, 1, 14);
      r.setValues([rowValues]);
      r.setBackground(bg);
      if (isTotal) r.setFontWeight('bold');

      // Color the WoW cells
      for (var i = 0; i < 6; i++) {
        var chgVal = parseFloat(d.chg[i]);
        if (!isNaN(chgVal) && Math.abs(chgVal) > 0.0001) {
          var cellBg = chgVal > 0 ? C.positive : C.negative;
          ws.getRange(rowNum, 9 + i).setBackground(cellBg);
        }
      }

      // Right-align numeric columns
      ws.getRange(rowNum, 2, 1, 6).setHorizontalAlignment('right');
      ws.getRange(rowNum, 9, 1, 6).setHorizontalAlignment('right');

      ws.setRowHeight(rowNum, 18);
      rowNum++;
    });
    rowNum++; // blank spacer row
  }

  writeSection('INSTRUMENT REVENUE ($M)', ['instChromium','instSpatial','instVisium','instXenium','instTotal']);
  writeSection('CONSUMABLES REVENUE ($M)', ['consChromium','consSpatial','consVisium','consXenium','consTotal']);
  writeSection('SERVICES & TOTAL REVENUE ($M)', ['services','totalRev']);
  writeSection('P&L ($M)', ['cogs','grossProfit','grossMargin','rd','sga','totalOpex','ebit','netIncome']);

  // Note row
  ws.getRange(rowNum, 1, 1, 14).merge()
    .setValue('Note: Product segment figures are averages across available analysts and may not tie to total revenue.')
    .setFontStyle('italic').setFontSize(8).setFontColor('#666666');

  // ── Column widths ──
  ws.setColumnWidth(1, 180);
  for (var c = 2; c <= 7; c++) ws.setColumnWidth(c, 75);
  ws.setColumnWidth(8, 20);
  for (var c = 9; c <= 14; c++) ws.setColumnWidth(c, 72);

  // ── Borders ──
  ws.getRange(3, 1, rowNum - 3, 14)
    .setBorder(true, true, true, true, true, false,
               C.border, SpreadsheetApp.BorderStyle.SOLID);

  // ── Freeze ──
  ws.setFrozenRows(3);
  ws.setFrozenColumns(1);

  // ── Header column color ──
  ws.getRange(4, 1, rowNum - 4, 1).setFontColor(C.sectionText);

  // Move this sheet to front
  ss.setActiveSheet(ws);
  ss.moveActiveSheet(1);
}

function writeAnalystTab(ss, data) {
  var tabName = 'Analyst Detail';
  var ws = ss.getSheetByName(tabName);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(tabName, 1);

  // Title
  ws.getRange(1, 1, 1, 13).merge()
    .setValue('TXG — Analyst Consensus Detail')
    .setBackground(C.navy).setFontColor(C.navyText)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 36);

  // Headers
  var headers = ['Firm', 'Analyst', 'Date', 'Rating', 'Price\nTarget',
                 'Q1\'26E', 'Q2\'26E', 'Q3\'26E', 'Q4\'26E', 'FY\'26E', 'FY\'27E',
                 'FY\'26\nGrowth', 'FY\'27\nGrowth'];
  var hRange = ws.getRange(2, 1, 1, 13);
  hRange.setValues([headers]);
  hRange.setBackground(C.navy).setFontColor(C.navyText)
    .setFontWeight('bold').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  ws.setRowHeight(2, 30);

  var rowNum = 3;
  data.analysts.forEach(function(a, idx) {
    if (a.isSummary) {
      // Mean / Median / FactSet rows
      var bg = C.totalBg;
      var vals = [a.label, '', '', '', fmtPT(a.pt),
                  fmtRev(a.q1), fmtRev(a.q2), fmtRev(a.q3), fmtRev(a.q4),
                  fmtRev(a.fy26), fmtRev(a.fy27), '', ''];
      var r = ws.getRange(rowNum, 1, 1, 13);
      r.setValues([vals]).setBackground(bg).setFontWeight('bold');
      ws.getRange(rowNum, 1, 1, 13).setBorder(true, true, true, true, false, false,
        C.navy, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    } else {
      var bg = idx % 2 === 0 ? C.neutral : C.altRow;
      var ratingColor = C.neutral;
      if (a.rating === 'Buy') ratingColor = C.analystBuy;
      else if (a.rating === 'Hold') ratingColor = C.analystHold;
      else if (a.rating === 'Sell' || a.rating === 'Underperform') ratingColor = C.analystSell;

      var dateStr = a.date instanceof Date
        ? Utilities.formatDate(a.date, Session.getScriptTimeZone(), 'MM/dd/yy')
        : (a.date || '');

      var vals = [a.firm || '', a.analyst || '', dateStr, a.rating || '',
                  fmtPT(a.pt),
                  fmtRev(a.q1), fmtRev(a.q2), fmtRev(a.q3), fmtRev(a.q4),
                  fmtRev(a.fy26), fmtRev(a.fy27),
                  fmtGrowth(a.growth26), fmtGrowth(a.growth27)];
      var r = ws.getRange(rowNum, 1, 1, 13);
      r.setValues([vals]).setBackground(bg);
      ws.getRange(rowNum, 4).setBackground(ratingColor).setHorizontalAlignment('center');
    }

    // Right-align numeric columns
    ws.getRange(rowNum, 5, 1, 9).setHorizontalAlignment('right');
    ws.setRowHeight(rowNum, 18);
    rowNum++;
  });

  // Column widths
  ws.setColumnWidth(1, 110);  // Firm
  ws.setColumnWidth(2, 140);  // Analyst
  ws.setColumnWidth(3, 75);   // Date
  ws.setColumnWidth(4, 65);   // Rating
  ws.setColumnWidth(5, 60);   // PT
  for (var c = 6; c <= 11; c++) ws.setColumnWidth(c, 72);
  ws.setColumnWidth(12, 72);
  ws.setColumnWidth(13, 72);

  ws.setFrozenRows(2);
  ws.setFrozenColumns(2);
}

function fmtPT(v) {
  if (!v || v === 'NA' || v === 'na') return 'NA';
  var n = parseFloat(v);
  return isNaN(n) ? String(v) : '$' + n.toFixed(0);
}

function fmtRev(v) {
  if (!v || v === 'na') return '';
  var n = parseFloat(v);
  return isNaN(n) ? '' : '$' + n.toFixed(1);
}

function fmtGrowth(v) {
  if (v === null || v === undefined || v === '') return '';
  var n = parseFloat(v);
  if (isNaN(n)) return '';
  return (n >= 0 ? '+' : '') + (n * 100).toFixed(1) + '%';
}
