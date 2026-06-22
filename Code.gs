/**
 * Pulse 2026 — Live Schedule
 * Google Apps Script (Code.gs)
 *
 * HOW TO USE:
 *  1. Go to script.google.com → New project
 *  2. Paste this into Code.gs
 *  3. Add a new HTML file called "Index" and paste Index.html content
 *  4. Deploy → New deployment → Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. Copy the web app URL
 *  6. Paste it into the GitHub wrapper HTML (src of the iframe)
 *
 * The Schedule Google Sheet must be shared with the Apps Script account
 * (or the sheet must be publicly readable).
 */

var SHEET_ID = '1ZFwvTm7ro_1GwRxG8nk_97Y0TMoGdWpR5UjknwtblEA';
var TAB_NAME = 'Schedule';

// Column indices (0-based) in the Schedule tab
var COL_ITEM      = 1;   // B — item number
var COL_STUDIO    = 2;   // C — studio name
var COL_AGE       = 3;   // D — age division
var COL_LEVEL     = 4;   // E — level
var COL_CAT       = 5;   // F — category
var COL_STYLE     = 6;   // G — discipline / style
var COL_TITLE     = 7;   // H — dance title
var COL_DANCER    = 8;   // I — dancer name(s)
var COL_SCRATCHED = 18;  // S — scratched (checkbox / "Yes")
var COL_ONSTAGE   = 19;  // T — on stage  (checkbox / "Yes")
var COL_DANCED    = 20;  // U — danced    (checkbox / "Yes")

// ── Serve the HTML page ──────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pulse 2026 — Live Schedule')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Data function called by the client every 5 s ─────────────────────────────
function getStageData() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  var data  = sheet.getDataRange().getValues();

  function isBool(val) {
    if (typeof val === 'boolean') return val;
    var s = String(val || '').toLowerCase().trim();
    return s === 'yes' || s === 'true' || s === '1';
  }

  function isItemRow(row) {
    return /^\d+$/.test(String(row[COL_ITEM] || '').trim());
  }

  // Detect prize giving / awards rows — scan ALL columns for relevant keywords
  var PRIZE_KEYWORDS = ['prize', 'award', 'ceremony', 'pg '];
  function isPrizeRow(row) {
    if (isItemRow(row)) return false;
    for (var c = 0; c < row.length; c++) {
      var v = String(row[c] || '').toLowerCase().trim();
      if (!v) continue;
      for (var k = 0; k < PRIZE_KEYWORDS.length; k++) {
        if (v.indexOf(PRIZE_KEYWORDS[k]) >= 0) return true;
      }
    }
    return false;
  }

  function isDisplayRow(row) {
    return isItemRow(row) || isPrizeRow(row);
  }

  function formatItem(row) {
    var item  = String(row[COL_ITEM]  || '').trim();
    var title = String(row[COL_TITLE] || '').trim();
    var age   = String(row[COL_AGE]   || '').trim();
    var level = String(row[COL_LEVEL] || '').trim();
    var cat   = String(row[COL_CAT]   || '').trim();
    var style = String(row[COL_STYLE] || '').trim();

    var meta = [age, level, cat, style].filter(Boolean).join(' · ');
    var line = '#' + item;
    if (title) line += '  ' + title;
    if (meta)  line += '  ·  ' + meta;
    return line;
  }

  function formatPrize(row) {
    // Scan all columns for meaningful text (skip empty, skip pure numbers, skip time-like values)
    var seen = {}, parts = [];
    for (var c = 0; c < row.length; c++) {
      var v = String(row[c] || '').trim();
      if (!v) continue;
      if (/^\d+$/.test(v)) continue;           // skip pure numbers
      if (/^\d{1,2}:\d{2}$/.test(v)) continue; // skip times like 13:47
      if (!seen[v]) { seen[v] = true; parts.push(v); }
    }
    return 'PRIZE GIVING  —  ' + (parts.join('  ·  ') || 'Awards Ceremony');
  }

  function formatRow(row) {
    return isItemRow(row) ? formatItem(row) : formatPrize(row);
  }

  var onStageText = null;
  var onStageIdx  = -1;

  for (var i = 1; i < data.length; i++) {
    if (!isDisplayRow(data[i])) continue;
    if (isBool(data[i][COL_ONSTAGE])) {
      onStageText = formatRow(data[i]);
      onStageIdx  = i;
      break;
    }
  }

  // Next 10 items (dance or prize giving) that are not yet danced or scratched
  var upNext      = [];
  var startSearch = onStageIdx >= 0 ? onStageIdx + 1 : 1;

  for (var j = startSearch; j < data.length && upNext.length < 10; j++) {
    if (!isDisplayRow(data[j]))           continue;
    if (isBool(data[j][COL_SCRATCHED]))   continue;
    if (isBool(data[j][COL_DANCED]))      continue;
    upNext.push(formatRow(data[j]));
  }

  return { onStage: onStageText, upNext: upNext };
}
