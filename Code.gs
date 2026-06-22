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

  // Detect prize giving / awards rows (non-numeric item column but has "prize" or "award" text)
  function isPrizeRow(row) {
    if (isItemRow(row)) return false; // already handled as a dance item
    var checkCols = [COL_ITEM, COL_TITLE, COL_STYLE, COL_CAT];
    for (var c = 0; c < checkCols.length; c++) {
      var v = String(row[checkCols[c]] || '').toLowerCase().trim();
      if (v.indexOf('prize') >= 0 || v.indexOf('award') >= 0) return true;
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
    // Collect whatever descriptive text is available
    var cols = [COL_ITEM, COL_TITLE, COL_STYLE, COL_CAT, COL_AGE, COL_LEVEL];
    var parts = [];
    for (var c = 0; c < cols.length; c++) {
      var v = String(row[cols[c]] || '').trim();
      if (v) parts.push(v);
    }
    // Deduplicate while keeping order
    var seen = {}, unique = [];
    for (var p = 0; p < parts.length; p++) {
      if (!seen[parts[p]]) { seen[parts[p]] = true; unique.push(parts[p]); }
    }
    return 'PRIZE GIVING  —  ' + (unique.join('  ·  ') || 'Awards');
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
