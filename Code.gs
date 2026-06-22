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

  function isItemRow(row) {
    return /^\d+$/.test(String(row[COL_ITEM] || '').trim());
  }

  function isBool(val) {
    if (typeof val === 'boolean') return val;
    var s = String(val || '').toLowerCase().trim();
    return s === 'yes' || s === 'true' || s === '1';
  }

  function formatItem(row) {
    var item   = String(row[COL_ITEM]   || '').trim();
    var title  = String(row[COL_TITLE]  || '').trim();
    var studio = String(row[COL_STUDIO] || '').trim();
    var age    = String(row[COL_AGE]    || '').trim();
    var level  = String(row[COL_LEVEL]  || '').trim();
    var cat    = String(row[COL_CAT]    || '').trim();
    var style  = String(row[COL_STYLE]  || '').trim();

    var meta = [age, level, cat, style].filter(Boolean).join(' · ');
    var line = '#' + item;
    if (title)  line += '  ' + title;
    if (meta)   line += '  ·  ' + meta;
    return line;
  }

  var onStageText = null;
  var onStageIdx  = -1;

  for (var i = 1; i < data.length; i++) {
    if (!isItemRow(data[i])) continue;
    if (isBool(data[i][COL_ONSTAGE])) {
      onStageText = formatItem(data[i]);
      onStageIdx  = i;
      break;
    }
  }

  // Next 10 items that are not danced or scratched
  var upNext     = [];
  var startSearch = onStageIdx >= 0 ? onStageIdx + 1 : 1;

  for (var j = startSearch; j < data.length && upNext.length < 10; j++) {
    if (!isItemRow(data[j]))          continue;
    if (isBool(data[j][COL_SCRATCHED])) continue;
    if (isBool(data[j][COL_DANCED]))    continue;
    upNext.push(formatItem(data[j]));
  }

  return { onStage: onStageText, upNext: upNext };
}
