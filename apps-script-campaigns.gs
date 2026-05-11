// ─────────────────────────────────────────────────────────────────────────────
//  Google Apps Script — Campaign Comparator
//  Music & Brand Consumer Comms
//
//  HOW TO USE:
//  1. Open your "2024-2026 Top Campaign Numbers" Google Sheet
//  2. Click Extensions → Apps Script
//  3. Delete any existing code and paste ALL of this in
//  4. Save (⌘S)
//  5. Deploy → New deployment → Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  6. Copy the URL and paste it into campaign-comparator.html
//     where it says PASTE_CAMPAIGNS_SCRIPT_URL_HERE
// ─────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  if (e.parameter && e.parameter.action === 'import') {
    return importCampaigns();
  }
  return ContentService
    .createTextOutput('Campaign Comparator Apps Script is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function importCampaigns() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var campaigns = [];

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var lastRow = sheet.getLastRow();
      if (lastRow <= 2) continue;

      // Try to extract year from sheet name (e.g. "2026 Campaigns" → 2026)
      var sheetYear = extractYear(sheet.getName());

      var data = sheet.getRange(3, 1, lastRow - 2, 14).getValues();

      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var name = String(row[0]).trim();
        if (!name) continue;

        // Skip header-like rows
        if (name.toLowerCase() === 'campaign') continue;

        var hits       = parseNum(row[1]);
        var imps       = parseNum(row[2]);
        var online     = parseNum(row[3]);
        var pressSoc   = parseNum(row[5]);
        var fanAcc     = parseNum(row[7]);
        var influencer = parseNum(row[9]);
        var reporterSoc= parseNum(row[11]);
        var notes      = String(row[13] || '').trim();

        // social = press social + fan account + reporter social
        var social = addNums(pressSoc, fanAcc, reporterSoc);

        if (!hits) continue; // skip rows with no hit data

        campaigns.push({
          name:       name,
          year:       sheetYear,
          hits:       hits,
          imps:       imps,
          online:     online,
          social:     social,
          influencer: influencer,
          notes:      notes
        });
      }
    }

    return jsonResponse({ status: 'ok', campaigns: campaigns });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// Extract a 4-digit year from a string, default to current year
function extractYear(str) {
  var match = String(str).match(/20\d{2}/);
  return match ? parseInt(match[0]) : new Date().getFullYear();
}

// Parse a number from a cell value (handles commas, Date objects, etc.)
function parseNum(val) {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return val === 0 ? null : val;
  var n = parseInt(String(val).replace(/,/g, ''));
  return isNaN(n) || n === 0 ? null : n;
}

// Sum multiple nullable numbers; returns null if all are null
function addNums() {
  var sum = 0, hasAny = false;
  for (var i = 0; i < arguments.length; i++) {
    if (arguments[i] !== null && arguments[i] !== undefined) {
      sum += arguments[i];
      hasAny = true;
    }
  }
  return hasAny ? sum : null;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
