// ─────────────────────────────────────────────────────────────────────────────
//  Google Apps Script — Project Tracker
//  Music & Brand Consumer Comms
//
//  HOW TO USE:
//  1. In your Google Sheet, click Extensions → Apps Script
//  2. Delete any existing code and paste ALL of this code in
//  3. Click Save (⌘S or the floppy disk icon)
//  4. Click Deploy → New deployment → Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  5. Click Deploy, accept permissions, copy the URL shown
//  6. Paste that URL into index.html where it says PASTE_YOUR_URL_HERE
// ─────────────────────────────────────────────────────────────────────────────

// ── CONFIG ──────────────────────────────────────────────────────────────────
// If your sheet tab has a specific name, put it here (e.g. "2026 Projects").
// Leave as "" to use the first tab automatically.
var SHEET_NAME = "2026 Project Tracker";

// ─────────────────────────────────────────────────────────────────────────────

var MONTHS = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December"
];

// doGet handles two things:
//   - No params → health check (open URL in browser to confirm script is live)
//   - ?action=import → returns all projects from the sheet as JSON
function doGet(e) {
  if (e.parameter && e.parameter.action === "import") {
    return importAllProjects();
  }
  return ContentService
    .createTextOutput("Apps Script is running. Ready to accept project entries.")
    .setMimeType(ContentService.MimeType.TEXT);
}

// Read all rows from the sheet and return them as a JSON array of project objects.
function importAllProjects() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allSheets = ss.getSheets();
    var allSheetNames = allSheets.map(function(s) { return s.getName(); });

    // Find the sheet: exact name → case-insensitive → first sheet
    var sheet = null;
    if (SHEET_NAME) {
      sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) {
        var target = SHEET_NAME.trim().toLowerCase();
        for (var j = 0; j < allSheets.length; j++) {
          if (allSheets[j].getName().trim().toLowerCase() === target) {
            sheet = allSheets[j];
            break;
          }
        }
      }
    }
    if (!sheet) sheet = allSheets[0];

    var sheetUsed = sheet.getName();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 2) return jsonResponse({ status: "ok", projects: [], sheetUsed: sheetUsed, allSheets: allSheetNames });

    // Skip rows 1–2 (merged title row + column header row)
    var data = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
    var projects = [];
    var currentYear = new Date().getFullYear();
    var monthHeaderPattern = /^([A-Z][a-z]+) (\d{4})$/;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var col1 = String(row[0]).trim(); // PROJECT
      var col4 = String(row[3]).trim(); // DATE

      // Month headers can be in col A (merged cell) OR col D (date column)
      // Check BEFORE skipping empty col A rows
      var headerCandidate = col1 || col4;
      var headerMatch = headerCandidate.match(monthHeaderPattern);
      if (headerMatch) {
        currentYear = parseInt(headerMatch[2]);
        continue;
      }

      // Skip rows with no project name
      if (!col1) continue;

      // Parse the date cell — could be a Date object (date-formatted cell)
      // or a plain text string like "May 8", "TBC", "June 4-6"
      var rawDate = row[3];
      var dateInfo;
      if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
        var y  = rawDate.getFullYear();
        var mo = String(rawDate.getMonth() + 1).padStart(2, "0");
        var dy = String(rawDate.getDate()).padStart(2, "0");
        dateInfo = { date: y + "-" + mo + "-" + dy, dateEnd: "" };
      } else {
        dateInfo = parseDateFromSheet(col4, currentYear);
      }

      projects.push({
        id:      Utilities.getUuid(),
        project: col1,
        dri:     String(row[1]).trim(),
        team:    String(row[2]).trim(),
        date:    dateInfo.date,
        dateEnd: dateInfo.dateEnd,
        support: String(row[4]).trim(),
        notes:   String(row[5]).trim(),
        savedAt: new Date().toISOString(),
        synced:  true
      });
    }

    return jsonResponse({ status: "ok", projects: projects, sheetUsed: sheetUsed, allSheets: allSheetNames });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// Parse a date string from the sheet (e.g. "April 8", "April 8-10",
// "July 8-September 30", "TBC", "") back into ISO date strings.
function parseDateFromSheet(dateStr, year) {
  if (!dateStr || dateStr === "" ) return { date: "", dateEnd: "" };
  if (/^tbc$/i.test(dateStr))     return { date: "TBC", dateEnd: "" };

  // Remove trailing noise like "FULL", "TBC", "or 2", etc. — keep first date part
  // Normalise en-dashes and em-dashes to hyphens
  dateStr = dateStr.replace(/[–—]/g, "-").trim();

  // Try "Month Day-Month Day" range (e.g. "July 8-September 30")
  var crossMonthRange = dateStr.match(/^([A-Z][a-z]+)\s+(\d+)-([A-Z][a-z]+)\s+(\d+)/);
  if (crossMonthRange) {
    return {
      date:    toISO(crossMonthRange[1], crossMonthRange[2], year),
      dateEnd: toISO(crossMonthRange[3], crossMonthRange[4], year)
    };
  }

  // Try "Month Day-Day" same-month range (e.g. "April 8-10")
  var sameMonthRange = dateStr.match(/^([A-Z][a-z]+)\s+(\d+)-(\d+)/);
  if (sameMonthRange) {
    return {
      date:    toISO(sameMonthRange[1], sameMonthRange[2], year),
      dateEnd: toISO(sameMonthRange[1], sameMonthRange[3], year)
    };
  }

  // Try plain "Month Day" (e.g. "April 8")
  var single = dateStr.match(/^([A-Z][a-z]+)\s+(\d+)/);
  if (single) {
    return {
      date:    toISO(single[1], single[2], year),
      dateEnd: ""
    };
  }

  // Couldn't parse — store the raw string so data isn't lost
  return { date: dateStr, dateEnd: "" };
}

// Build an ISO date string "YYYY-MM-DD" from month name, day, and year
function toISO(monthName, day, year) {
  var mo = MONTHS.indexOf(monthName);
  if (mo === -1) return "";
  return year + "-" + String(mo + 1).padStart(2, "0") + "-" + String(parseInt(day)).padStart(2, "0");
}

// Main handler — receives a new project or a delete request from the web app.
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Handle delete action
    if (data.action === "delete") {
      return deleteProjectRow(data);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SHEET_NAME
      ? ss.getSheetByName(SHEET_NAME)
      : ss.getSheets()[0];

    if (!sheet) {
      return jsonResponse({ status: "error", message: "Sheet not found. Update SHEET_NAME in the script." });
    }

    // Format date for display in the sheet
    var dateDisplay = formatDate(data.date, data.dateEnd);

    // Row data — columns: PROJECT, DRI, TEAM, DATE, SUPPORT, MATERIALS/NOTES
    var rowData = [
      data.project || "",
      data.dri     || "",
      data.team    || "",
      dateDisplay,
      data.support || "",
      data.notes   || ""
    ];

    // Figure out which month this belongs to
    var label = getMonthLabel(data.date);

    // Find or create the month header row in the sheet
    var headerRow = findOrCreateMonthHeader(sheet, label);

    // Find where to insert within that month's block
    var insertAt = findInsertionRow(sheet, headerRow);

    // Insert a blank row and write the data
    if (insertAt <= sheet.getLastRow()) {
      sheet.insertRowAfter(insertAt - 1);
    }
    sheet.getRange(insertAt, 1, 1, rowData.length).setValues([rowData]);

    // Flush to minimise race conditions with simultaneous edits
    SpreadsheetApp.flush();

    return jsonResponse({ status: "ok", insertedAt: insertAt });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── Delete a project row by matching project name + DRI ──────────────────────
function deleteProjectRow(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
    if (!sheet) return jsonResponse({ status: "error", message: "Sheet not found." });

    var lastRow = sheet.getLastRow();
    if (lastRow === 0) return jsonResponse({ status: "ok", deleted: false });

    var values = sheet.getRange(1, 1, lastRow, 2).getValues();
    var targetProject = String(data.project || "").trim().toLowerCase();
    var targetDri     = String(data.dri     || "").trim().toLowerCase();

    for (var i = values.length - 1; i >= 0; i--) {
      var rowProject = String(values[i][0]).trim().toLowerCase();
      var rowDri     = String(values[i][1]).trim().toLowerCase();
      if (rowProject === targetProject && (!targetDri || rowDri === targetDri)) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return jsonResponse({ status: "ok", deleted: true, row: i + 1 });
      }
    }

    return jsonResponse({ status: "ok", deleted: false });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── Find or create a month header row ────────────────────────────────────────
function findOrCreateMonthHeader(sheet, label) {
  var lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    // Empty sheet — write the header in row 1
    writeMonthHeader(sheet, 1, label);
    return 1;
  }

  var colA = sheet.getRange(1, 1, lastRow, 1).getValues();

  // Look for the month label in column A
  for (var i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim() === label) {
      return i + 1; // 1-indexed
    }
  }

  // Not found — find the right chronological position and insert
  var insertBefore = findChronologicalPosition(colA, label);

  if (insertBefore <= lastRow) {
    sheet.insertRowBefore(insertBefore);
  }

  writeMonthHeader(sheet, insertBefore, label);
  return insertBefore;
}

// Write and style a month header row
function writeMonthHeader(sheet, row, label) {
  var range = sheet.getRange(row, 1, 1, 6);
  sheet.getRange(row, 1).setValue(label);
  range.setBackground("#C9DAF8");          // light blue, matches existing sheet
  range.setFontWeight("bold");
  range.setFontColor("#000000");
}

// Find the chronologically correct row to insert a new month header
function findChronologicalPosition(colA, newLabel) {
  var newKey = labelToKey(newLabel);

  for (var i = 0; i < colA.length; i++) {
    var cell = String(colA[i][0]).trim();
    var existingKey = labelToKey(cell);
    if (existingKey !== Infinity && existingKey > newKey) {
      return i + 1; // insert before this row
    }
  }

  return colA.length + 1; // append at end
}

// Convert "April 2026" → sortable number; "Upcoming / TBC" → Infinity
function labelToKey(label) {
  if (!label || label === "Upcoming / TBC") return Infinity;
  var parts = label.trim().split(" ");
  if (parts.length < 2) return Infinity;
  var mo = MONTHS.indexOf(parts[0]);
  var yr = parseInt(parts[1]);
  if (mo === -1 || isNaN(yr)) return Infinity;
  return yr * 12 + mo;
}

// ── Find where to insert within a month's block ───────────────────────────────
function findInsertionRow(sheet, headerRow) {
  var lastRow = sheet.getLastRow();
  var monthPattern = /^[A-Z][a-z]+ \d{4}$/;

  for (var r = headerRow + 1; r <= lastRow; r++) {
    var val = String(sheet.getRange(r, 1).getValue()).trim();
    // Stop at next month header or an empty cell that signals end of block
    if (monthPattern.test(val)) return r;
    if (val === "") return r;
  }

  return lastRow + 1;
}

// ── Date formatting ───────────────────────────────────────────────────────────
function formatDate(date, dateEnd) {
  if (!date || date === "TBC") return "TBC";

  function parse(s) {
    var p = s.split("-").map(Number);
    return new Date(p[0], p[1] - 1, p[2]);
  }

  var d = parse(date);
  var startStr = MONTHS[d.getMonth()] + " " + d.getDate();

  if (dateEnd) {
    var de = parse(dateEnd);
    var endStr;
    if (de.getMonth() === d.getMonth() && de.getFullYear() === d.getFullYear()) {
      endStr = String(de.getDate()); // "May 8-10"
    } else {
      endStr = MONTHS[de.getMonth()] + " " + de.getDate(); // "July 8-September 30"
    }
    return startStr + "-" + endStr;
  }

  return startStr;
}

// ── Month label from ISO date ─────────────────────────────────────────────────
function getMonthLabel(date) {
  if (!date || date === "TBC") return "Upcoming / TBC";
  var p = date.split("-").map(Number);
  var d = new Date(p[0], p[1] - 1, p[2]);
  return MONTHS[d.getMonth()] + " " + d.getFullYear();
}

// ── JSON response helper ──────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
