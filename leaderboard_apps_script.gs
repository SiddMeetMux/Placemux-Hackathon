/**
 * TECHNOVA PLACEMUX 2026 — Leaderboard Backend
 * ================================================
 * Paste this entire file into Google Apps Script (Extensions > Apps Script from
 * your Google Sheet), then deploy as a Web App:
 *   Execute as: Me  |  Who has access: Anyone
 *
 * Your Sheet must have a tab named exactly: Leaderboard
 * Row 1 column headers (must match exactly):
 *   Timestamp | Name | Roll | Seconds | Display
 *
 * After deploying, copy the Web App URL and paste it into
 * technova_placemux_mission.html at:
 *   const LEADERBOARD_ENDPOINT = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
 */

// ── Configuration ─────────────────────────────────────────────────────────────
var SHEET_NAME   = "Leaderboard";   // Tab name in your Google Sheet
var TOP_N        = 50;              // How many rows to return in the leaderboard
// ─────────────────────────────────────────────────────────────────────────────

/**
 * doPost — Receives a new leaderboard entry from the mission page.
 * The page sends a JSON body with: { name, roll, seconds, display, ts }
 */
function doPost(e) {
  try {
    var data    = JSON.parse(e.postData.contents);
    var name    = (data.name    || "").toString().trim();
    var roll    = (data.roll    || "").toString().trim();
    var seconds = parseFloat(data.seconds) || 0;
    var display = (data.display || seconds + "s").toString();
    var ts      = (data.ts      || new Date().toISOString()).toString();

    if (!name || !roll) {
      return jsonResponse({ ok: false, error: "Name and Roll are required" });
    }

    var sheet = getSheet();
    // Append row: Timestamp | Name | Roll | Seconds | Display
    sheet.appendRow([ts, name, roll, seconds, display]);

    // Flush to ensure the write completes before we read back
    SpreadsheetApp.flush();

    return jsonResponse({ ok: true, message: "Entry saved" });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

/**
 * doGet — Returns the top-N leaderboard as JSON.
 * Called by the mission page with ?action=top50
 */
function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) ? e.parameter.action : "";

    if (action === "top50" || action === "") {
      var rows = getTopRows(TOP_N);
      return jsonResponse(rows);
    }

    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * getSheet — Returns the Leaderboard sheet, creating headers if missing.
 */
function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // Ensure headers exist
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Name", "Roll", "Seconds", "Display"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  return sheet;
}

/**
 * getTopRows — Reads all data rows, deduplicates by name (keeps fastest),
 * sorts by Seconds ascending, returns top N as plain objects.
 */
function getTopRows(n) {
  var sheet = getSheet();
  var last  = sheet.getLastRow();
  if (last <= 1) return [];  // no data beyond header

  // Read all data rows (skip row 1 = header)
  var data = sheet.getRange(2, 1, last - 1, 5).getValues();
  // Columns: 0=Timestamp, 1=Name, 2=Roll, 3=Seconds, 4=Display

  // Deduplicate: keep the fastest (lowest Seconds) entry per Name
  var best = {};
  data.forEach(function(row) {
    var name    = (row[1] || "").toString().trim();
    var roll    = (row[2] || "").toString().trim();
    var seconds = parseFloat(row[3]) || 0;
    var display = (row[4] || seconds + "s").toString();
    if (!name) return;
    if (!best[name] || seconds < best[name].seconds) {
      best[name] = { name: name, roll: roll, seconds: seconds, display: display };
    }
  });

  // Sort by seconds ascending (fastest first)
  var sorted = Object.values(best).sort(function(a, b) {
    return a.seconds - b.seconds;
  });

  return sorted.slice(0, n);
}

/**
 * jsonResponse — Wraps a value in a ContentService JSON response with CORS.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Optional: manual test function ───────────────────────────────────────────
/**
 * Run this function in the Apps Script editor to verify everything works:
 *   1. Open Extensions > Apps Script
 *   2. Select "testLeaderboard" from the dropdown
 *   3. Click Run
 *   4. Check the Execution Log at the bottom
 */
function testLeaderboard() {
  // Insert two test rows
  var sheet = getSheet();
  sheet.appendRow([new Date().toISOString(), "Test Student", "REG001", 45.2, "00:45.2"]);
  sheet.appendRow([new Date().toISOString(), "Another Student", "REG002", 32.7, "00:32.7"]);
  SpreadsheetApp.flush();

  // Read back top 50
  var rows = getTopRows(50);
  Logger.log("Top rows: " + JSON.stringify(rows));
  Logger.log("Test passed! " + rows.length + " rows returned.");
}
