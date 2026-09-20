/**
 * TECHNOVA PLACEMUX 2026 — LEADERBOARD APPS SCRIPT BACKEND
 * =========================================================
 * Deploys as a Google Apps Script Web App backed by Google Sheets.
 * Stores Name, Roll No / College, Elapsed Time, Display Time, and Mission Code.
 * Serves live ranked Top 50, Top 100, and All participant results with deduplication.
 *
 * HOW TO DEPLOY:
 * 1. In your Google Sheet, click Extensions > Apps Script.
 * 2. Paste this entire code into the editor (replace any existing code).
 * 3. Click "Deploy" > "New deployment".
 * 4. Select type "Web app":
 *    - Description: Leaderboard v1
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click "Deploy", authorize access, and copy the Web App URL.
 * 6. Paste that URL into technova_placemux_mission.html and admin_leaderboard.html.
 */

const SHEET_NAME = "Leaderboard";
const HEADERS = ["Timestamp", "Name", "Roll", "Seconds", "Display", "Code"];

/**
 * Ensures the target sheet and header row exist.
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#1267e8").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Handles incoming POST requests (participant submitting their final time & details)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  // Wait up to 10 seconds for concurrent submissions to prevent row overwrite collisions
  const hasLock = lock.tryLock(10000);
  
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    const name = String(data.name || "").trim();
    const roll = String(data.roll || "").trim();
    const seconds = parseFloat(data.seconds) || 0;
    const display = String(data.display || "").trim();
    const code = String(data.code || "").trim();
    const ts = data.ts || new Date().toISOString();

    if (!name || !roll || seconds <= 0) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Missing required fields (name, roll, or valid seconds)"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = getOrCreateSheet();
    sheet.appendRow([ts, name, roll, seconds, display, code]);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Result saved to leaderboard",
      entry: { name, roll, seconds, display, code, ts }
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Handles incoming GET requests (fetching Top 50, Top 100, or All records)
 * Examples:
 *   ?action=top50
 *   ?action=top100
 *   ?action=all
 *   &dedupe=false (to see every attempt rather than only each student's fastest)
 */
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    }

    // Read all rows starting from row 2 (skip headers)
    const range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    const rawValues = range.getValues();

    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action.toLowerCase() : "top50";
    const dedupeParam = (e && e.parameter && e.parameter.dedupe) ? e.parameter.dedupe.toLowerCase() : "true";
    const shouldDedupe = dedupeParam !== "false";

    let rows = rawValues.map((r, idx) => {
      return {
        rowId: idx + 2,
        timestamp: r[0] ? new Date(r[0]).toISOString() : "",
        name: String(r[1] || "").trim(),
        roll: String(r[2] || "").trim(),
        seconds: parseFloat(r[3]) || 999999,
        display: String(r[4] || "").trim(),
        code: String(r[5] || "").trim()
      };
    }).filter(r => r.name && r.roll && r.seconds < 999999);

    // Optional Deduplication: Keep the fastest time for each student (grouped by roll number)
    if (shouldDedupe) {
      const bestMap = new Map();
      rows.forEach(item => {
        const key = item.roll.toLowerCase();
        if (!bestMap.has(key) || item.seconds < bestMap.get(key).seconds) {
          bestMap.set(key, item);
        }
      });
      rows = Array.from(bestMap.values());
    }

    // Sort by seconds ascending (fastest first)
    rows.sort((a, b) => a.seconds - b.seconds);

    // Assign rank
    rows.forEach((item, index) => {
      item.rank = index + 1;
    });

    // Limit based on action
    let result = rows;
    if (action === "top50") {
      result = rows.slice(0, 50);
    } else if (action === "top100") {
      result = rows.slice(0, 100);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
