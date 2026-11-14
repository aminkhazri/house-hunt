/**
 * House Hunt API — Google Apps Script web app.
 *
 * Turns one tab of a Google Sheet ("Homes") into a tiny JSON API:
 *   GET  ?action=list                       -> { ok, homes: [...] }
 *   GET  ?action=csv                        -> our current list as CSV (for Excel)
 *   GET  ?action=csv&include=removed        -> same, plus removed homes
 *   POST { action: "add",     home, passcode }    -> { ok, home }
 *   POST { action: "remove",  id,   passcode }    -> { ok }
 *   POST { action: "restore", id,   passcode }    -> { ok }
 *
 * Reading is open to anyone with the URL; writing needs the passcode stored
 * in Script Properties (key: PASSCODE). Removing never deletes a row — it
 * just sets removed = TRUE, so every change can be undone from the sheet.
 *
 * Setup: run setup() once from the editor (see README).
 */

const SHEET_NAME = "Homes";
// Where setup() downloads the starter list from (the same file the site uses).
const SEED_URL = "https://raw.githubusercontent.com/aminkhazri/house-hunt/main/explorer/data/homes.json";
const COLUMNS = [
  "id", "address", "city", "zip", "county", "neighborhood", "price", "originalPrice",
  "status", "mls", "daysOnMarket", "mortgage", "safety", "yearBuilt", "sqft", "beds",
  "baths", "hoa", "notes", "elementary", "elementaryScore", "middle", "middleScore",
  "high", "highScore", "list", "redfinUrl", "photoUrl", "addedBy", "addedAt",
  "removed", "removedAt",
  // Family page columns: rank 1 = top choice; blank rank = not on the top list.
  "rank", "why", "watch", "deadline",
];
const NUMBER_FIELDS = ["price", "mortgage", "yearBuilt", "sqft", "beds", "baths", "hoa",
  "elementaryScore", "middleScore", "highScore", "rank"];

/* ---------- One-time setup ---------- */

/**
 * Creates the Homes tab, fills it with the starter list (Seed.gs if present,
 * otherwise data/homes.json from the GitHub repo) and makes
 * a passcode if you haven't set one. Safe to run again: it never overwrites
 * an existing Homes tab.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    Logger.log('A "%s" tab already exists, so the starter list was not re-imported.', SHEET_NAME);
  } else {
    sheet = ss.insertSheet(SHEET_NAME, 0);
    // Store the label columns as plain text so Sheets doesn't turn them into dates.
    ["list", "zip", "originalPrice", "daysOnMarket", "mls"].forEach((key) =>
      sheet.getRange(1, COLUMNS.indexOf(key) + 1, sheet.getMaxRows(), 1).setNumberFormat("@"));
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS])
      .setFontWeight("bold").setBackground("#dcefe6");
    sheet.setFrozenRows(1);
    // Seed.gs is optional: paste it in to work offline, otherwise download from GitHub.
    const seed = typeof SEED_HOMES !== "undefined"
      ? SEED_HOMES
      : JSON.parse(UrlFetchApp.fetch(SEED_URL).getContentText());
    const rows = seed.map((home) => toRow_(Object.assign({}, home, { removed: false })));
    if (rows.length) sheet.getRange(2, 1, rows.length, COLUMNS.length).setValues(rows);
    sheet.autoResizeColumns(1, 6);
    Logger.log("Created %s with %s homes.", SHEET_NAME, rows.length);
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("PASSCODE")) {
    const words = ["porch", "maple", "cozy", "garden", "sunny", "cedar", "brick", "acorn"];
    const code = words[Math.floor(Math.random() * words.length)] + "-" + Math.floor(1000 + Math.random() * 9000);
    props.setProperty("PASSCODE", code);
    Logger.log("Your family passcode is: %s  (change it in Project Settings → Script Properties)", code);
  } else {
    Logger.log("Passcode already set. See Project Settings → Script Properties.");
  }
}

/* ---------- HTTP entry points ---------- */

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = p.action || "list";
    if (action === "csv") return csv_(readHomes_(), p.include === "removed");
    if (action !== "list") return json_({ ok: false, error: "Unknown action." });
    return json_({ ok: true, homes: readHomes_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const expected = PropertiesService.getScriptProperties().getProperty("PASSCODE");
    if (!expected || String(body.passcode || "") !== expected) {
      return json_({ ok: false, error: "bad_passcode" });
    }

    lock.waitLock(10000); // one write at a time, so two reviewers can't collide
    switch (body.action) {
      case "add": return json_({ ok: true, home: addHome_(body.home) });
      case "remove": setRemoved_(body.id, true); return json_({ ok: true });
      case "restore": setRemoved_(body.id, false); return json_({ ok: true });
      default: return json_({ ok: false, error: "Unknown action." });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Sheet helpers ---------- */

function sheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('No "Homes" tab yet. Run setup() in the Apps Script editor.');
  ensureColumns_(sheet);
  return sheet;
}

/** Add any columns from COLUMNS that an older Homes tab doesn't have yet (in order, at the end). */
function ensureColumns_(sheet) {
  const width = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, width).getValues()[0];
  const missing = COLUMNS.filter((c) => header.indexOf(c) < 0);
  if (missing.length) {
    sheet.getRange(1, width + 1, 1, missing.length).setValues([missing])
      .setFontWeight("bold").setBackground("#fff6d1");
  }
}

/** 1-based column number for a header name. */
function col_(sheet, name) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const i = header.indexOf(name);
  if (i < 0) throw new Error("Missing column: " + name);
  return i + 1;
}

function readHomes_() {
  const values = sheet_().getDataRange().getValues();
  const header = values.shift();
  return values
    .filter((row) => row[0] !== "")
    .map((row) => {
      const home = {};
      header.forEach((key, i) => {
        let v = row[i];
        // Sheets auto-converts text like "October 2026" or "Added Nov 14" into dates.
        // Turn the list label back into words; keep other dates as ISO strings.
        if (v instanceof Date && key === "list") {
          const tz = Session.getScriptTimeZone();
          v = v.getDate() === 1
            ? Utilities.formatDate(v, tz, "MMMM yyyy")
            : "Added " + Utilities.formatDate(v, tz, "MMM d");
        } else if (v instanceof Date) {
          v = v.toISOString();
        }
        if (key === "removed") v = v === true || String(v).toUpperCase() === "TRUE";
        else if (NUMBER_FIELDS.indexOf(key) >= 0) v = v === "" ? null : Number(v);
        else v = v === null ? "" : String(v);
        home[key] = v;
      });
      return home;
    });
}

function addHome_(input) {
  const home = cleanHome_(input);
  const sheet = sheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
  if (ids.indexOf(home.id) >= 0) throw new Error("That home is already on the list.");
  sheet.appendRow(toRow_(home));
  return home;
}

function setRemoved_(id, removed) {
  const sheet = sheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
  const index = ids.indexOf(String(id));
  if (index < 0) throw new Error("Couldn't find that home. Refresh the page and try again.");
  const row = index + 2;
  sheet.getRange(row, col_(sheet, "removed")).setValue(removed);
  sheet.getRange(row, col_(sheet, "removedAt")).setValue(removed ? new Date().toISOString() : "");
}

/** Never trust the browser: trim text, coerce numbers, allow only http(s) links. */
function cleanHome_(input) {
  const h = input || {};
  const text = (v, max) => String(v == null ? "" : v).trim().slice(0, max || 200);
  const num = (v) => (v === "" || v == null || isNaN(Number(v)) ? null : Number(v));
  const url = (v) => (/^https?:\/\//i.test(text(v, 500)) ? text(v, 500) : "");

  const address = text(h.address, 120);
  const zip = text(h.zip, 5);
  if (!address) throw new Error("Address is required.");
  if (!/^\d{5}$/.test(zip)) throw new Error("Zip should be 5 digits.");

  return {
    id: (address + "-" + zip).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, ""),
    address: address, city: text(h.city, 60), zip: zip, county: text(h.county, 60),
    neighborhood: text(h.neighborhood, 80), price: num(h.price), originalPrice: "",
    status: "Active", mls: text(h.mls, 40), daysOnMarket: "", mortgage: num(h.mortgage),
    safety: text(h.safety, 3), yearBuilt: num(h.yearBuilt), sqft: num(h.sqft),
    beds: num(h.beds), baths: num(h.baths), hoa: num(h.hoa), notes: text(h.notes, 1000),
    elementary: "", elementaryScore: null, middle: "", middleScore: null, high: "", highScore: null,
    list: text(h.list, 40), redfinUrl: url(h.redfinUrl), photoUrl: url(h.photoUrl),
    addedBy: text(h.addedBy, 60) || "A reviewer", addedAt: new Date().toISOString(),
    removed: false, removedAt: "",
  };
}

function toRow_(home) {
  return COLUMNS.map((key) => (home[key] == null ? "" : home[key]));
}

/**
 * Our list as CSV, ranked homes first (1, 2, 3...), then the rest by price.
 * Excel's Data → From Web reads this URL and can refresh it automatically.
 */
function csv_(homes, includeRemoved) {
  const list = homes
    .filter((h) => includeRemoved || !h.removed)
    .sort((a, b) => (a.rank || 999) - (b.rank || 999) || (a.price || 0) - (b.price || 0));
  const head = ["Rank", "Address", "Town", "Zip", "Price", "Beds", "Baths", "Sq Ft", "Year Built",
    "Safety", "Status", "Why we like it", "Keep in mind", "Deadline", "HOA $/mo",
    "Elementary", "Middle", "High", "Redfin", "On our list"];
  const cell = (v) => {
    const s = v == null ? "" : String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const rows = list.map((h) => [
    h.rank || "", h.address, h.city, h.zip, h.price, h.beds, h.baths, h.sqft, h.yearBuilt,
    h.safety, h.status, h.why, h.watch, h.deadline, h.hoa,
    h.elementary, h.middle, h.high, h.redfinUrl, h.removed ? "Removed" : "Yes",
  ].map(cell).join(","));
  return ContentService.createTextOutput([head.join(",")].concat(rows).join("\n"))
    .setMimeType(ContentService.MimeType.CSV);
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
