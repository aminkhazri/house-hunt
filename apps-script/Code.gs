/**
 * House Hunt API — Google Apps Script web app.
 *
 * Turns one tab of a Google Sheet ("Homes") into a tiny JSON API:
 *   GET  ?action=list                       -> { ok, homes: [...] }
 *   GET  ?action=csv                        -> our current list as CSV (for Excel)
 *   GET  ?action=csv&include=removed        -> same, plus removed homes
 *   GET  ?action=lookup&url=REDFIN_URL      -> { ok, home } preview from a Redfin link (web or app short link)
 *   POST { action: "add",     home }              -> { ok, home }
 *   POST { action: "addRedfin", url, why, addedBy }  -> { ok, home }
 *   POST { action: "remove",  id, reason, passcode } -> { ok }   (reason is saved in removedReason)
 *   POST { action: "move", id, direction: "up"|"down" } -> { ok, order: [ids] }
 *   POST { action: "moveTo", id, position: 1..n }       -> { ok, order: [ids] }
 *   POST { action: "restore", id }                -> { ok }
 *   POST { action: "comment", id, name, text }    -> { ok, comment }  (saved in the Comments tab)
 *   POST { action: "vote", id, name, hearts: 1..5 } -> { ok, vote }   (one vote per person per home, Votes tab)
 *   POST { action: "upload", homeId, name, mime, data(base64) } -> { ok, file: { id, type } }  (tour photo / voice note, saved in Drive)
 *   POST { action: "tour", tour: { id, homeId, name, answers, notes, files } } -> { ok, tour }  (Tours tab)
 *   POST { action: "radar", url, decision: "skip" }  -> { ok }   (hide a "New today" match)
 *   POST { action: "settings", downPct, rate }    -> { ok, settings }  (the monthly-cost numbers)
 *
 * The morning job is morningCheck() (run setupTriggers() once to schedule it). It re-reads every
 * listing (price, status, open houses, days on market…), looks for new matches, works out
 * drive times, and emails the family a short summary when something changed.
 *
 * Reading and most changes are open to anyone with the URL; only "remove" needs the passcode stored
 * in Script Properties (key: PASSCODE). Removing never deletes a row — it
 * just sets removed = TRUE, so every change can be undone from the sheet.
 *
 * Setup: run setup() once from the editor (see README).
 */

const SHEET_NAME = "Homes";
const COMMENTS_SHEET = "Comments";
const COMMENT_COLUMNS = ["id", "homeId", "name", "text", "at"];
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
  "removedReason",
  // Filled in by the daily Redfin check: the price before the latest change, and when it changed.
  "lastPrice", "priceChangedAt",
  // Also filled in every morning from the Redfin listing page.
  "lat", "lng", "garage", "dom", "taxRate", "insurance", "drops", "openHouses", "drives", "statusChangedAt",
];
const NUMBER_FIELDS = ["price", "mortgage", "yearBuilt", "sqft", "beds", "baths", "hoa",
  "elementaryScore", "middleScore", "highScore", "rank", "lastPrice",
  "lat", "lng", "garage", "dom", "taxRate", "insurance", "drops"];

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
    if (action === "csv") { const hs = readHomes_(); return csv_(hs, p.include === "removed", readComments_(), safetyFor_(hs.filter((h) => h.rank).map((h) => h.zip))); }
    if (action === "lookup") return json_({ ok: true, home: redfinLookup_(p.url) });
    if (action !== "list") return json_({ ok: false, error: "Unknown action." });
    return ContentService.createTextOutput(listJson_()).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    // Uploads only write to Drive, not the Sheet, so they don't wait for the lock.
    if (body.action === "upload") return json_({ ok: true, file: saveUpload_(body) });
    const expected = PropertiesService.getScriptProperties().getProperty("PASSCODE");
    // Only removing a home needs the family passcode. Adding, moving, commenting and
    // bringing a home back are open to anyone with the page (all easy to undo).
    const open = body.action !== "remove";
    if (!open && (!expected || String(body.passcode || "") !== expected)) {
      return json_({ ok: false, error: "bad_passcode" });
    }

    lock.waitLock(10000); // one write at a time, so two reviewers can't collide
    clearListCache_();    // something is about to change, so the saved copy of the list is out of date
    switch (body.action) {
      case "add": return json_({ ok: true, home: addHome_(body.home) });
      case "addRedfin": {
        const home = addFromRedfin_(body);
        decideRadar_(home.redfinUrl, "added"); // if it came from the "New today" box, take it out of there
        return json_({ ok: true, home: home });
      }
      case "move": return json_({ ok: true, order: moveHome_(body.id, body.direction) });
      case "moveTo": return json_({ ok: true, order: moveHome_(body.id, null, Number(body.position)) });
      case "remove": setRemoved_(body.id, true, body.reason); return json_({ ok: true });
      case "restore": setRemoved_(body.id, false); return json_({ ok: true });
      case "comment": return json_({ ok: true, comment: addComment_(body) });
      case "vote": return json_({ ok: true, vote: saveVote_(body) });
      case "tour": return json_({ ok: true, tour: saveTour_(body.tour) });
      case "radar": return json_({ ok: true, radar: decideRadar_(body.url, body.decision) });
      case "settings": return json_({ ok: true, settings: saveCostSettings_(body) });
      default: return json_({ ok: false, error: "Unknown action." });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Fast list: a saved copy of the answer ---------- */

const LIST_CACHE_KEY = "LIST_JSON";
const LIST_CACHE_SECONDS = 300; // 5 minutes; any change made through the page clears it right away

/**
 * The family page's list as JSON text. Reading the Sheet, the comments and the safety
 * grades takes a while, so we keep the finished answer in CacheService and reuse it
 * until something changes. Only homes that are (or were) on our ranked list are sent.
 */
function listJson_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(LIST_CACHE_KEY);
  if (hit) return hit;
  const homes = readHomes_().filter((h) => h.rank);
  const ids = homes.map((h) => h.id);
  const mine = (x) => ids.indexOf(x.homeId) >= 0;
  const props = PropertiesService.getScriptProperties();
  const json = JSON.stringify({
    ok: true,
    homes: homes,
    comments: readComments_().filter(mine),
    safety: safetyFor_(homes.map((h) => h.zip)),
    pricesChecked: props.getProperty("PRICES_CHECKED") || "",
    votes: readTab_(VOTES_TAB, VOTE_COLUMNS).filter(mine),
    tours: readTab_(TOURS_TAB, TOUR_COLUMNS).filter(mine).map((t) => Object.assign(t, {
      answers: parseJson_(t.answers, {}), files: parseJson_(t.files, []),
    })),
    radar: radarForPage_(),
    market: parseJson_(props.getProperty("MARKET"), {}),
    settings: publicSettings_(),
  });
  if (json.length < 90000) cache.put(LIST_CACHE_KEY, json, LIST_CACHE_SECONDS); // cache values max out at 100 KB
  return json;
}

function clearListCache_() {
  CacheService.getScriptCache().remove(LIST_CACHE_KEY);
}

/** Runs by itself whenever someone edits the Google Sheet by hand, so the page shows the edit right away. */
function onEdit() {
  clearListCache_();
  CacheService.getScriptCache().remove("SETTINGS"); // in case the Settings tab was edited
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

function setRemoved_(id, removed, reason) {
  const sheet = sheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
  const index = ids.indexOf(String(id));
  if (index < 0) throw new Error("Couldn't find that home. Refresh the page and try again.");
  const row = index + 2;
  sheet.getRange(row, col_(sheet, "removed")).setValue(removed);
  sheet.getRange(row, col_(sheet, "removedAt")).setValue(removed ? new Date().toISOString() : "");
  sheet.getRange(row, col_(sheet, "removedReason")).setValue(removed ? String(reason || "").trim().slice(0, 500) : "");
}

/* ---------- Safety grades from CrimeGrade (by zip code) ---------- */

const SAFETY_DAYS = 7; // look each zip up again after a week

/** CrimeGrade's page for a zip code, e.g. https://crimegrade.org/safest-places-in-27527/ */
function crimeGradeUrl_(zip) {
  return "https://crimegrade.org/safest-places-in-" + zip + "/";
}

/**
 * Grades for each zip code: { "27527": { grade, score, saferThan, violent, property, url, checked } }.
 * Saved in Script Properties so we only ask CrimeGrade about a zip once a week.
 */
function safetyFor_(zips) {
  const props = PropertiesService.getScriptProperties();
  const out = {};
  const unique = zips.filter((z, i) => /^\d{5}$/.test(z) && zips.indexOf(z) === i);
  unique.forEach((zip) => {
    const key = "SAFETY_" + zip;
    let info = null;
    try { info = JSON.parse(props.getProperty(key) || "null"); } catch (e) { info = null; }
    const stale = !info || Date.now() - new Date(info.checked).getTime() > SAFETY_DAYS * 864e5;
    if (stale) {
      try {
        const fresh = crimeGradeLookup_(zip);
        if (fresh) { info = fresh; props.setProperty(key, JSON.stringify(info)); }
      } catch (e) { /* keep the old numbers if CrimeGrade is down */ }
    }
    if (info) out[zip] = info;
  });
  return out;
}

/** Read the zip's CrimeGrade page and pull out the grades and safety scores. */
function crimeGradeLookup_(zip) {
  const res = UrlFetchApp.fetch(crimeGradeUrl_(zip), {
    muteHttpExceptions: true,
    headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36" },
  });
  if (res.getResponseCode() !== 200) return null;
  // Turn the HTML into plain text: drop styles/scripts, then tags, then squeeze the spaces.
  const text = res.getContentText()
    .replace(/<(style|script)[\s\S]*?<\/\1>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&trade;|&#8482;/g, "™").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&")
    .replace(/\s+/g, " ");
  // e.g. "its overall CrimeGrade™ is A, safer than the neighborhoods where 72% of Americans live.
  //       Violent crime grades A+ and property crime A."  or  "...Violent and property crime both grade A here."
  const overall = text.match(/overall CrimeGrade\S* (?:is|of) ([A-F][+-]?)/);
  if (!overall) return null;
  const pct = text.match(/where (\d+)% of Americans live/);
  const both = text.match(/Violent and property crime both grade ([A-F][+-]?)/);
  const split = text.match(/Violent crime grades ([A-F][+-]?) and property crime ([A-F][+-]?)/);
  const scores = [];
  const re = /Safety score (\d+) of 1,000/g;
  let m;
  while ((m = re.exec(text)) && scores.length < 3) scores.push(Number(m[1]));
  return {
    grade: overall[1],
    score: scores[0] || null,                       // 0-1000, higher is safer
    saferThan: pct ? Number(pct[1]) : null,         // % of Americans' neighborhoods this zip is safer than
    violent: both ? both[1] : split ? split[1] : "",
    property: both ? both[1] : split ? split[2] : "",
    url: crimeGradeUrl_(zip),
    checked: new Date().toISOString(),
  };
}

/* ---------- Comments ---------- */

/** The Comments tab (one row per comment). Created the first time someone comments. */
function commentsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(COMMENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(COMMENTS_SHEET);
    sheet.getRange(1, 1, 1, COMMENT_COLUMNS.length).setValues([COMMENT_COLUMNS])
      .setFontWeight("bold").setBackground("#e8f1fd");
    sheet.setFrozenRows(1);
    sheet.getRange(1, 4, sheet.getMaxRows(), 1).setNumberFormat("@"); // keep comments as plain text
  }
  return sheet;
}

/** Every comment, oldest first. Rows you delete in the Sheet simply disappear from the page. */
function readComments_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMMENTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, COMMENT_COLUMNS.length).getValues()
    .filter((r) => r[1] !== "" && r[3] !== "")
    .map((r) => ({
      id: String(r[0]), homeId: String(r[1]), name: String(r[2]), text: String(r[3]),
      at: r[4] instanceof Date ? r[4].toISOString() : String(r[4]),
    }));
}

function addComment_(body) {
  const text = String(body.text || "").trim().slice(0, 500);
  const name = String(body.name || "").trim().slice(0, 40) || "Someone";
  if (!text) throw new Error("Please write a comment first.");
  const sheet = sheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat().map(String);
  if (ids.indexOf(String(body.id)) < 0) throw new Error("Couldn't find that home. Refresh the page and try again.");
  const comment = { id: Utilities.getUuid().slice(0, 8), homeId: String(body.id), name: name, text: text, at: new Date().toISOString() };
  commentsSheet_().appendRow(COMMENT_COLUMNS.map((k) => comment[k]));
  return comment;
}

/* ---------- Reorder ---------- */

/**
 * Move a home one spot up or down our ranked list, then renumber the list
 * 1, 2, 3... so gaps left by removed homes disappear. Returns the new order (ids).
 */
function moveHome_(id, direction, position) {
  const sheet = sheet_();
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const cId = header.indexOf("id"), cRank = header.indexOf("rank"), cRemoved = header.indexOf("removed");
  const isRemoved = (v) => v === true || String(v).toUpperCase() === "TRUE";

  const ranked = [];
  for (let r = 1; r < values.length; r++) {
    const rank = Number(values[r][cRank]);
    if (rank > 0 && !isRemoved(values[r][cRemoved])) ranked.push({ row: r, id: String(values[r][cId]), rank: rank });
  }
  ranked.sort((a, b) => a.rank - b.rank);

  const i = ranked.findIndex((h) => h.id === String(id));
  if (i < 0) throw new Error("Couldn't find that home on our list. Refresh the page and try again.");
  let j;
  if (position) {
    // Drag and drop: take the home out and put it at the requested spot (1 = top).
    j = Math.max(0, Math.min(ranked.length - 1, Math.round(position) - 1));
  } else {
    j = direction === "up" ? i - 1 : i + 1;
    if (j < 0) throw new Error("It's already our top choice.");
    if (j >= ranked.length) throw new Error("It's already at the bottom of the list.");
  }
  const [home] = ranked.splice(i, 1);
  ranked.splice(j, 0, home);

  ranked.forEach((h, n) => sheet.getRange(h.row + 1, cRank + 1).setValue(n + 1));
  return ranked.map((h) => h.id);
}

/* ---------- Add from a Redfin link ---------- */

const REDFIN_URL = /^https:\/\/(www\.)?redfin\.com\/[A-Z]{2}\/[^?#\s]+\/home\/\d+/;

/**
 * The Redfin phone app shares short links like https://redf.in/rFnamM, often inside
 * a sentence ("Check out this home I found on Redfin! https://redf.in/..."). Pull the
 * link out of whatever was pasted, and if it's a short link, follow its redirects
 * (one hop at a time, reading each Location header) until we reach redfin.com.
 */
function resolveShortLink_(pasted) {
  const found = String(pasted || "").match(/https?:\/\/[^\s"'<>]+/i);
  let url = found ? found[0] : String(pasted || "").trim();
  if (/^(www\.)?redf\.in\//i.test(url)) url = "https://" + url; // pasted without https://
  for (let hop = 0; hop < 5 && /^https?:\/\/(www\.)?redf\.in\//i.test(url); hop++) {
    const res = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
    const headers = res.getAllHeaders();
    const next = headers.Location || headers.location;
    if (!next) break;
    url = String(Array.isArray(next) ? next[0] : next);
  }
  return url.replace(/^http:/, "https:");
}

/**
 * Read a Redfin listing page and pull out the basics. Redfin's page summary
 * (meta description) looks like:
 *   "For Sale: 5 beds, 3.5 baths · 2339 sq. ft. · 479 Curling Crk Dr, Clayton, NC 27527 · $385,000 · MLS# 10193225 · ..."
 */
function redfinLookup_(rawUrl) {
  const match = resolveShortLink_(rawUrl).match(REDFIN_URL);
  if (!match) throw new Error("That doesn't look like a Redfin home link. In the Redfin app, tap Share, then Copy link, and paste it here.");
  const url = match[0];

  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36" },
  });
  if (res.getResponseCode() !== 200) throw new Error("Redfin didn't answer. Please try again in a minute.");
  const html = res.getContentText();
  const meta = (name) => {
    const m = html.match(new RegExp('<meta[^>]+(?:name|property)="' + name + '"[^>]+content="([^"]*)"'));
    return m ? decode_(m[1]) : "";
  };
  const desc = meta("description");
  const num = (re) => { const m = desc.match(re); return m ? Number(m[1].replace(/,/g, "")) : null; };

  // Address: from the summary if possible, otherwise from the link itself.
  const addr = desc.match(/([^·∙:]+?),\s*([^,·∙]+),\s*([A-Z]{2})\s+(\d{5})/);
  const path = url.split("/"); // ["https:", "", "www.redfin.com", "NC", "Clayton", "479-Curling-Creek-Dr-27527", "home", "187212408"]
  const slugParts = path[5].split("-");
  const zip = addr ? addr[4] : slugParts.pop();
  const street = addr ? addr[1].trim() : slugParts.join(" ");
  const city = addr ? addr[2].trim() : path[4].replace(/-/g, " ");
  const year = html.match(/\\?"yearBuilt\\?"\s*:\s*(\d{4})/);
  // Status: the summary starts "For Sale:" or "Sold:" (nothing = off the market), and the
  // page's own data says whether a for-sale home is already pending.
  const prefix = (desc.match(/^([A-Za-z ]+):/) || [])[1] || "";
  const listing = (html.match(/"listingStatus"\s*:\s*"([a-z_ ]+)"/i) || [])[1] || "";
  const status = /^sold/i.test(prefix) ? "Sold"
    : !/sale/i.test(prefix) ? "Off market"
    : /pending|contingent/i.test(listing) ? "Pending" : "Active";

  // Details from the data Redfin puts inside the page. Inside that data the quotes are
  // often escaped (\"name\":value), so every pattern allows an optional backslash.
  const field = (name, pattern) => {
    const m = html.match(new RegExp(name + '\\\\?"\\s*:\\s*' + (pattern || "(-?[\\d.]+)")));
    return m ? Number(m[1]) : null;
  };
  const garage = html.match(/(\d+) garage spaces?/i) || html.match(/Parking: (\d+) car garage/i);
  // Upcoming open houses: [{start, end}] as times in milliseconds.
  const openHouses = [];
  const oh = html.match(/openHouses\\?"\s*:\s*\[([^\]]*)\]/);
  if (oh) {
    const re = /startTime\\?"\s*:\s*(\d+)\s*,\s*\\?"endTime\\?"\s*:\s*(\d+)/g;
    let m;
    while ((m = re.exec(oh[1]))) if (Number(m[2]) > Date.now()) openHouses.push({ start: Number(m[1]), end: Number(m[2]) });
  }

  const school = elementarySchool_(html);

  return {
    elementary: school ? school.name : "", elementaryScore: school ? school.rating : null,
    lat: field("latitude"), lng: field("longitude"),
    garage: garage ? Number(garage[1]) : null,
    dom: field("cumulativeDaysOnMarket", "(\\d+)"),
    hoa: field("monthlyHoaDues"),
    taxRate: field("propertyTaxRate"),            // % of the price per year, Redfin's estimate
    insurance: field("homeInsuranceYearlyEstimate", "(\\d+)"), // $ per year, Redfin's estimate
    rate30: field("thirtyYearFixed"),             // today's 30-year fixed rate Redfin shows (Bankrate)
    openHouses: openHouses,
    address: street, city: city, zip: zip,
    price: num(/\$([\d,]+)/),
    beds: num(/([\d.]+)\s*beds?/i),
    baths: num(/([\d.]+)\s*baths?/i),
    sqft: num(/([\d,]+)\s*sq\.?\s*ft/i),
    yearBuilt: year ? Number(year[1]) : null,
    mls: (desc.match(/MLS#\s*([\w-]+)/) || [])[1] || "",
    status: status,
    redfinUrl: url,
    photoUrl: /^https:\/\//.test(meta("og:image")) ? meta("og:image") : "",
  };
}

/**
 * The elementary school assigned to the home, from Redfin's "servingThisHomeSchools" data:
 * [{ servesHome: true, greatSchoolsRating: 8, gradeRanges: "PreK-5", name: "… Elementary School" }, …].
 * That data sits inside the page as text that has been escaped more than once, so we strip
 * the backslashes first and then read each school. Returns { name, rating } or null.
 */
function elementarySchool_(html) {
  const at = html.indexOf("servingThisHomeSchools");
  if (at < 0) return null;
  const text = html.slice(at, at + 8000).replace(/\\/g, "");
  const end = text.indexOf("]");
  const schools = (end > 0 ? text.slice(0, end) : text).split('{"servesHome"').slice(1).map((chunk) => ({
    serves: /^:true/.test(chunk),
    rating: Number((chunk.match(/"greatSchoolsRating":(\d+)/) || [])[1]) || null,
    grades: (chunk.match(/"gradeRanges":"([^"]*)"/) || [])[1] || "",
    name: (chunk.match(/"name":"([^"]*)"/) || [])[1] || "",
  }));
  // Elementary = grades that stop at 5th or lower (PreK-5, K-5…), or "Elementary" in the name.
  const isElementary = (x) => /elementary/i.test(x.name) || /^(PreK|K|\d)-?([1-5])?$/i.test(x.grades.replace(/\s/g, "")) && Number((x.grades.match(/(\d+)$/) || [])[1]) <= 5;
  const pick = schools.find((x) => x.serves && isElementary(x)) || null;
  return pick && pick.name ? { name: pick.name, rating: pick.rating } : null;
}

/** Add (or bring back) a home from its Redfin link, at the bottom of our ranked list. */
function addFromRedfin_(body) {
  const info = redfinLookup_(body.url);
  const sheet = sheet_();
  const homes = readHomes_();
  const id = (info.address + "-" + info.zip).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  const sameStreet = (h) => h.zip === info.zip && h.address.toLowerCase().replace(/\W/g, "").slice(0, 6) === info.address.toLowerCase().replace(/\W/g, "").slice(0, 6);
  const existing = homes.find((h) => h.id === id || h.redfinUrl === info.redfinUrl || sameStreet(h));
  const nextRank = homes.reduce((max, h) => (h.rank && !h.removed ? Math.max(max, h.rank) : max), 0) + 1;
  const text = (v, max) => String(v == null ? "" : v).trim().slice(0, max);
  const why = text(body.why, 500) || "Added by " + (text(body.addedBy, 60) || "the family") + ".";

  if (existing && existing.rank && !existing.removed) throw new Error(existing.address + " is already on our list (#" + existing.rank + ").");

  // Safety grades are by area, so borrow the grade most homes in the same zip have.
  const grades = homes.filter((h) => h.zip === info.zip && h.safety).map((h) => h.safety);
  const safety = grades.sort((a, b) => grades.filter((g) => g === b).length - grades.filter((g) => g === a).length)[0] || "";

  if (existing) {
    // Already in the sheet (maybe removed or unranked): refresh it and put it back on the list.
    const values = sheet.getDataRange().getValues();
    const header = values[0];
    const r = values.findIndex((row, i) => i > 0 && String(row[0]) === existing.id);
    const set = (key, v) => { const c = header.indexOf(key); if (c >= 0 && v !== null && v !== "") sheet.getRange(r + 1, c + 1).setValue(v); };
    set("price", info.price); set("photoUrl", info.photoUrl); set("redfinUrl", info.redfinUrl);
    set("status", info.status); set("rank", nextRank); set("why", why);
    const extra = detailFields_(info);
    Object.keys(extra).forEach((k) => set(k, extra[k]));
    sheet.getRange(r + 1, header.indexOf("removed") + 1).setValue(false);
    return Object.assign({}, existing, info, { rank: nextRank, why: why, removed: false });
  }

  const home = Object.assign(cleanHome_(Object.assign({}, info, {
    addedBy: text(body.addedBy, 60), list: "Added " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d"),
  })), { safety: safety, rank: nextRank, why: why, watch: "", deadline: "" }, detailFields_(info));
  sheet.appendRow(headerRow_(sheet, home));
  return home;
}

/** The extra listing details we keep in the Homes tab (open houses are saved as JSON text). */
function detailFields_(info) {
  return {
    lat: info.lat, lng: info.lng, garage: info.garage, dom: info.dom, hoa: info.hoa,
    elementary: info.elementary || null, elementaryScore: info.elementaryScore,
    taxRate: info.taxRate, insurance: info.insurance,
    openHouses: JSON.stringify(info.openHouses || []),
  };
}

/** Build a row in the sheet's actual column order (safer than assuming COLUMNS order). */
function headerRow_(sheet, home) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return header.map((key) => (home[key] == null ? "" : home[key]));
}

function decode_(s) {
  return s.replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">");
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
    mls: text(h.mls, 40), daysOnMarket: "", mortgage: num(h.mortgage),
    safety: text(h.safety, 3), yearBuilt: num(h.yearBuilt), sqft: num(h.sqft),
    beds: num(h.beds), baths: num(h.baths), hoa: num(h.hoa), notes: text(h.notes, 1000),
    elementary: "", elementaryScore: null, middle: "", middleScore: null, high: "", highScore: null,
    list: text(h.list, 40), redfinUrl: url(h.redfinUrl), photoUrl: url(h.photoUrl), status: text(h.status, 20) || "Active",
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
function csv_(homes, includeRemoved, comments, safety) {
  const zipSafety = (h) => (safety || {})[h.zip] || {};
  const notesFor = (id) => (comments || []).filter((c) => c.homeId === id).map((c) => c.name + ": " + c.text).join(" | ");
  const list = homes
    .filter((h) => includeRemoved || !h.removed)
    .sort((a, b) => (a.rank || 999) - (b.rank || 999) || (a.price || 0) - (b.price || 0));
  const head = ["Rank", "Address", "Town", "Zip", "Price", "Beds", "Baths", "Sq Ft", "Year Built",
    "Safety", "Status", "Why we like it", "Keep in mind", "Deadline", "HOA $/mo",
    "Elementary", "Middle", "High", "Safety score (of 1000)", "CrimeGrade", "Redfin", "On our list", "Why we removed it", "Family comments"];
  const cell = (v) => {
    const s = v == null ? "" : String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const rows = list.map((h) => [
    h.rank || "", h.address, h.city, h.zip, h.price, h.beds, h.baths, h.sqft, h.yearBuilt,
    zipSafety(h).grade || h.safety, h.status, h.why, h.watch, h.deadline, h.hoa,
    h.elementary, h.middle, h.high, zipSafety(h).score || "", crimeGradeUrl_(h.zip), h.redfinUrl, h.removed ? "Removed" : "Yes", h.removed ? h.removedReason : "", notesFor(h.id),
  ].map(cell).join(","));
  return ContentService.createTextOutput([head.join(",")].concat(rows).join("\n"))
    .setMimeType(ContentService.MimeType.CSV);
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- "On our page" tab: a read-only, color-coded view of the homes on the family page ---------- */

const PAGE_TAB = "On our page";

/**
 * Run once from the editor. Adds an "On our page" tab that always shows exactly the homes
 * the family page shows (ranked and not removed), in the same order. Everything on it is
 * built from formulas that read the Homes and Comments tabs, so it updates by itself and
 * holds no data of its own. Safe to run again: it rebuilds the tab. Logs the tab's link.
 *
 * Tricks used: FILTER + SORT (only our homes, in rank order), ARRAYFORMULA (one formula
 * fills a whole column), HYPERLINK (short clickable links), COUNTIF across tabs (comment
 * counts), and conditional formatting (stripes, a gold #1 row, color scales).
 */
function setupPageTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const homes = sheet_();
  commentsSheet_(); // make sure the Comments tab exists for the comment counts
  // Look columns up by header name, so the formulas keep working if columns move.
  const letter = (name) => homes.getRange(1, col_(homes, name)).getA1Notation().replace(/\d+/g, "");
  const r = (name) => `Homes!${letter(name)}2:${letter(name)}`;

  // B..N come from one FILTER formula. L, M and N are helper columns we hide.
  const pulled = [
    ["Home", "address"], ["Town", "city"], ["Price", "price"], ["Beds", "beds"], ["Baths", "baths"],
    ["Sq ft", "sqft"], ["Built", "yearBuilt"], ["Why we like it", "why"], ["Keep in mind", "watch"],
    ["Our notes", "notes"], ["zip", "zip"], ["redfin", "redfinUrl"], ["id", "id"],
  ];
  const keep = `(${r("rank")}<>"")*((${r("removed")}&"")<>"TRUE")`; // has a rank AND isn't removed
  const main = `=IFERROR(SORT(FILTER({${pulled.map(([, k]) => r(k)).join(",")}}, ${keep}), FILTER(${r("rank")}, ${keep}), TRUE), "")`;

  let tab = ss.getSheetByName(PAGE_TAB);
  if (!tab) tab = ss.insertSheet(PAGE_TAB, 0); // first tab, so it's what people see first
  tab.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => p.remove());
  tab.clear();
  tab.clearConditionalFormatRules();
  tab.showColumns(1, tab.getMaxColumns());

  const headers = ["#"].concat(pulled.map(([label]) => label), ["Redfin", "Safety", "Comments"]);
  tab.getRange(1, 1, 1, headers.length).setValues([headers]);
  tab.getRange("A2").setFormula('=ARRAYFORMULA(IF(B2:B="", "", ROW(B2:B)-1))');   // 1, 2, 3… like the page
  tab.getRange("B2").setFormula(main);
  tab.getRange("O2").setFormula('=ARRAYFORMULA(IF(M2:M="", "", HYPERLINK(M2:M, "Open ↗")))');
  tab.getRange("P2").setFormula('=ARRAYFORMULA(IF(L2:L="", "", HYPERLINK("https://crimegrade.org/safest-places-in-"&L2:L&"/", "Safety map ↗")))');
  tab.getRange("Q2").setFormula('=ARRAYFORMULA(IF(N2:N="", "", IF(COUNTIF(Comments!B2:B, N2:N)=0, "—", "💬 "&COUNTIF(Comments!B2:B, N2:N))))');
  tab.hideColumns(12, 3); // L, M, N: zip, Redfin address, id

  // ---- Look ----
  const NAVY = "#12355f", STRIPE = "#f2f6fb", GOLD = "#fdf2cc";
  const all = tab.getRange("A1:Q");
  all.setFontFamily("Nunito").setFontSize(11).setVerticalAlignment("top");
  tab.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground(NAVY).setFontColor("#ffffff")
    .setVerticalAlignment("middle").setWrap(true);
  tab.setRowHeight(1, 34);
  tab.setFrozenRows(1);
  tab.setFrozenColumns(2); // # and Home stay put while scrolling sideways
  tab.getRange("A2:A").setHorizontalAlignment("center").setFontWeight("bold").setFontColor(NAVY);
  tab.getRange("B2:B").setFontWeight("bold").setFontColor(NAVY);
  tab.getRange("D2:D").setNumberFormat("$#,##0").setFontWeight("bold");
  tab.getRange("E2:F").setHorizontalAlignment("center");
  tab.getRange("G2:G").setNumberFormat("#,##0");
  tab.getRange("H2:H").setNumberFormat("0").setHorizontalAlignment("center");
  tab.getRange("I2:K").setWrap(true);
  tab.getRange("J2:J").setFontColor("#9a3412");   // Keep in mind: red, like the page
  tab.getRange("K2:K").setFontColor("#5b3a8c");   // Our notes: purple, like the page
  tab.getRange("O2:Q").setHorizontalAlignment("center");
  [34, 190, 105, 95, 50, 55, 70, 60, 320, 250, 250, 60, 60, 60, 90, 110, 95].forEach((w, i) => tab.setColumnWidth(i + 1, w));
  tab.setTabColor(NAVY);
  tab.getRange("A1:Q").setBorder(null, null, null, null, false, true, "#d3deec", SpreadsheetApp.BorderStyle.SOLID);

  // ---- Color coding (conditional formatting) ----
  const rows = [tab.getRange("A2:Q")];
  const rule = () => SpreadsheetApp.newConditionalFormatRule();
  const rules = [
    // #1 home: gold row (listed first so it wins over the stripes)
    rule().whenFormulaSatisfied("=$A2=1").setBackground(GOLD).setRanges(rows).build(),
    // Stripes on every other home
    rule().whenFormulaSatisfied('=AND($B2<>"", ISEVEN(ROW()))').setBackground(STRIPE).setRanges(rows).build(),
    // Price: green = cheaper, orange = pricier
    rule().setGradientMinpoint("#c9ecd6").setGradientMidpointWithValue("#ffffff", SpreadsheetApp.InterpolationType.PERCENTILE, "50")
      .setGradientMaxpoint("#fbd7b8").setRanges([tab.getRange("D2:D")]).build(),
    // Size and year built: the bigger / newer, the greener
    rule().setGradientMinpoint("#ffffff").setGradientMaxpoint("#9fd8b4").setRanges([tab.getRange("G2:G")]).build(),
    rule().setGradientMinpoint("#ffffff").setGradientMaxpoint("#9fd8b4").setRanges([tab.getRange("H2:H")]).build(),
    // 4+ bedrooms or 3+ baths: green and bold
    rule().whenNumberGreaterThanOrEqualTo(4).setFontColor("#1f7a45").setBold(true).setRanges([tab.getRange("E2:E")]).build(),
    rule().whenNumberGreaterThanOrEqualTo(3).setFontColor("#1f7a45").setBold(true).setRanges([tab.getRange("F2:F")]).build(),
    // Homes with comments stand out
    rule().whenTextStartsWith("💬").setBackground("#e6eefb").setFontColor(NAVY).setBold(true).setRanges([tab.getRange("Q2:Q")]).build(),
  ];
  tab.setConditionalFormatRules(rules);

  // Warn anyone who tries to type here: changes belong on the page (or the Homes tab).
  tab.protect().setDescription("Built automatically from the Homes tab. Make changes on the family page.").setWarningOnly(true);

  const url = ss.getUrl().replace(/\/edit.*$/, "") + "/edit#gid=" + tab.getSheetId();
  Logger.log("On our page tab ready: %s", url);
  return url;
}

/* ---------- Daily Redfin check: prices and status ---------- */

const HISTORY_TAB = "Price history";

/**
 * Runs once a day from a time-driven trigger (Triggers, in the left menu of the editor).
 * For every home on our page it re-reads the Redfin listing and, if the price or the
 * status (Active / Pending / Sold / Off market) changed, updates the Homes tab and adds
 * a line to the "Price history" tab. Safe to run by hand any time.
 */
function refreshPrices() {
  // 1) Read the listings first, without holding the lock (this is the slow part).
  const fresh = {};
  let failed = 0;
  readHomes_().filter((h) => h.rank && !h.removed && REDFIN_URL.test(h.redfinUrl)).forEach((h) => {
    try {
      const info = redfinLookup_(h.redfinUrl);
      // No beds in the summary means we got something other than the listing (a "please
      // verify" page, say). Skip it, so a bad read can never mark a home as off the market.
      if (info.beds === null) failed++; else fresh[h.id] = info;
    } catch (e) { failed++; }
    Utilities.sleep(400); // be polite to Redfin
  });

  // 2) Then take the lock (so we can't collide with someone using the page) and write the changes.
  //    We change the values in memory, then write each changed column back in one go:
  //    one setValues() per column is much faster than one setValue() per cell.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const changes = [];
  try {
    const sheet = sheet_();
    const values = sheet.getDataRange().getValues();
    const header = values[0];
    const col = (key) => header.indexOf(key);
    const touched = {};
    const now = new Date().toISOString();
    for (let r = 1; r < values.length; r++) {
      const info = fresh[String(values[r][0])];
      if (!info) continue;
      const set = (key, v) => { if (values[r][col(key)] !== v) { values[r][col(key)] = v; touched[key] = true; } };
      const name = values[r][col("address")] + ", " + values[r][col("city")];

      const oldStatus = String(values[r][col("status")] || "");
      if (info.status && info.status !== oldStatus) {
        set("status", info.status);
        set("statusChangedAt", now);
        changes.push([now, name, "Status", oldStatus || "(blank)", info.status, "", info.redfinUrl]);
      }

      // Only trust the price while the home is for sale, and ignore anything that looks
      // like a misread (less than half or more than double the price we have).
      const oldPrice = Number(values[r][col("price")]) || 0;
      const forSale = info.status === "Active" || info.status === "Pending";
      const sane = info.price && (!oldPrice || (info.price > oldPrice * 0.5 && info.price < oldPrice * 2));
      if (forSale && sane && info.price !== oldPrice) {
        set("price", info.price);
        if (oldPrice) { set("lastPrice", oldPrice); set("priceChangedAt", now); }
        if (oldPrice && info.price < oldPrice) set("drops", (Number(values[r][col("drops")]) || 0) + 1);
        changes.push([now, name, "Price", oldPrice || "", info.price, oldPrice ? info.price - oldPrice : "", info.redfinUrl]);
      }

      // Quiet details: no history line, just keep them current.
      const extra = detailFields_(info);
      Object.keys(extra).forEach((k) => { if (extra[k] !== null && extra[k] !== undefined) set(k, extra[k]); });
    }
    Object.keys(touched).forEach((key) => {
      const c = col(key);
      sheet.getRange(2, c + 1, values.length - 1, 1).setValues(values.slice(1).map((row) => [row[c]]));
    });
    if (changes.length) {
      const log = historySheet_();
      log.getRange(log.getLastRow() + 1, 1, changes.length, changes[0].length).setValues(changes);
    }
    // Remember today's 30-year rate from Redfin for the monthly-cost estimate.
    const rate = Object.keys(fresh).map((id) => fresh[id].rate30).filter((x) => x > 1 && x < 20)[0];
    const props = PropertiesService.getScriptProperties();
    if (rate) props.setProperty("RATE30", String(rate));
    props.setProperty("PRICES_CHECKED", now);
    clearListCache_(); // so the page shows the new prices right away
  } finally {
    lock.releaseLock();
  }
  Logger.log("Checked %s homes, %s changes, %s could not be read.", Object.keys(fresh).length, changes.length, failed);
  return changes;
}

/** The "Price history" tab: one line per price or status change. Created the first time it's needed. */
function historySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HISTORY_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_TAB);
    sheet.getRange(1, 1, 1, 7).setValues([["When", "Home", "What changed", "Before", "Now", "Difference", "Redfin"]])
      .setFontWeight("bold").setBackground("#12355f").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.getRange("D2:F").setNumberFormat("$#,##0;-$#,##0");
    [170, 220, 110, 100, 100, 100, 320].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

/* =====================================================================
   Settings tab: the family's must-haves and other knobs, in plain words.
   Edit the Value column in the Sheet; the page and the morning job read it.
   ===================================================================== */

const SETTINGS_TAB = "Settings";
const SETTINGS = [
  // key, label shown in the Sheet, default, what it means
  ["budget", "Budget (most we want to pay)", 400000, "Homes above this get a red \"over budget\" tick on the page."],
  ["beds", "Must have: bedrooms", 3, "At least this many bedrooms."],
  ["baths", "Must have: bathrooms", 2.5, "At least this many bathrooms."],
  ["sqft", "Must have: square feet", 2250, "At least this many square feet."],
  ["garage", "Must have: garage spaces", 2, "At least this many garage spaces."],
  ["school", "Must have: elementary school rating", 8, "GreatSchools rating out of 10 for the assigned elementary school (8 = 8, 9 and 10 are OK). The radar hides homes below it."],
  ["towns", "New-listing radar: towns", "Cary, Clayton, Wake Forest, Holly Springs, Angier, Garner, Fuquay-Varina", "Towns the morning job searches on Redfin, separated by commas."],
  ["radarMax", "New-listing radar: highest price", 450000, "The radar shows new homes up to this price."],
  ["downPct", "Down payment (%)", 20, "Used for the monthly cost. Also editable on the page."],
  ["rate", "Mortgage rate (%)", "", "Leave blank to use today's 30-year fixed rate from Redfin."],
  ["family", "Family voters", "Amin, Aurora, Rachel, David", "Names for the hearts vote, separated by commas."],
  ["emailTo", "Morning email also goes to", "Katie.M.Ernst@gmail.com, David.L.Ernst@gmail.com", "The Sheet's owner always gets it. Separate addresses with commas. Leave blank for just the owner."],
  ["places", "Drive times to", "RDU Airport = Raleigh-Durham International Airport, NC; Downtown Raleigh = Fayetteville St, Raleigh, NC; WakeMed = WakeMed Raleigh Campus, Raleigh, NC", "Short name = address, separated by semicolons."],
  ["pageUrl", "Family page link", "https://aminkhazri.github.io/house-hunt/", "Used in the morning email."],
];

function settingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_TAB);
    sheet.getRange(1, 1, 1, 3).setValues([["Setting", "Value", "What it means"]])
      .setFontWeight("bold").setBackground("#12355f").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.getRange("B:B").setNumberFormat("@").setHorizontalAlignment("left"); // keep values as typed
    [260, 380, 520].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.setTabColor("#e3b94f");
  }
  // Add any setting the tab doesn't have yet (keeps values people already changed).
  const labels = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat() : [];
  const missing = SETTINGS.filter(([, label]) => labels.indexOf(label) < 0);
  if (missing.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3)
      .setValues(missing.map(([, label, def, help]) => [label, String(def), help]));
  }
  return sheet;
}

/** All settings as { key: value } (numbers where the default is a number). Cached for a minute. */
function settings_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get("SETTINGS");
  if (hit) return JSON.parse(hit);
  const sheet = settingsSheet_();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const out = {};
  SETTINGS.forEach(([key, label, def]) => {
    const row = rows.find((r) => r[0] === label);
    let v = row ? String(row[1]).trim() : String(def);
    if (typeof def === "number") v = v === "" || isNaN(Number(v.replace(/[$,%]/g, ""))) ? def : Number(v.replace(/[$,%]/g, ""));
    out[key] = v;
  });
  cache.put("SETTINGS", JSON.stringify(out), 60);
  return out;
}

const splitList_ = (s, sep) => String(s || "").split(sep || ",").map((x) => x.trim()).filter(String);

/** "RDU Airport = Raleigh-Durham…; Downtown = …" -> [{ label, address }] */
function places_(s) {
  return splitList_(s, ";").map((p) => {
    const i = p.indexOf("=");
    return i > 0 ? { label: p.slice(0, i).trim(), address: p.slice(i + 1).trim() } : { label: p, address: p };
  }).filter((p) => p.label && p.address).slice(0, 5);
}

/** What the page needs to know (no email addresses). */
function publicSettings_() {
  const s = settings_();
  const rate30 = Number(PropertiesService.getScriptProperties().getProperty("RATE30")) || null;
  return {
    budget: s.budget, beds: s.beds, baths: s.baths, sqft: s.sqft, garage: s.garage, school: s.school,
    radarMax: s.radarMax, towns: splitList_(s.towns), downPct: s.downPct,
    rate: s.rate || rate30 || 6.5, rateFromRedfin: !s.rate && !!rate30,
    family: splitList_(s.family), places: places_(s.places).map((p) => p.label),
  };
}

/** The page's monthly-cost box can change the down payment and the rate (blank rate = Redfin's). */
function saveCostSettings_(body) {
  const sheet = settingsSheet_();
  const labels = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const write = (key, v) => {
    const label = SETTINGS.find((x) => x[0] === key)[1];
    sheet.getRange(labels.indexOf(label) + 2, 2).setValue(String(v));
  };
  const down = Number(body.downPct);
  if (!(down >= 0 && down <= 100)) throw new Error("Down payment should be between 0 and 100%.");
  write("downPct", down);
  const rate = String(body.rate == null ? "" : body.rate).trim();
  if (rate !== "" && !(Number(rate) > 0 && Number(rate) < 20)) throw new Error("The rate should be a percent, like 6.5.");
  write("rate", rate);
  CacheService.getScriptCache().remove("SETTINGS");
  return publicSettings_();
}

/* ---------- Small helpers for the extra tabs ---------- */

function parseJson_(text, fallback) {
  try { return text ? JSON.parse(text) : fallback; } catch (e) { return fallback; }
}

/** A tab with a header row, created the first time it's needed. */
function tab_(name, columns, color) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, columns.length).setValues([columns])
      .setFontWeight("bold").setBackground(color || "#e8f1fd");
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, sheet.getMaxRows(), columns.length).setNumberFormat("@"); // plain text, no surprises
  }
  return sheet;
}

/** Every row of a tab as objects ({} per row, keyed by the column names). */
function readTab_(name, columns) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, columns.length).getValues()
    .filter((r) => r[0] !== "")
    .map((r) => {
      const o = {};
      columns.forEach((c, i) => { o[c] = r[i] instanceof Date ? r[i].toISOString() : String(r[i]); });
      return o;
    });
}

function homeExists_(id) {
  const sheet = sheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat().map(String);
  if (ids.indexOf(String(id)) < 0) throw new Error("Couldn't find that home. Refresh the page and try again.");
  return ids.indexOf(String(id)) + 2;
}

/* ---------- Family votes: 1 to 5 hearts per person per home ---------- */

const VOTES_TAB = "Votes";
const VOTE_COLUMNS = ["homeId", "name", "hearts", "at"];

function saveVote_(body) {
  homeExists_(body.id);
  let name = String(body.name || "").trim().slice(0, 40);
  if (!name) throw new Error("Please pick your name.");
  // "rachel" and "Rachel " count as the same voter as "Rachel".
  const fam = splitList_(settings_().family);
  name = fam.find((f) => f.toLowerCase() === name.toLowerCase()) || name;
  const hearts = Math.round(Number(body.hearts));
  if (!(hearts >= 0 && hearts <= 5)) throw new Error("Pick 1 to 5 hearts.");

  const sheet = tab_(VOTES_TAB, VOTE_COLUMNS, "#fde2e4");
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues() : [];
  const i = rows.findIndex((r) => String(r[0]) === String(body.id) && String(r[1]).toLowerCase() === name.toLowerCase());
  const vote = { homeId: String(body.id), name: name, hearts: String(hearts), at: new Date().toISOString() };
  if (hearts === 0) { if (i >= 0) sheet.deleteRow(i + 2); return vote; } // 0 = take my vote back
  const row = VOTE_COLUMNS.map((k) => vote[k]);
  if (i >= 0) sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return vote;
}

/* ---------- Tour notes: checklist answers, notes, photos and voice notes ---------- */

const TOURS_TAB = "Tours";
const TOUR_COLUMNS = ["id", "homeId", "name", "at", "answers", "notes", "files"];

function saveTour_(input) {
  const t = input || {};
  homeExists_(t.homeId);
  const id = String(t.id || "");
  if (!/^[a-z0-9-]{6,24}$/.test(id)) throw new Error("Bad tour id.");
  const answers = {};
  Object.keys(t.answers || {}).slice(0, 30).forEach((k) => {
    answers[String(k).slice(0, 30)] = String(t.answers[k]).slice(0, 20);
  });
  const files = (Array.isArray(t.files) ? t.files : []).slice(0, 40)
    .filter((f) => f && /^[\w-]{10,}$/.test(f.id) && /^(photo|audio)$/.test(f.type))
    .map((f) => ({ id: f.id, type: f.type }));
  const tour = {
    id: id, homeId: String(t.homeId), name: String(t.name || "").trim().slice(0, 40) || "Someone",
    at: new Date().toISOString(), answers: JSON.stringify(answers),
    notes: String(t.notes || "").trim().slice(0, 3000), files: JSON.stringify(files),
  };
  const sheet = tab_(TOURS_TAB, TOUR_COLUMNS, "#e6f4ea");
  const ids = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String) : [];
  const row = TOUR_COLUMNS.map((k) => tour[k]);
  const i = ids.indexOf(id);
  if (i >= 0) sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return Object.assign({}, tour, { answers: answers, files: files });
}

/** The Drive folder for tour photos and voice notes (made the first time). */
function tourFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("TOUR_FOLDER");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* deleted: make a new one */ } }
  const folder = DriveApp.createFolder("Home Hunt – tour photos & voice notes");
  props.setProperty("TOUR_FOLDER", folder.getId());
  return folder;
}

/**
 * Save one photo or voice note sent from the page (as base64 text) into the Drive folder.
 * The file is set to "anyone with the link can view", so the family page can show it
 * without anyone signing in to Google.
 */
function saveUpload_(body) {
  const mime = String(body.mime || "").toLowerCase().split(";")[0];
  const type = /^image\/(jpeg|png|webp|heic|heif)$/.test(mime) ? "photo"
    : /^audio\/(webm|mp4|mpeg|ogg|aac|x-m4a|m4a|wav)$/.test(mime) ? "audio" : "";
  if (!type) throw new Error("Only photos and voice notes can be uploaded.");
  const data = String(body.data || "");
  if (!data || data.length > 14e6) throw new Error("That file is too big (10 MB at most).");
  const row = homeExists_(body.homeId);
  const address = sheet_().getRange(row, col_(sheet_(), "address")).getValue();
  const ext = { "image/jpeg": "jpg", "image/png": "png", "image/webp": "webp", "audio/webm": "webm", "audio/mp4": "m4a", "audio/mpeg": "mp3", "audio/ogg": "ogg" }[mime] || mime.split("/")[1];
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH.mm.ss");
  const who = String(body.name || "").trim().slice(0, 40) || "Someone";
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mime, address + " – " + who + " – " + stamp + "." + ext);
  const file = tourFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: file.getId(), type: type };
}

/* =====================================================================
   New-listing radar: every morning, search Redfin for new homes that
   match our must-haves, and put them in a "New today" box on the page.
   ===================================================================== */

const RADAR_TAB = "New listings";
const RADAR_COLUMNS = ["url", "address", "city", "zip", "price", "beds", "baths", "sqft", "garage", "dom",
  "ppsf", "hoa", "lat", "lng", "photoUrl", "openHouse", "foundAt", "status", "school", "schoolScore"];
const RADAR_DAYS = 14;  // a match stays in "New today" for two weeks unless someone adds or skips it
const UA = { "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36" };

/** Redfin's id for a town ("Clayton" -> 3340), looked up once and remembered. */
function regionId_(town) {
  const props = PropertiesService.getScriptProperties();
  const key = "REGION_" + town.toUpperCase().replace(/\W+/g, "_");
  const saved = props.getProperty(key);
  if (saved) return saved;
  const res = UrlFetchApp.fetch("https://www.redfin.com/stingray/do/location-autocomplete?v=2&location=" +
    encodeURIComponent(town + ", NC"), { muteHttpExceptions: true, headers: UA });
  // The answer starts with "{}&&" (a guard against old JSON tricks), then the JSON.
  const m = res.getContentText().match(/"id":"2_(\d+)","type":"?2"?,"name":"([^"]+)"/);
  if (!m || m[2].toLowerCase() !== town.toLowerCase()) return "";
  props.setProperty(key, m[1]);
  return m[1];
}

/** One town's matching homes for sale, from Redfin's "Download all" CSV. */
function radarSearch_(town, s) {
  const id = regionId_(town);
  if (!id) { Logger.log("Radar: Redfin doesn't know the town %s", town); return null; }
  const q = {
    al: 1, market: "raleigh", region_id: id, region_type: 6, uipt: 1, status: 9, sf: "1,2,3,5,6,7", v: 8,
    num_homes: 350, page_number: 1, max_price: s.radarMax, num_beds: s.beds, num_baths: s.baths,
    min_listing_approx_size: s.sqft,
  };
  const url = "https://www.redfin.com/stingray/api/gis-csv?" + Object.keys(q).map((k) => k + "=" + encodeURIComponent(q[k])).join("&");
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: UA });
  if (res.getResponseCode() !== 200) { Logger.log("Radar: %s answered %s", town, res.getResponseCode()); return null; }
  const rows = Utilities.parseCsv(res.getContentText()).filter((r) => r.length > 10);
  const head = rows.shift() || [];
  const at = (name) => head.findIndex((h) => h.indexOf(name) === 0);
  const c = { url: at("URL"), address: at("ADDRESS"), city: at("CITY"), zip: at("ZIP"), price: at("PRICE"),
    beds: at("BEDS"), baths: at("BATHS"), sqft: at("SQUARE FEET"), dom: at("DAYS ON MARKET"), ppsf: at("$/SQUARE FEET"),
    hoa: at("HOA"), status: at("STATUS"), oh: at("NEXT OPEN HOUSE START"), lat: at("LATITUDE"), lng: at("LONGITUDE"), type: at("PROPERTY TYPE") };
  const n = (v) => (v === "" || v == null || isNaN(Number(v)) ? null : Number(v));
  // The town from the link (/NC/Fuquay-Varina/…) is more reliable than the CSV's CITY column,
  // which is sometimes scrambled for new-construction listings.
  const townOf = (url, fallback) => decodeURIComponent((String(url).split("/")[4] || "")).replace(/-/g, " ") || fallback;
  return rows.map((r) => ({
    url: r[c.url], address: r[c.address], city: townOf(r[c.url], r[c.city]), zip: r[c.zip], price: n(r[c.price]),
    beds: n(r[c.beds]), baths: n(r[c.baths]), sqft: n(r[c.sqft]), dom: n(r[c.dom]), ppsf: n(r[c.ppsf]),
    hoa: n(r[c.hoa]), status: r[c.status], type: r[c.type], lat: n(r[c.lat]), lng: n(r[c.lng]),
    openHouse: r[c.oh] ? new Date(r[c.oh]).getTime() || "" : "",
  })).filter((h) => REDFIN_URL.test(h.url) && /single family/i.test(h.type || "Single Family"));
}

const homeNum_ = (url) => (String(url).match(/\/home\/(\d+)/) || [])[1] || "";

/**
 * The radar itself. Returns the homes that are new this morning (for the email).
 * The first run only records what's already for sale, so the box doesn't fill with
 * hundreds of old listings: only homes listed in the last 3 days show up as new.
 */
function radarCheck_() {
  const s = settings_();
  // 1) Ask Redfin about every town (slow part, no lock).
  const found = [];
  let townsRead = 0;
  splitList_(s.towns).forEach((town) => {
    try { const list = radarSearch_(town, s); if (list) { townsRead++; found.push.apply(found, list); } }
    catch (e) { Logger.log("Radar: %s failed: %s", town, e.message); }
    Utilities.sleep(500);
  });
  if (!townsRead) return [];

  // 2) The typical $ per square foot and days on market in each town (for "might negotiate").
  const median = (xs) => { const a = xs.filter((x) => x > 0).sort((p, q) => p - q); return a.length ? a[Math.floor(a.length / 2)] : null; };
  const market = {};
  found.forEach((h) => { (market[h.city] = market[h.city] || { p: [], d: [] }); market[h.city].p.push(h.ppsf); market[h.city].d.push(h.dom); });
  Object.keys(market).forEach((city) => {
    market[city] = { ppsf: median(market[city].p), dom: median(market[city].d), count: market[city].p.length };
  });
  PropertiesService.getScriptProperties().setProperty("MARKET", JSON.stringify(market));

  // 3) Compare with what we've seen before and with our own list.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const fresh = [];
  try {
    const sheet = tab_(RADAR_TAB, RADAR_COLUMNS, "#e9f8ee");
    const old = readTab_(RADAR_TAB, RADAR_COLUMNS);
    const baseline = old.length === 0;
    const byNum = {};
    old.forEach((r) => { byNum[homeNum_(r.url)] = r; });
    const ours = {};
    readHomes_().forEach((h) => { ours[homeNum_(h.redfinUrl)] = true; ours[(h.address.toLowerCase().replace(/\W/g, "").slice(0, 8)) + h.zip] = true; });
    const now = new Date().toISOString();
    const seenToday = {};

    found.forEach((h) => {
      const num = homeNum_(h.url);
      if (!num || seenToday[num]) return;
      seenToday[num] = true;
      const prev = byNum[num];
      if (prev) { // already known: keep the numbers current
        Object.assign(prev, { price: h.price, dom: h.dom, ppsf: h.ppsf, openHouse: h.openHouse });
        if (prev.status === "gone") prev.status = "seen";
        // The school minimum in Settings changed: bring back recent matches that pass it now, hide ones that don't.
        const recent = Date.now() - new Date(prev.foundAt).getTime() < RADAR_DAYS * 864e5;
        const score = Number(prev.schoolScore);
        if (/^school below/.test(prev.status) && score >= s.school && recent) prev.status = "new";
        else if (prev.status === "new" && score && score < s.school) prev.status = "school below " + s.school;
        return;
      }
      if (ours[num] || ours[h.address.toLowerCase().replace(/\W/g, "").slice(0, 8) + h.zip]) return;
      const row = Object.assign({}, h, { garage: "", photoUrl: "", foundAt: now,
        status: baseline && !(h.dom !== null && h.dom <= 3) ? "seen" : "new" });
      byNum[num] = row;
      if (row.status === "new") fresh.push(row);
    });
    // New matches that disappeared from Redfin's results (sold, pending, price went up…)
    Object.keys(byNum).forEach((num) => { if (byNum[num].status === "new" && !seenToday[num]) byNum[num].status = "gone"; });

    // 4) Look at each new match's listing for the photo, the garage and the elementary school
    //    (at most 15 a day). Matches already in the box that were found before the school
    //    check existed get looked at too.
    const toCheck = fresh.concat(Object.keys(byNum).map((k) => byNum[k])
      .filter((r) => r.status === "new" && (r.schoolScore === "" || r.schoolScore === undefined)));
    toCheck.filter((r, i, a) => a.indexOf(r) === i).slice(0, 15).forEach((row) => {
      try {
        const info = redfinLookup_(row.url);
        row.photoUrl = info.photoUrl || row.photoUrl || "";
        row.garage = info.garage == null ? "" : info.garage;
        row.school = info.elementary || "";
        row.schoolScore = info.elementaryScore == null ? "unknown" : info.elementaryScore;
        if (info.garage !== null && info.garage < s.garage) row.status = "small garage";
        else if (info.elementaryScore !== null && info.elementaryScore < s.school) row.status = "school below " + s.school;
      } catch (e) { /* keep it as it is */ }
      Utilities.sleep(400);
    });

    const all = Object.keys(byNum).map((k) => byNum[k]);
    sheet.getRange(1, 1, 1, RADAR_COLUMNS.length).setValues([RADAR_COLUMNS]); // header, in case new columns were added
    sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), RADAR_COLUMNS.length).clearContent();
    if (all.length) sheet.getRange(2, 1, all.length, RADAR_COLUMNS.length)
      .setValues(all.map((r) => RADAR_COLUMNS.map((k) => (r[k] == null ? "" : String(r[k])))));
    clearListCache_();
  } finally {
    lock.releaseLock();
  }
  const kept = fresh.filter((r) => r.status === "new");
  Logger.log("Radar: %s towns, %s matching homes, %s new.", townsRead, found.length, kept.length);
  return kept;
}

/** The matches the page shows in its "New today" box. */
function radarForPage_() {
  const cutoff = Date.now() - RADAR_DAYS * 864e5;
  return readTab_(RADAR_TAB, RADAR_COLUMNS)
    .filter((r) => r.status === "new" && new Date(r.foundAt).getTime() > cutoff)
    .sort((a, b) => new Date(b.foundAt) - new Date(a.foundAt) || Number(a.dom) - Number(b.dom))
    .slice(0, 20)
    .map((r) => {
      const o = {};
      RADAR_COLUMNS.forEach((k) => {
        o[k] = ["price", "beds", "baths", "sqft", "garage", "dom", "ppsf", "hoa", "lat", "lng", "openHouse", "schoolScore"].indexOf(k) >= 0
          ? (r[k] === "" || isNaN(Number(r[k])) ? null : Number(r[k])) : r[k];
      });
      return o;
    });
}

/** "Not for us" on the page, or "added" once it's on our list. */
function decideRadar_(url, decision) {
  const status = { added: "added", skip: "skipped", "new": "new" }[decision] || "skipped"; // "new" = undo a skip
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RADAR_TAB);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const num = homeNum_(url);
  const urls = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const i = urls.findIndex((u) => homeNum_(u) === num);
  if (!num || i < 0) return null;
  sheet.getRange(i + 2, RADAR_COLUMNS.indexOf("status") + 1).setValue(status);
  return status;
}

/* ---------- Drive times (Google Maps, built into Apps Script) ---------- */

/**
 * Minutes from each home to the places in Settings ("Drive times to"), saved as JSON in
 * the drives column. Only recalculated when a home is new or the list of places changes.
 */
function updateDrives_() {
  const places = places_(settings_().places);
  const key = places.map((p) => p.label + "=" + p.address).join(";");
  const sheet = sheet_();
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const c = (k) => header.indexOf(k);
  let lookups = 0, changed = false;
  for (let r = 1; r < values.length && lookups < 80; r++) {
    const row = values[r];
    const removed = row[c("removed")] === true || String(row[c("removed")]).toUpperCase() === "TRUE";
    if (!Number(row[c("rank")]) || removed) continue;
    if (parseJson_(row[c("drives")], {}).key === key) continue;
    const origin = row[c("lat")] && row[c("lng")] ? row[c("lat")] + "," + row[c("lng")]
      : row[c("address")] + ", " + row[c("city")] + ", NC " + row[c("zip")];
    const mins = {};
    places.forEach((p) => {
      try {
        lookups++;
        const d = Maps.newDirectionFinder().setOrigin(origin).setDestination(p.address)
          .setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();
        const leg = d.routes && d.routes[0] && d.routes[0].legs[0];
        if (leg) mins[p.label] = Math.round(leg.duration.value / 60);
      } catch (e) { Logger.log("Drive time failed: %s", e.message); }
    });
    row[c("drives")] = JSON.stringify({ key: key, mins: mins });
    changed = true;
  }
  if (changed) {
    const col = c("drives");
    sheet.getRange(2, col + 1, values.length - 1, 1).setValues(values.slice(1).map((row) => [row[col]]));
    clearListCache_();
  }
  Logger.log("Drive times: %s lookups.", lookups);
}

/* =====================================================================
   Morning email: a short summary to the family when something changed.
   ===================================================================== */

function digestRecipients_() {
  const me = Session.getEffectiveUser().getEmail();
  const others = splitList_(settings_().emailTo).filter((x) => /^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(x));
  return [me].concat(others).filter((x, i, a) => x && a.map((y) => y.toLowerCase()).indexOf(x.toLowerCase()) === i);
}

/**
 * Build the email. changes = rows from refreshPrices(), fresh = new radar matches.
 * Open houses are added on Thursdays and Fridays (a weekend preview), or whenever
 * there's other news. Returns null when there's nothing worth an email.
 */
function buildDigest_(changes, fresh) {
  const s = settings_();
  // Apps Script's toLocaleString() skips the thousands commas, so add them by hand.
  const commas = (n) => String(Math.round(Number(n))).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const money = (n) => "$" + commas(n);
  const tz = Session.getScriptTimeZone();
  const day = Number(Utilities.formatDate(new Date(), tz, "u")); // 1 = Monday … 7 = Sunday
  const drops = changes.filter((c) => c[2] === "Price" && c[5] < 0);
  const rises = changes.filter((c) => c[2] === "Price" && c[5] > 0);
  const status = changes.filter((c) => c[2] === "Status");

  const soon = [];
  readHomes_().filter((h) => h.rank && !h.removed).forEach((h) => {
    parseJson_(h.openHouses, []).forEach((o) => {
      if (o.start > Date.now() && o.start < Date.now() + 4 * 864e5) soon.push({ h: h, o: o });
    });
  });
  soon.sort((a, b) => a.o.start - b.o.start);

  const news = drops.length + rises.length + status.length + fresh.length;
  if (!news && !(soon.length && (day === 4 || day === 5))) return null;

  const li = (html) => `<li style="margin:0 0 8px">${html}</li>`;
  const section = (title, color, items) => items.length
    ? `<h3 style="margin:18px 0 6px;color:${color};font-family:Georgia,serif">${title}</h3><ul style="padding-left:18px;margin:0">${items.join("")}</ul>` : "";
  const link = (url, text) => `<a href="${url}" style="color:#1f4e8c">${text}</a>`;
  const when = (ms) => Utilities.formatDate(new Date(ms), tz, "EEE MMM d, h:mm a");

  const body = [
    section("▼ Price drops", "#17803d", drops.map((c) => li(`<b>${c[1]}</b>: ${money(c[3])} → <b>${money(c[4])}</b> (down ${money(-c[5])}) · ${link(c[6], "Redfin")}`))),
    section("▲ Price went up", "#b45309", rises.map((c) => li(`<b>${c[1]}</b>: ${money(c[3])} → ${money(c[4])} · ${link(c[6], "Redfin")}`))),
    section("Status changes", "#12355f", status.map((c) => li(`<b>${c[1]}</b> is now <b>${c[4]}</b> (was ${c[3]}) · ${link(c[6], "Redfin")}`))),
    section("🆕 New homes that match", "#12355f", fresh.slice(0, 15).map((h) => li(
      `<b>${h.address}, ${h.city}</b>: ${money(h.price)} · ${h.beds} bed · ${h.baths} bath · ${commas(h.sqft)} sq ft` +
      `${h.garage !== "" && h.garage != null ? " · " + h.garage + "-car garage" : ""}` +
      `${Number(h.schoolScore) ? " · Elementary " + h.schoolScore + "/10" : ""} · ${link(h.url, "Redfin")}`))),
    section("Open houses coming up", "#12355f", soon.map((x) => li(`<b>${when(x.o.start)}</b> – ${Utilities.formatDate(new Date(x.o.end), tz, "h:mm a")}: ${x.h.address}, ${x.h.city}`))),
  ].join("");

  const parts = [];
  if (drops.length) parts.push(drops.length === 1 ? "price drop on " + drops[0][1].split(",")[0] : drops.length + " price drops");
  if (fresh.length) parts.push(fresh.length + (fresh.length === 1 ? " new match" : " new matches"));
  if (status.length) parts.push(status.length === 1 ? status[0][1].split(",")[0] + " is " + status[0][4].toLowerCase() : status.length + " status changes");
  if (rises.length) parts.push(rises.length + (rises.length === 1 ? " price increase" : " price increases"));
  if (!parts.length) parts.push(soon.length + (soon.length === 1 ? " open house" : " open houses") + " this weekend");

  const html = `<div style="font-family:Nunito,Segoe UI,Arial,sans-serif;font-size:16px;color:#16243a;max-width:620px">
    <h2 style="font-family:Georgia,serif;color:#12355f;margin:0 0 6px">Good morning! Here's what changed in our home hunt.</h2>
    ${body}
    <p style="margin:22px 0 0"><a href="${s.pageUrl}" style="display:inline-block;background:#12355f;color:#fff;text-decoration:none;font-weight:bold;padding:12px 20px;border-radius:10px">Open our home hunt page</a></p>
    <p style="color:#5a6b80;font-size:13px;margin-top:18px">This email is sent by our home hunt page's morning check, only on days something changes.</p></div>`;
  return { subject: "🏡 Home hunt: " + parts.join(", "), html: html };
}

function sendDigest_(changes, fresh) {
  const mail = buildDigest_(changes || [], fresh || []);
  if (!mail) { Logger.log("Nothing new today, so no email."); return false; }
  MailApp.sendEmail({ to: digestRecipients_().join(","), subject: mail.subject, htmlBody: mail.html, name: "Our Home Hunt" });
  Logger.log("Sent: %s", mail.subject);
  return true;
}

/** Run from the editor to see the email: sends a sample only to you (the Sheet's owner). */
function sendTestEmail() {
  const homes = readHomes_().filter((h) => h.rank && !h.removed);
  const h = homes[0];
  const sample = [[new Date().toISOString(), h.address + ", " + h.city, "Price", h.price + 5000, h.price, -5000, h.redfinUrl]];
  const fresh = radarForPage_().slice(0, 2);
  const mail = buildDigest_(sample, fresh);
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: "[TEST] " + mail.subject, htmlBody: mail.html, name: "Our Home Hunt" });
  Logger.log("Test email sent to %s. Real emails go to: %s", Session.getEffectiveUser().getEmail(), digestRecipients_().join(", "));
}

/* =====================================================================
   The morning job and its schedule.
   ===================================================================== */

/** Everything the page needs fresh each morning. Each step is separate, so one failing doesn't stop the rest. */
function morningCheck() {
  const step = (name, fn, fallback) => { try { return fn(); } catch (e) { Logger.log("%s failed: %s", name, e.message); return fallback; } };
  const changes = step("Price check", refreshPrices, []);
  const fresh = step("Radar", radarCheck_, []);
  step("Drive times", updateDrives_);
  step("Email", () => sendDigest_(changes, fresh));
  clearListCache_();
}

/** Run once from the editor: replaces the old daily price check with morningCheck at 6–7 am. */
function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter((t) => ["refreshPrices", "morningCheck"].indexOf(t.getHandlerFunction()) >= 0)
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("morningCheck").timeBased().everyDays(1).atHour(6).create();
  settingsSheet_();
  Logger.log("morningCheck will run every day between 6 and 7 am (%s).", Session.getScriptTimeZone());
}
