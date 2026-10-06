/**
 * @OnlyCurrentDoc
 *
 * Training Data hub for Morning Coach.
 *
 * Paste this whole file into the Sheet's script editor (Extensions → Apps Script),
 * save, then use the "Training sync" menu that appears in the Sheet.
 *
 * What it does
 *   - Receives morning checks from the Morning Coach app (doPost) → "Mornings" tab
 *   - Pulls Garmin data from intervals.icu every 3 hours → "Wellness" + "Activities" tabs
 *   - Hands the app the latest coach note + last 7 days of wellness (doGet)
 *
 * Secrets (intervals.icu API key, app token) live in Script Properties, never in
 * the Sheet and never in the public GitHub repo. @OnlyCurrentDoc above limits
 * this script to THIS spreadsheet: it cannot open your other Drive files.
 */

const TABS = {
  mornings:   { name: "Mornings",    key: "date", headers: ["date", "stiff_min", "pain", "prev_day", "session_done", "phase", "plan_week", "step", "updated_at"] },
  wellness:   { name: "Wellness",    key: "date", headers: ["date", "resting_hr", "hrv_rmssd", "hrv_sdnn", "sleep_h", "sleep_score", "sleep_quality", "avg_sleep_hr", "readiness", "spo2", "respiration", "weight_kg", "fitness_ctl", "fatigue_atl", "form_tsb", "ramp_rate", "steps", "vo2max", "updated_at"] },
  activities: { name: "Activities",  key: "id",   headers: ["id", "date", "type", "name", "distance_mi", "moving_min", "pace_per_mi", "avg_hr", "max_hr", "cadence", "elev_gain_ft", "load", "decoupling_pct", "rpe", "feel", "temp_f", "updated_at"] },
  notes:      { name: "Coach notes", key: "date", headers: ["date", "note", "source"] }
};
const SYNC_DAYS = 14;       // each sync re-pulls the last two weeks, so late Garmin uploads get picked up
const BACKFILL_DAYS = 90;   // first sync
const INTERVALS = "https://intervals.icu/api/v1";

// ---------------------------------------------------------------- menu (runs when the Sheet opens)
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Training sync")
    .addItem("1 · Connect intervals.icu", "connectIntervals")
    .addItem("2 · Show app connection link", "showAppLink")
    .addSeparator()
    .addItem("Sync Garmin now", "syncNow")
    .addToUi();
}

function connectIntervals() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("Connect intervals.icu",
    "Paste your intervals.icu API key (Settings → Developer Settings at intervals.icu):",
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) return;
  const props = PropertiesService.getScriptProperties();
  props.setProperty("INTERVALS_API_KEY", key);
  setup_();
  const n = syncIntervals_(BACKFILL_DAYS);
  ui.alert("Connected", `Pulled ${n.wellness} days of wellness and ${n.activities} activities. ` +
    "This now repeats every 3 hours on its own.\n\nNext: Training sync → 2 · Show app connection link.", ui.ButtonSet.OK);
}

function showAppLink() {
  const ui = SpreadsheetApp.getUi();
  setup_();
  const props = PropertiesService.getScriptProperties();
  let url = props.getProperty("WEBAPP_URL") || ScriptApp.getService().getUrl() || "";
  // getUrl() can return the editor-only "/dev" address; the app needs the "/exec" one.
  if (!/\/exec$/.test(url)) {
    const res = ui.prompt("One more paste",
      "Paste the Web app URL from the deployment screen (it ends in /exec):", ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    url = res.getResponseText().trim();
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
      ui.alert("That doesn't look like a Web app URL. It should start with https://script.google.com and end with /exec.");
      return;
    }
    props.setProperty("WEBAPP_URL", url);
  }
  const token = props.getProperty("APP_TOKEN");
  // One string to paste into Morning Coach. The token is what stops anyone else
  // who finds the URL from reading or writing your data.
  const html = HtmlService.createHtmlOutput(
    `<p style="font:14px sans-serif">Copy this link into Morning Coach → ⚙ → Sync link:</p>` +
    `<textarea style="width:100%;height:110px;font:12px monospace" onclick="this.select()">${url}?token=${token}</textarea>`
  ).setWidth(460).setHeight(220);
  ui.showModalDialog(html, "Morning Coach connection link");
}

function syncNow() {
  const n = syncIntervals_(SYNC_DAYS);
  SpreadsheetApp.getUi().alert(`Synced: ${n.wellness} wellness days, ${n.activities} activities.`);
}

// ---------------------------------------------------------------- setup (safe to run any number of times)
function setup_() {
  const ss = SpreadsheetApp.getActive();
  for (const t of Object.values(TABS)) {
    let sh = ss.getSheetByName(t.name);
    if (!sh) sh = ss.insertSheet(t.name);
    sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
    // Key columns as plain text: otherwise Sheets turns "2026-10-06" into a Date
    // object and lookups by date stop matching.
    sh.getRange("A:A").setNumberFormat("@");
    if (t === TABS.activities) sh.getRange("B:B").setNumberFormat("@");
  }
  const blank = ss.getSheetByName("Sheet1");
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("APP_TOKEN")) props.setProperty("APP_TOKEN", Utilities.getUuid().replace(/-/g, ""));

  // Exactly one recurring trigger, even if setup runs twice.
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "scheduledSync")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("scheduledSync").timeBased().everyHours(3).create();
}

function scheduledSync() { syncIntervals_(SYNC_DAYS); }

// ---------------------------------------------------------------- intervals.icu → Sheet
function syncIntervals_(days) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty("INTERVALS_API_KEY");
  if (!key) throw new Error("No intervals.icu API key yet. Use Training sync → 1 · Connect intervals.icu.");
  const athlete = props.getProperty("INTERVALS_ATHLETE_ID") || "0";   // 0 = the key's owner
  const tz = Session.getScriptTimeZone();
  const newest = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const oldest = Utilities.formatDate(new Date(Date.now() - days * 864e5), tz, "yyyy-MM-dd");
  const opts = {
    headers: { Authorization: "Basic " + Utilities.base64Encode("API_KEY:" + key) },
    muteHttpExceptions: true
  };
  const get = path => {
    const r = UrlFetchApp.fetch(INTERVALS + path, opts);
    if (r.getResponseCode() !== 200) throw new Error(`intervals.icu ${r.getResponseCode()} on ${path.split("?")[0]}: ${r.getContentText().slice(0, 200)}`);
    return JSON.parse(r.getContentText());
  };
  const stamp = new Date().toISOString();
  try {
    const well = get(`/athlete/${athlete}/wellness?oldest=${oldest}&newest=${newest}`);
    const wRows = well.map(w => wellnessRow_(w, stamp));
    const actFields = "id,start_date_local,type,name,distance,moving_time,average_heartrate,max_heartrate,average_cadence,total_elevation_gain,icu_training_load,decoupling,icu_rpe,feel,average_temp,average_weather_temp";
    const acts = get(`/athlete/${athlete}/activities?oldest=${oldest}&newest=${newest}&fields=${actFields}`);
    const aRows = acts.map(a => activityRow_(a, stamp));
    upsert_(TABS.wellness, wRows);
    upsert_(TABS.activities, aRows);
    props.setProperty("LAST_SYNC", stamp);
    props.deleteProperty("LAST_ERROR");
    return { wellness: wRows.length, activities: aRows.length };
  } catch (err) {
    props.setProperty("LAST_ERROR", `${stamp} ${err.message}`);
    throw err;
  }
}

const r1_ = x => (x == null || x === "" ? "" : Math.round(x * 10) / 10);
const r0_ = x => (x == null || x === "" ? "" : Math.round(x));

function wellnessRow_(w, stamp) {
  const form = w.ctl != null && w.atl != null ? w.ctl - w.atl : null;
  return [w.id, r0_(w.restingHR), r1_(w.hrv), r1_(w.hrvSDNN), w.sleepSecs ? r1_(w.sleepSecs / 3600) : "",
          r0_(w.sleepScore), r0_(w.sleepQuality), r0_(w.avgSleepingHR), r0_(w.readiness), r0_(w.spO2),
          r1_(w.respiration), r1_(w.weight), r1_(w.ctl), r1_(w.atl), r1_(form), r1_(w.rampRate),
          r0_(w.steps), r1_(w.vo2max), stamp];
}

function activityRow_(a, stamp) {
  const mi = a.distance ? a.distance / 1609.344 : 0;
  const min = a.moving_time ? a.moving_time / 60 : 0;
  let pace = "";
  if (mi > 0.1 && min > 0) {
    const p = min / mi, mm = Math.floor(p), ss = Math.round((p - mm) * 60);
    pace = ss === 60 ? `${mm + 1}:00` : `${mm}:${String(ss).padStart(2, "0")}`;
  }
  const c = a.average_weather_temp != null ? a.average_weather_temp : a.average_temp;
  return [a.id, String(a.start_date_local || "").slice(0, 10), a.type || "", a.name || "",
          mi ? Math.round(mi * 100) / 100 : "", r1_(min), pace, r0_(a.average_heartrate), r0_(a.max_heartrate),
          r0_(a.average_cadence), a.total_elevation_gain != null ? r0_(a.total_elevation_gain * 3.28084) : "",
          r0_(a.icu_training_load), r1_(a.decoupling), r0_(a.icu_rpe), r0_(a.feel),
          c != null ? r0_(c * 9 / 5 + 32) : "", stamp];
}

// ---------------------------------------------------------------- generic upsert: replace a row by key, else append; keep sorted
function upsert_(tab, rows) {
  if (!rows.length) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(tab.name);
  const width = tab.headers.length;
  const last = sh.getLastRow();
  const keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]) : [];
  const index = new Map(keys.map((k, i) => [k, i + 2]));
  const appends = [];
  for (const row of rows) {
    const k = String(row[0]);
    row[0] = k;
    const at = index.get(k);
    if (at) sh.getRange(at, 1, 1, width).setValues([row]);
    else { appends.push(row); index.set(k, -1); }
  }
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, width).setValues(appends);
  const n = sh.getLastRow() - 1;
  const sortCol = tab === TABS.activities ? 2 : 1;
  if (n > 1) sh.getRange(2, 1, n, width).sort({ column: sortCol, ascending: true });
}

// ---------------------------------------------------------------- web app: Morning Coach talks to these
function doGet(e) {
  if (!authorized_(e.parameter.token)) return json_({ ok: false, error: "bad token" });
  return json_(snapshot_());
}

// Browsers send this as text/plain so no CORS preflight is needed (Apps Script can't answer one).
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad json" }); }
  if (!authorized_(body.token)) return json_({ ok: false, error: "bad token" });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);   // two quick taps on the phone shouldn't write the same day twice
  try {
    const stamp = new Date().toISOString();
    const rows = (body.logs || [])
      .filter(l => /^\d{4}-\d{2}-\d{2}$/.test(l.date))
      .map(l => [l.date, num_(l.stiff), num_(l.pain), l.prev || "", l.done ? "yes" : "",
                 num_(l.phase), num_(l.week), num_(l.step), stamp]);
    if (rows.length) { setup_ifNeeded_(); upsert_(TABS.mornings, rows); }
    return json_(Object.assign(snapshot_(), { saved: rows.length }));
  } finally { lock.releaseLock(); }
}

function setup_ifNeeded_() {
  if (!SpreadsheetApp.getActive().getSheetByName(TABS.mornings.name)) setup_();
}

function snapshot_() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  const rowsOf = (tab, n) => {
    const sh = ss.getSheetByName(tab.name);
    if (!sh || sh.getLastRow() < 2) return [];
    const last = sh.getLastRow(), count = Math.min(n, last - 1);
    return sh.getRange(last - count + 1, 1, count, tab.headers.length).getDisplayValues()
      .map(r => Object.fromEntries(tab.headers.map((h, i) => [h, r[i]])));
  };
  const notes = rowsOf(TABS.notes, 50).filter(r => r.note);
  return {
    ok: true,
    coachNote: notes.length ? notes[notes.length - 1] : null,
    wellness: rowsOf(TABS.wellness, 14),
    lastSync: props.getProperty("LAST_SYNC") || null,
    lastError: props.getProperty("LAST_ERROR") || null
  };
}

function authorized_(t) {
  const want = PropertiesService.getScriptProperties().getProperty("APP_TOKEN");
  return !!want && t === want;
}
const num_ = v => (v == null || v === "" || isNaN(v) ? "" : Number(v));
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
