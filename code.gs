/**
 * Pôle Sponsor Aquil'Art — Google Apps Script
 * 1. Colle ce fichier dans Code.gs et Index.html dans un fichier HTML nommé "Index".
 * 2. Lance une fois la fonction setup() (menu Exécuter) et accepte les autorisations.
 * 3. Déployer > Nouveau déploiement > Application Web
 *    Exécuter en tant que : Moi · Qui a accès : Tous les utilisateurs disposant d'un compte Google.
 */

const COLS = ["tasks","prospects","reports","referrals","shifts","spots","members","notes","offers","activity","config","people","validations"];
const TZ = "Europe/Paris";
const MEMBERS = [["mathis","Mathis"],["chloe","Chloé"],["alban","Alban"],["stanislas","Stanislas"],["liam","Liam"],["hippolyte","Hippolyte"],["gabriel","Gabriel"]];

/* ---------- web app ---------- */
/* Le site est lu en direct depuis GitHub (Mathisb01/pole-sponsor). Le Drive et le fichier Index servent de secours. */
const SITE_FILE = "PoleSponsor_Index.html";
const GITHUB_URL = "https://raw.githubusercontent.com/Mathisb01/pole-sponsor/main/index.html";
function siteHtml_() {
  // 1. version en ligne sur GitHub (mise à jour automatiquement par Claude)
  try {
    const r = UrlFetchApp.fetch(GITHUB_URL + "?t=" + Math.floor(Date.now() / 60000), { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) return HtmlService.createHtmlOutput(r.getContentText("UTF-8"));
  } catch (e) {}
  // 2. secours : fichier PoleSponsor_Index.html du Drive
  try {
    const it = DriveApp.getFilesByName(SITE_FILE); let best = null;
    while (it.hasNext()) { const f = it.next(); if (!f.isTrashed() && (!best || f.getLastUpdated() > best.getLastUpdated())) best = f; }
    if (best) return HtmlService.createHtmlOutput(best.getBlob().getDataAsString("UTF-8"));
  } catch (e) {}
  // 3. secours : fichier Index du projet
  return HtmlService.createHtmlOutputFromFile("Index");
}
function autoriser() { Logger.log("GitHub : " + UrlFetchApp.fetch(GITHUB_URL).getResponseCode() + " · Drive : " + DriveApp.getFilesByName(SITE_FILE).hasNext()); }
/* Icône de l'onglet : le logo Aquil'Art (fichier PoleSponsor_favicon.png de ton Drive, partagé en lecture) */
// Colle ici le lien direct de l'image (il doit finir par .png), entre les guillemets :
const FAVICON_URL = "";
const FAVICON_ID = "1paogAxxZVY0UmOdhbrHVEj5eH6i52h1-";
const FAVICON_B64 = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAkFBMVEVpXiSTayumlF6hjTDh6uGYb03NqFVfZR9lZxFYHBnPnSyn0+Pe0aWlrZ9dZh0sIyFqYFMjKgvKt5PfyljhxzGfcxA4XB43Txeioxs6URZmIAZxcXF6kGf//wAA/wBYPAt5gSGPLiCJdSSIZhz///8AAFU9RyFRXEVVqgCFiTW6sESGhqL/AAD/f/8AAACrilGhK9D9AAAAMHRSTlPx/fz9/vz+oxH+//7/+WD//g////8eCqMX1BIC/gEBV+wIQ7wBA2RrA0MaJAECAP5QBn0aAAACL0lEQVR42mWUiXYbIQxFJTGCGQKzesnS7Onewv//XZ+Ia6enHHvGSJcnIYSpnEZfPo6H85TeX/f4Xu/Lfv/yssfrx8l0Acrv74Ekn4ZQAP0P0Af6dHUeXQekvwD9f/4T0b8D38pofjN3VznjYQPEaASV6zIScTZ3nqdpmBsJCaIeTio99Em8T12eyFUaJoO9KtHY9wVUDwF1MaZ5qlxppgnr1Xuh5rQEHfvMAIwkrVPyfvMVEkiUyoFYauq87xrA0zR5LGexLBaEGINkAOvWIUVSmgYVr6l2SSWMCxVEUJ+Sb8DEPAya1rUzwFEwBaTuY4xbgg+DB4X8ynUVxj7oEAwAgY/qMGR1PIivVQwIB3oYSaTCrZzTqohlQshCRdRCHAKjSk4Vmt2auq3pkAHeDoRQBmbnlDUnVDBFyZJ19gLdegLgdZgPYoDP85CTVHWOnAELFCpH7xB5TWndRYdiRmK4UUrUYST2zC01pLDGzUdWUFIdtolSj0/knLMCQTqtEqNsuy0KjNQAlJJwlpI257zsWkl2sXrzBzRMjySU4yZxh4G1wmzbsgyQQm89GZRRXAdr243gJN1s3LH15F35WpkpGkQz3PZtWdGXYv1wd/9TeUDTRFhtfZNilA8BWlf/KreYoVm4mSseCI/X8vde3JUlBBMdiOZ5hhg0wnG53KzbUpbjDdENWg5nTE90c3Kfrx6mn8dxfH0s5fEVP95NHy9vu2bPt+XtrSzPH/8N/gCddUE45MUPNQAAAABJRU5ErkJggg==";
function doGet() {
  const out = siteHtml_();
  if (FAVICON_URL) { try { out.setFaviconUrl(FAVICON_URL); } catch (e) {} }
  else try { out.setFaviconUrl("data:image/png;base64," + FAVICON_B64); }
  catch (e) { try { out.setFaviconUrl("https://lh3.googleusercontent.com/d/" + FAVICON_ID + "#.png"); } catch (e2) {} }
  return out
    .setTitle("Pôle Sponsor Aquil'Art")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- storage: one tab per collection (id | data JSON | updatedAt) ---------- */
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (!id) throw new Error("Lance d'abord la fonction setup().");
  return SpreadsheetApp.openById(id);
}
function sheet_(col) {
  const ss = ss_();
  let sh = ss.getSheetByName(col);
  if (!sh) { sh = ss.insertSheet(col); sh.appendRow(["id","data","updatedAt"]); sh.setFrozenRows(1); sh.getRange("A:C").setNumberFormat("@"); }
  return sh;
}
function rows_(col) {
  const sh = sheet_(col);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 2).getValues().filter(r => r[0]).map(r => {
    let d = {}; try { d = JSON.parse(r[1] || "{}"); } catch (e) {}
    d.id = String(r[0]); return d;
  });
}
function readAll_() { const out = {}; COLS.forEach(c => out[c] = rows_(c)); return out; }
function bump_() { const p = PropertiesService.getScriptProperties(); p.setProperty("V", String(+(p.getProperty("V") || 0) + 1)); }
function check_(col, id) {
  if (COLS.indexOf(col) < 0) throw new Error("Collection inconnue");
  if (!/^[\w.\-:@+~]{1,80}$/.test(String(id))) throw new Error("Identifiant invalide");
}
function findRow_(sh, id) {
  const n = sh.getLastRow(); if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function clean_(d) { const o = Object.assign({}, d || {}); delete o.id; return o; }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { const r = fn(); bump_(); return r; } finally { l.releaseLock(); } }

function setDoc(col, id, data) {
  check_(col, id);
  return withLock_(() => {
    const sh = sheet_(col), row = findRow_(sh, id), now = new Date().toISOString();
    const vals = [[String(id), JSON.stringify(clean_(data)), now]];
    if (row < 0) sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues(vals); else sh.getRange(row, 1, 1, 3).setValues(vals);
    return true;
  });
}
function updateDoc(col, id, patch) {
  check_(col, id);
  return withLock_(() => {
    const sh = sheet_(col), row = findRow_(sh, id), now = new Date().toISOString();
    let cur = {};
    if (row > 0) { try { cur = JSON.parse(sh.getRange(row, 2).getValue() || "{}"); } catch (e) {} }
    const next = Object.assign(cur, clean_(patch));
    const vals = [[String(id), JSON.stringify(next), now]];
    if (row < 0) sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues(vals); else sh.getRange(row, 1, 1, 3).setValues(vals);
    return true;
  });
}
function deleteDoc(col, id) {
  check_(col, id);
  return withLock_(() => { const sh = sheet_(col), row = findRow_(sh, id); if (row > 0) sh.deleteRow(row); return true; });
}

/* ---------- sync + présence (pastille en ligne) ---------- */
function sync(v, member) {
  const cache = CacheService.getScriptCache(), now = Date.now();
  let pres = {}; try { pres = JSON.parse(cache.get("presence") || "{}"); } catch (e) {}
  if (member) { pres[member] = now; try { trackSession_(member, now); } catch (e) {} }
  Object.keys(pres).forEach(k => { if (now - pres[k] > 45000) delete pres[k]; });
  cache.put("presence", JSON.stringify(pres), 600);
  const cur = +(PropertiesService.getScriptProperties().getProperty("V") || 0);
  const out = { v: cur, online: Object.keys(pres) };
  if (v !== cur) out.data = readAll_();
  return out;
}

/* ---------- historique des connexions (onglet "Connexions") ---------- */
function logSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName("Connexions");
  if (!sh) { sh = ss.insertSheet("Connexions"); sh.appendRow(["membre","début","dernière activité"]); sh.setFrozenRows(1); sh.getRange("A:C").setNumberFormat("@"); }
  return sh;
}
function trackSession_(member, now) {
  const cache = CacheService.getScriptCache(), key = "sess_" + member;
  let s = null; try { s = JSON.parse(cache.get(key) || "null"); } catch (e) {}
  const iso = new Date(now).toISOString();
  if (!s || now - s.last > 5 * 60000) {
    const l = LockService.getScriptLock(); l.waitLock(10000);
    try { const sh = logSheet_(); const row = sh.getLastRow() + 1; sh.getRange(row, 1, 1, 3).setValues([[member, iso, iso]]); s = { row: row, last: now, wrote: now }; }
    finally { l.releaseLock(); }
  } else {
    s.last = now;
    if (now - s.wrote > 2 * 60000) { logSheet_().getRange(s.row, 3).setValue(iso); s.wrote = now; }
  }
  cache.put(key, JSON.stringify(s), 900);
}
function getLogins() {
  const sh = logSheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  const from = Math.max(2, n - 499);
  return sh.getRange(from, 1, n - from + 1, 3).getValues().filter(r => r[0]).map(r => ({ m: String(r[0]), start: String(r[1]), end: String(r[2]) }));
}

/* ---------- installation ---------- */
function setup() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty("SHEET_ID");
  if (!id) { id = SpreadsheetApp.create("Pôle Sponsor Aquil'Art — données").getId(); p.setProperty("SHEET_ID", id); }
  const ss = SpreadsheetApp.openById(id);
  COLS.forEach(c => {
    const sh = sheet_(c);
    const seed = c === "people" ? PEOPLE_SEED : SEED[c];
    if (seed && sh.getLastRow() < 2) {
      const now = new Date().toISOString();
      const vals = Object.keys(seed).map(k => [k, JSON.stringify(seed[k]), now]);
      if (vals.length) sh.getRange(2, 1, vals.length, 3).setValues(vals);
    }
  });
  logSheet_();
  let m = ss.getSheetByName("Membres");
  if (!m) {
    m = ss.insertSheet("Membres", 0);
    m.getRange(1, 1, 1, 3).setValues([["id","prénom","email (pour les rappels)"]]).setFontWeight("bold");
    m.getRange(2, 1, MEMBERS.length, 2).setValues(MEMBERS);
    m.setFrozenRows(1); m.autoResizeColumns(1, 3);
  }
  const def = ss.getSheetByName("Feuille 1") || ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("rappelDimanche").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(18).inTimezone(TZ).create();
  ScriptApp.newTrigger("lancementLundi").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).inTimezone(TZ).create();
  bump_();
  Logger.log("C'est prêt. Données : " + ss.getUrl());
}

/* ---------- rappels par mail ---------- */
function ymd_(d) { return Utilities.formatDate(d || new Date(), TZ, "yyyy-MM-dd"); }
function addDays_(s, n) { const d = new Date(s + "T12:00:00"); d.setDate(d.getDate() + n); return ymd_(d); }
function monday_() { const s = ymd_(); const d = new Date(s + "T12:00:00"); d.setDate(d.getDate() - (d.getDay() + 6) % 7); return ymd_(d); }
function due_(p) {
  const L = p.lastContact, t = ymd_();
  const m = { a_contacter: 0, contacte: 5, relance1: 7, relance2: 3, canal: 5, discussion: 4 };
  if (!(p.stage in m)) return null;
  if (p.stage === "a_contacter" || !L) return t;
  return addDays_(L, m[p.stage]);
}
function emails_() {
  const sh = ss_().getSheetByName("Membres"); const out = {};
  if (!sh) return out;
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getValues().forEach(r => { if (r[0] && r[2]) out[r[0]] = { name: r[1], email: String(r[2]).trim() }; });
  return out;
}
function url_() { try { return ScriptApp.getService().getUrl() || ""; } catch (e) { return ""; } }

function rappelDimanche() {
  const all = readAll_(), mon = monday_(), em = emails_(), link = url_();
  const done = [], late = [];
  MEMBERS.forEach(([id, name]) => {
    const rep = all.reports.some(r => r.who === id && r.week === mon);
    const contacts = all.prospects.filter(p => p.owner === id && p.lastContact && p.lastContact >= mon).length;
    (rep ? done : late).push(name + " (" + contacts + "/7)");
    if (!rep && em[id]) MailApp.sendEmail(em[id].email, "Ton report de la semaine 📝",
      "Salut " + name + " !\n\nPetit rappel : remplis ton report de la semaine sur le site avant ce soir (2 minutes, en bas de « Ma semaine »).\n\nCette semaine : " + contacts + " contact(s) sur 7.\n\n" + link + "\n\nMerci !\nLe pôle sponsor");
  });
  if (em.mathis) MailApp.sendEmail(em.mathis.email, "Reports du dimanche — pôle sponsor",
    "Reports faits : " + (done.join(", ") || "personne") + "\nPas encore : " + (late.join(", ") || "personne") + "\n\n" + link);
}

function lancementLundi() {
  const all = readAll_(), t = ymd_(), mon = monday_(), em = emails_(), link = url_();
  const defi = all.config.filter(c => c.id === "defi")[0];
  const defiTxt = defi && defi.title && defi.week === mon ? "\nDéfi de la semaine : " + defi.title + "\n" : "";
  const recap = [];
  MEMBERS.forEach(([id, name]) => {
    const rel = all.prospects.filter(p => p.owner === id).filter(p => { const d = due_(p); return d && d <= t; });
    const tasks = all.tasks.filter(x => x.owner === id && !x.done);
    const lateT = tasks.filter(x => x.due && x.due < t);
    const wait = all.people.filter(x => x.assignee === id && Object.keys(x.banks || {}).some(b => x.banks[b] === "En cours" || x.banks[b] === "Demandé"));
    recap.push(name + " : " + rel.length + " relances, " + lateT.length + " tâches en retard, " + wait.length + " personnes à relancer pour leurs comptes");
    if (!em[id]) return;
    const lines = ["Salut " + name + ", c'est lundi, let's go 💪", ""];
    lines.push("Relances à faire aujourd'hui : " + rel.length);
    rel.slice(0, 10).forEach(p => lines.push("  • " + p.name));
    if (tasks.length) { lines.push("", "Tes tâches : " + tasks.length + (lateT.length ? " (dont " + lateT.length + " en retard)" : "")); tasks.slice(0, 8).forEach(x => lines.push("  • " + x.title)); }
    if (wait.length) { lines.push("", "Comptes à relancer : " + wait.map(x => x.name + " (" + Object.keys(x.banks).filter(b => x.banks[b] === "En cours" || x.banks[b] === "Demandé").join(", ") + ")").join(", ")); }
    lines.push(defiTxt, "Objectif : 7 contacts cette semaine.", "", link);
    MailApp.sendEmail(em[id].email, "Ta semaine au pôle sponsor 🏛️", lines.join("\n"));
  });
  if (em.mathis) MailApp.sendEmail(em.mathis.email, "Lancement de semaine — récap équipe", recap.join("\n") + (defiTxt || "\nPas de défi cette semaine : lance-en un depuis Pilotage.") + "\n\n" + link);
}


/* ---------- mise à jour v2 : comptes de l'asso ----------
 * À lancer UNE fois après avoir collé la nouvelle version (menu Exécuter > miseAJour).
 * - importe le tableau des comptes (banques + paris sportifs) dans l'onglet "people"
 * - y reporte les filleuls déjà suivis sur le site (statut + qui s'en occupe)
 * - fusionne les offres de paris sportifs en une seule
 * - ajoute un type aux tâches Hello Bank / paris existantes
 */
function miseAJour() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("MIGRATED_V2")) miseAJourV2_();
  if (!p.getProperty("MIGRATED_V3")) { fusionnerDoublons_(); p.setProperty("MIGRATED_V3", "1"); }
  if (!p.getProperty("MIGRATED_V4")) { lierMembres_(); p.setProperty("MIGRATED_V4", "1"); }
  if (!p.getProperty("MIGRATED_V5")) { nettoyageV5_(); p.setProperty("MIGRATED_V5", "1"); }
  Logger.log("Tout est à jour.");
}

/* v3 : fusion des doublons de noms (Hippo, Chloé) */
const MERGES = [
  { name: "Hippolyte Bruneau", from: ["hippo", "hippo bruneau", "hippolyte bruneau"] },
  { name: "Hippo Blanloeil", from: ["hippo 2", "hippo blanloeil", "hippo blanloeil"] },
  { name: "Chloé Lazam", from: ["chloe", "chloe lazam"] },
  { name: "Chloé Lnd", from: ["chloe 2", "chloe lnd"] }
];
function fusionnerDoublons_() {
  const norm = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
  const RANK = { "": 0, "Pas fait": 1, "Refus": 1, "Demandé": 2, "En cours": 3, "Fait": 4 };
  const all = rows_("people"); const tasks = rows_("tasks");
  MERGES.forEach(m => {
    const group = all.filter(x => m.from.indexOf(norm(x.name)) >= 0);
    if (!group.length) return;
    const keep = group[0], rest = group.slice(1);
    const merged = { name: m.name, banks: {}, paris: [], comment: "", note: "", assignee: null, member: null };
    group.forEach(g => {
      Object.keys(g.banks || {}).forEach(b => { if ((RANK[g.banks[b]] || 0) >= (RANK[merged.banks[b]] || 0)) merged.banks[b] = g.banks[b]; });
      (g.paris || []).forEach(x => { if (merged.paris.indexOf(x) < 0) merged.paris.push(x); });
      if (g.comment) merged.comment = merged.comment ? merged.comment + " · " + g.comment : g.comment;
      if (g.note && !merged.note) merged.note = g.note;
      if (g.assignee && !merged.assignee) merged.assignee = g.assignee;
      if (g.member && !merged.member) merged.member = g.member;
    });
    setDoc("people", keep.id, merged);
    rest.forEach(r => {
      tasks.filter(t => t.personId === r.id).forEach(t => updateDoc("tasks", t.id, { personId: keep.id }));
      deleteDoc("people", r.id);
    });
  });
}

/* v4 : relie les membres du pôle à leur fiche dans la liste de l'asso */
const LINKS = { "chloe lnd": "chloe", "hippo blanloeil": "hippolyte", "mathis": "mathis", "alban": "alban", "gabriel": "gabriel", "liam": "liam", "stan": "stanislas" };
function lierMembres_() {
  const norm = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
  const all = rows_("people"); const ids = Object.keys(LINKS).map(k => LINKS[k]);
  all.forEach(x => {
    const want = LINKS[norm(x.name)] || null;
    if (want && x.member !== want) updateDoc("people", x.id, { member: want });
    else if (!want && x.member && ids.indexOf(x.member) >= 0) updateDoc("people", x.id, { member: null });
  });
}

/* v5 : retire les anciens de l'asso, et passe les comptes importés du sheet en "Déjà fait" */
const PARTIS = ["tim","noemie","louise","leo","emma","manon","blanche 2a","chloe lazam","apo","baptiste"];
function nettoyageV5_() {
  const norm = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
  const tasks = rows_("tasks");
  const logged = {}; rows_("validations").forEach(v => logged[norm(v.person) + "|" + v.account] = true);
  rows_("people").forEach(x => {
    if (PARTIS.indexOf(norm(x.name)) >= 0) {
      tasks.filter(t => t.personId === x.id).forEach(t => deleteDoc("tasks", t.id));
      deleteDoc("people", x.id); return;
    }
    const banks = Object.assign({}, x.banks || {});
    Object.keys(banks).forEach(b => { if (banks[b] === "Fait" && !logged[norm(x.name) + "|" + b]) banks[b] = "Déjà fait"; });
    const paris = [], deja = (x.parisDeja || []).slice();
    (x.paris || []).forEach(s => { if (logged[norm(x.name) + "|" + s]) paris.push(s); else if (deja.indexOf(s) < 0) deja.push(s); });
    updateDoc("people", x.id, { banks: banks, paris: paris, parisDeja: deja });
  });
}

function miseAJourV2_() {
  const p = PropertiesService.getScriptProperties();
  sheet_("people"); sheet_("validations");
  const norm = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
  const existing = rows_("people");
  const byName = {}; existing.forEach(x => byName[norm(x.name)] = x);
  if (!existing.length) Object.keys(PEOPLE_SEED).forEach(id => { const d = Object.assign({}, PEOPLE_SEED[id], { id: id }); byName[norm(d.name)] = d; });
  const RANK = { "": 0, "Pas fait": 1, "Refus": 1, "Demandé": 2, "En cours": 3, "Fait": 4 };
  const MAP = { "validé": "Fait", "en attente": "En cours", "commencé": "En cours", "demandé": "Demandé", "à demander": "", "refus": "Refus" };
  const ALIAS = { "hippo b. (blaneuil)": "hippo blanloeil", "hippo (bruneau)": "hippo bruneau", "chloe l.": "chloe lnd", "blanche": "blanche 1a" };
  const BANKS = ["Hello Bank","Fortuneo","Boursorama","Revolut","N26"];
  rows_("referrals").forEach(r => {
    if (!r.person || r.person === "—") return;
    let k = norm(r.person); k = ALIAS[k] || k;
    let person = byName[k];
    if (!person) { const id = k.replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || ("p" + Date.now()); person = { id: id, name: r.person, banks: {}, paris: [], comment: "", note: "", assignee: null, member: null }; byName[k] = person; }
    const s = MAP[r.status] !== undefined ? MAP[r.status] : "";
    if (BANKS.indexOf(r.offer) >= 0) {
      const cur = (person.banks || {})[r.offer] || "";
      if ((RANK[s] || 0) > (RANK[cur] || 0)) { person.banks = person.banks || {}; person.banks[r.offer] = s; }
    } else if (s === "Fait") {
      person.paris = person.paris || []; if (person.paris.indexOf(r.offer) < 0) person.paris.push(r.offer);
    }
    if (r.by && r.by !== "bureau" && !person.assignee) person.assignee = r.by;
    if (s === "Fait" && r.validatedAt) appendRow_("validations", { person: person.name, account: r.offer, by: r.by, at: r.validatedAt });
  });
  const sh = sheet_("people"); const now = new Date().toISOString();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  const vals = Object.keys(byName).map(k => { const d = Object.assign({}, byName[k]); const id = d.id; delete d.id; return [id, JSON.stringify(d), now]; });
  if (vals.length) sh.getRange(2, 1, vals.length, 3).setValues(vals);
  // offres : une seule fiche paris sportifs
  ["winamax","unibet","betclic","parionssport","pmu"].forEach(id => { try { deleteDoc("offers", id); } catch (e) {} });
  setDoc("offers", "paris", { order: 6, kind: "paris", name: "Comptes de paris sportifs",
    code: "Winamax : AIM4B8\nUnibet : LHCBV\nBetclic : https://go.onelink.me/2887093520/6c3132b8?af_sub5=WALN3X6U\nParions Sport : https://www.enligne.parionssport.fdj.fr/inscription/?campaign=130126&parrain=49D7B3B06C248D88\nPMU : pas de lien de parrainage (appli orange)",
    gain: "", conditions: "Créer son compte avec le code ou le lien du site choisi.",
    etapes: "1. Choisir un site où la personne n'a pas encore de compte\n2. S'inscrire avec le code ou le lien\n3. Valider son identité\n4. Cocher le site sur sa fiche dans Parrainages",
    pitch: "Création en 5 minutes, ça aide directement l'asso.", pieges: "Entrer le code dès l'inscription : impossible de l'ajouter après." });
  // types des tâches existantes
  [["t12","banques"],["t13","banques"],["t19","paris"]].forEach(([id, type]) => { try { updateDoc("tasks", id, { type: type }); } catch (e) {} });
  bump_();
  p.setProperty("MIGRATED_V2", "1");
}
function appendRow_(col, data) {
  const sh = sheet_(col); const id = "v" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([[id, JSON.stringify(data), new Date().toISOString()]]);
}
const PEOPLE_SEED = {"lucas": {"name": "Lucas", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Revolut": "Déjà fait", "Fortuneo": "Déjà fait", "N26": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Betclic", "Winamax", "Unibet", "PMU", "Parions Sport"]}, "maxence": {"name": "Maxence", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Revolut": "Déjà fait", "Fortuneo": "Déjà fait", "N26": "Pas fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Betclic", "Winamax", "Unibet", "PMU", "Parions Sport", "Betsson"]}, "gus": {"name": "Gus", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Pas fait", "Revolut": "Déjà fait", "Fortuneo": "Déjà fait", "N26": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Betclic", "Winamax", "Unibet", "PMU", "Parions Sport", "Betsson", "Bwin", "Vbet", "NetBet", "ZEbet"]}, "mathis": {"name": "Mathis", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Revolut": "Déjà fait", "Fortuneo": "Déjà fait", "N26": "Pas fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": "mathis", "parisDeja": ["Betclic", "Winamax", "Unibet", "PMU", "Parions Sport", "Betsson"]}, "anais": {"name": "Anais", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Revolut": "Déjà fait", "Fortuneo": "Déjà fait", "N26": "Pas fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Winamax", "Unibet", "Parions Sport", "Olybet"]}, "blanche-1a": {"name": "Blanche 1A", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "En cours", "Revolut": "Pas fait", "Fortuneo": "Déjà fait", "N26": "Pas fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Unibet", "Parions Sport"]}, "matheo": {"name": "Mathéo", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Pas fait", "Revolut": "Déjà fait", "Fortuneo": "Demandé", "N26": "Pas fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Winamax", "Unibet", "Parions Sport"]}, "milan": {"name": "Milan", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Revolut": "Pas fait", "Fortuneo": "Déjà fait", "N26": "Pas fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Unibet", "Parions Sport"]}, "paul": {"name": "Paul", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Betclic", "Winamax", "Unibet", "Parions Sport"]}, "romy": {"name": "Romy", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "En cours", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Betclic", "Winamax", "Unibet", "PMU", "Parions Sport", "Betsson"]}, "hippo-bruneau": {"name": "Hippolyte Bruneau", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Demandé", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "hippo-blanloeil": {"name": "Hippo Blanloeil", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Demandé", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": "hippolyte", "parisDeja": ["Winamax", "PMU"]}, "clem": {"name": "Clem", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "albane": {"name": "Albane", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Fortuneo": "Demandé"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "liam": {"name": "Liam", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": "liam", "parisDeja": []}, "prune": {"name": "Prune", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Fortuneo": "En cours"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": ["Winamax", "Betsson"]}, "edmond": {"name": "Edmond", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Demandé", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "maxime": {"name": "Maxime", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Déjà fait", "Fortuneo": "Demandé"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "chloe-lnd": {"name": "Chloé Lnd", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "En cours", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": "chloe", "parisDeja": ["Betclic", "Winamax", "Unibet", "PMU", "Bet365", "Betsson", "Bwin", "Vbet"]}, "alban": {"name": "Alban", "banks": {"Hello Bank": "En cours", "Fortuneo": "Pas fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": "alban", "parisDeja": ["Betclic", "Winamax"]}, "gabriel": {"name": "Gabriel", "banks": {"Hello Bank": "En cours"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": "gabriel", "parisDeja": ["Betclic", "Winamax"]}, "jade": {"name": "Jade", "banks": {}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "juliette": {"name": "Juliette", "banks": {}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "stan": {"name": "Stan", "banks": {"Hello Bank": "Déjà fait", "Fortuneo": "Déjà fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": "stanislas", "parisDeja": ["Betclic", "Winamax", "Unibet", "PMU", "Parions Sport", "Bet365", "Betsson", "Bwin", "Vbet"]}, "lily-rose": {"name": "Lily-Rose", "banks": {}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "nathan": {"name": "Nathan", "banks": {}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "nico": {"name": "Nico", "banks": {"Boursorama": "Déjà fait", "Hello Bank": "Pas fait"}, "paris": [], "comment": "", "note": "", "assignee": null, "member": null, "parisDeja": []}, "olympe": {"name": "Olympe", "banks": {}, "paris": [], "comment": "", "note": "mineure jusqu'au 24/12", "assignee": null, "member": null, "parisDeja": []}};

/* ---------- données de départ ---------- */
const SEED = {
"prospects": {
"socgen": {
"ask": "Financement",
"cat": "grands",
"channel": "Mail",
"contact": "stephanie.baes@sgcib.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Société Générale",
"notes": "",
"owner": "mathis",
"stage": "a_contacter"
},
"archimede": {
"ask": "Lieu pour un concert",
"cat": "bars",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Péniche Archimède",
"notes": "",
"owner": "chloe",
"stage": "a_contacter"
},
"aristote": {
"ask": "Lieu pour un concert",
"cat": "bars",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Péniche Aristote",
"notes": "",
"owner": "chloe",
"stage": "a_contacter"
},
"baron": {
"ask": "Lieu ou soirée partenaire",
"cat": "bars",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Le Baron",
"notes": "Y aller quand il y a du monde.",
"owner": "chloe",
"stage": "a_contacter"
},
"creamy": {
"ask": "Produits ou lots",
"cat": "food",
"channel": "Mail",
"contact": "creamydailylille@gmail.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Creamy Daily",
"notes": "",
"owner": "liam",
"stage": "a_contacter"
},
"bulle": {
"ask": "Places de concert à faire gagner",
"cat": "salles",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "La Bulle Café",
"notes": "Mail envoyé par Mathis.",
"owner": "stanislas",
"stage": "contacte"
},
"waterdrop": {
"ask": "Produits ou lots",
"cat": "food",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Waterdrop",
"notes": "",
"owner": "liam",
"stage": "a_contacter"
},
"bose": {
"ask": "Enceinte ou casque à faire gagner",
"cat": "marques",
"channel": "Mail",
"contact": "Joanne_Berthiaume@bose.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Bose",
"notes": "",
"owner": "hippolyte",
"stage": "a_contacter"
},
"tmc": {
"ask": "Financement",
"cat": "grands",
"channel": "Mail",
"contact": "mc.poisseau@tmc-innovation.fr",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "TMC Innovation",
"notes": "Contacté par Prune au printemps.",
"owner": "mathis",
"stage": "contacte"
},
"catchee": {
"ask": "",
"cat": "grands",
"channel": "Insta",
"contact": "Insta",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Catchee",
"notes": "",
"owner": "mathis",
"stage": "contacte"
},
"benscookies": {
"ask": "Une boîte de cookies pour les gagnants",
"cat": "food",
"channel": "Mail",
"contact": "Esquermoise@benscookies.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Ben's Cookies",
"notes": "Démarché en physique aussi, à recontacter.",
"owner": "liam",
"stage": "contacte"
},
"starwax": {
"ask": "Financement ou lots",
"cat": "grands",
"channel": "Mail",
"contact": "etienne.sacilotto@free.fr",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Starwax",
"notes": "En discussion depuis l'ancien mandat (Prune).",
"owner": "mathis",
"stage": "discussion"
},
"circus": {
"ask": "Lieu pour un concert",
"cat": "bars",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Le Circus",
"notes": "",
"owner": "chloe",
"stage": "a_contacter"
},
"redbull": {
"ask": "Canettes pour les événements",
"cat": "grands",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": "2026-10-01",
"name": "Red Bull",
"notes": "",
"owner": "mathis",
"stage": "signe"
},
"cales": {
"ask": "Lieu pour un concert",
"cat": "bars",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Les Cales de Paulette",
"notes": "",
"owner": "chloe",
"stage": "a_contacter"
},
"odelice": {
"ask": "Deal : 50 étudiants, 25 menus un mercredi",
"cat": "food",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Ô Délice",
"notes": "Argument : le kebab le plus proche de l'IÉSEG.",
"owner": "liam",
"stage": "a_contacter"
},
"tropicana": {
"ask": "Boissons pour les événements",
"cat": "food",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Tropicana (et marques de jus)",
"notes": "",
"owner": "liam",
"stage": "a_contacter"
},
"jbl": {
"ask": "Enceinte à faire gagner",
"cat": "marques",
"channel": "Mail",
"contact": "enquiry@imssg.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "JBL",
"notes": "",
"owner": "hippolyte",
"stage": "a_contacter"
},
"seloger": {
"ask": "Financement",
"cat": "grands",
"channel": "Mail",
"contact": "geoffrey.bereyziat@seloger.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "SeLoger",
"notes": "",
"owner": "mathis",
"stage": "a_contacter"
},
"cdiscount": {
"ask": "Lots",
"cat": "grands",
"channel": "Mail",
"contact": "monique.ledieu@cdiscount.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Cdiscount",
"notes": "Contacté par Prune au printemps.",
"owner": "mathis",
"stage": "contacte"
},
"merveilleux": {
"ask": "Invendus de fin de journée",
"cat": "food",
"channel": "Physique",
"contact": "Rue de la Monnaie",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Aux Merveilleux de Fred",
"notes": "",
"owner": "liam",
"stage": "a_contacter"
},
"opera": {
"ask": "Places de spectacle à faire gagner",
"cat": "salles",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Opéra de Lille",
"notes": "",
"owner": "stanislas",
"stage": "a_contacter"
},
"decathlon": {
"ask": "Lots",
"cat": "marques",
"channel": "Mail",
"contact": "theo.ducbert@decathlon.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Decathlon",
"notes": "Contacté par Prune au printemps.",
"owner": "hippolyte",
"stage": "contacte"
},
"chaleur": {
"ask": "",
"cat": "grands",
"channel": "LinkedIn",
"contact": "LinkedIn",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Chaleur Tournante",
"notes": "",
"owner": "mathis",
"stage": "contacte"
},
"gimber": {
"ask": "Produits ou lots",
"cat": "paris",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Gimber",
"notes": "",
"owner": "gabriel",
"stage": "a_contacter"
},
"traknard": {
"ask": "Lieu pour un concert",
"cat": "bars",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Traknard",
"notes": "",
"owner": "chloe",
"stage": "a_contacter"
},
"citya": {
"ask": "Financement",
"cat": "grands",
"channel": "Mail",
"contact": "ydavid@citya.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Citya",
"notes": "",
"owner": "mathis",
"stage": "a_contacter"
},
"boulanger": {
"ask": "Lots high-tech",
"cat": "marques",
"channel": "Mail",
"contact": "katia.gradel@boulanger.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Boulanger",
"notes": "",
"owner": "hippolyte",
"stage": "a_contacter"
},
"buzzlife": {
"ask": "Partenariat événementiel",
"cat": "paris",
"channel": "Mail",
"contact": "lebuzzlife20@gmail.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Le Buzz Life",
"notes": "Démarché en physique, à recontacter.",
"owner": "gabriel",
"stage": "contacte"
},
"cinemas": {
"ask": "Places de cinéma",
"cat": "salles",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Cinémas lillois",
"notes": "À préciser : UGC, Majestic, Le Métropole…",
"owner": "stanislas",
"stage": "a_contacter"
},
"lou": {
"ask": "Pizzas pour un événement",
"cat": "food",
"channel": "Insta",
"contact": "Insta",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Lou Pizzeria",
"notes": "Très actif sur Insta.",
"owner": "liam",
"stage": "a_contacter"
},
"teambreak": {
"ask": "Sessions offertes à faire gagner",
"cat": "paris",
"channel": "Mail",
"contact": "tblille@team-break.fr",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Team Break",
"notes": "",
"owner": "gabriel",
"stage": "a_contacter"
},
"mauroy": {
"ask": "Places de concert ou match",
"cat": "salles",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Stade Pierre Mauroy",
"notes": "L'ancien brouillon disait Fnac Darty : utiliser le générateur.",
"owner": "stanislas",
"stage": "a_contacter"
},
"cafepeinture": {
"ask": "Atelier offert ou réduction",
"cat": "marques",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Café peinture",
"notes": "",
"owner": "hippolyte",
"stage": "a_contacter"
},
"passculture": {
"ask": "Offres culturelles pour nos membres",
"cat": "salles",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Pass Culture",
"notes": "",
"owner": "stanislas",
"stage": "a_contacter"
},
"motte": {
"ask": "Boissons pour les événements",
"cat": "grands",
"channel": "Mail",
"contact": "hellolafamille@bieresmottecordonnier.fr",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Motte Cordonnier",
"notes": "Contacté par Prune au printemps, sans suite.",
"owner": "mathis",
"stage": "contacte"
},
"splendid": {
"ask": "Places de concert à faire gagner",
"cat": "salles",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Le Splendid",
"notes": "L'ancien brouillon disait Fnac Darty : utiliser le générateur.",
"owner": "stanislas",
"stage": "a_contacter"
},
"flunch": {
"ask": "Bons repas à faire gagner",
"cat": "food",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Flunch",
"notes": "",
"owner": "liam",
"stage": "a_contacter"
},
"furet": {
"ask": "Lots (livres, BD)",
"cat": "marques",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Le Furet du Nord",
"notes": "",
"owner": "hippolyte",
"stage": "a_contacter"
},
"mandebar": {
"ask": "Lieu pour un concert",
"cat": "bars",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Mandebar",
"notes": "",
"owner": "chloe",
"stage": "a_contacter"
},
"zenith": {
"ask": "Places de concert à faire gagner",
"cat": "salles",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Zénith de Lille",
"notes": "L'ancien brouillon disait Fnac Darty : utiliser le générateur.",
"owner": "stanislas",
"stage": "a_contacter"
},
"blacklab": {
"ask": "Lots",
"cat": "marques",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Black Lab",
"notes": "",
"owner": "hippolyte",
"stage": "a_contacter"
},
"buddies": {
"ask": "Produits ou lots",
"cat": "food",
"channel": "Physique",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Buddies",
"notes": "",
"owner": "liam",
"stage": "a_contacter"
},
"alcide": {
"ask": "",
"cat": "paris",
"channel": "Téléphone",
"contact": "Gauthier",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Alcide (Gauthier)",
"notes": "Contact perso de Mathis, voir avec lui.",
"owner": "gabriel",
"stage": "a_contacter"
},
"aeronef": {
"ask": "Places de concert à faire gagner",
"cat": "salles",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "L'Aéronef",
"notes": "Mail envoyé par Mathis.",
"owner": "stanislas",
"stage": "contacte"
},
"fnac": {
"ask": "Places de concert ou lots",
"cat": "marques",
"channel": "Mail",
"contact": "suzie.kempa@fnacdarty.com",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Fnac Darty",
"notes": "Mail envoyé par Mathis.",
"owner": "hippolyte",
"stage": "contacte"
},
"soulmade": {
"ask": "Produits ou lots",
"cat": "food",
"channel": "Insta",
"contact": "Insta",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Soûl Made Lille",
"notes": "",
"owner": "liam",
"stage": "a_contacter"
},
"lofigirl": {
"ask": "Merch ou visibilité",
"cat": "paris",
"channel": "Mail",
"contact": "",
"createdAt": "2026-11-03",
"lastContact": null,
"name": "Lofi Girl",
"notes": "",
"owner": "gabriel",
"stage": "a_contacter"
}
},
"referrals": {
"r09": {
"by": "chloe",
"offer": "Hello Bank",
"person": "Maxence",
"status": "demandé"
},
"r14": {
"by": "chloe",
"offer": "Hello Bank",
"person": "Edmond",
"status": "demandé"
},
"r06": {
"by": "mathis",
"offer": "Hello Bank",
"person": "Nathan",
"status": "à demander"
},
"r05": {
"by": "mathis",
"offer": "Hello Bank",
"person": "Juliette",
"status": "à demander"
},
"r01": {
"by": "mathis",
"offer": "Hello Bank",
"person": "Alban",
"status": "en attente"
},
"r07": {
"by": "mathis",
"offer": "Hello Bank",
"person": "Hugo",
"status": "en attente"
},
"r13": {
"by": "chloe",
"offer": "Hello Bank",
"person": "Hippo (Bruneau)",
"status": "demandé"
},
"r10": {
"by": "chloe",
"offer": "Hello Bank",
"person": "Hippo B. (Blaneuil)",
"status": "demandé"
},
"r08": {
"by": "mathis",
"offer": "Hello Bank",
"person": "Jade",
"status": "validé"
},
"r16": {
"by": "chloe",
"offer": "Hello Bank",
"person": "Chloé L.",
"status": "en attente"
},
"r11": {
"by": "chloe",
"offer": "Hello Bank",
"person": "Blanche",
"status": "en attente"
},
"r02": {
"by": "mathis",
"offer": "Hello Bank",
"person": "Liam",
"status": "à demander"
},
"r03": {
"by": "mathis",
"offer": "Hello Bank",
"person": "Gabriel",
"status": "en attente"
},
"r15": {
"by": "chloe",
"offer": "Hello Bank",
"person": "Maxime",
"status": "demandé"
},
"r12": {
"by": "chloe",
"offer": "Hello Bank",
"person": "Romy",
"status": "en attente"
},
"r04": {
"by": "mathis",
"offer": "Hello Bank",
"person": "Lily-Rose",
"status": "à demander"
}
},
"tasks": {
"t06": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-10",
"notes": "Voir avec le bureau : dates des concerts, soirées, expos du mandat.",
"owner": "mathis",
"prio": "normale",
"title": "Fixer le calendrier des événements à vendre aux sponsors"
},
"t08": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-10",
"notes": "Script dans Boîte à outils > Script terrain. Plaquette imprimée.",
"owner": "chloe",
"prio": "normale",
"title": "Démarcher Mandebar, Le Circus et Traknard en duo"
},
"t13": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-05",
"notes": "",
"owner": "alban",
"prio": "normale",
"title": "Mettre à jour les filleuls dans Parrainages"
},
"t19": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-08",
"notes": "",
"owner": "gabriel",
"prio": "normale",
"title": "Finaliser les comptes paris à créer"
},
"t18": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-09",
"notes": "Ex. marques d'instruments, de pinceaux, magasins de musique lillois.",
"owner": "hippolyte",
"prio": "normale",
"title": "Ajouter 5 marques de matériel musique ou art au pipeline"
},
"t10": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-07",
"notes": "Générateur de mail dans la boîte à outils.",
"owner": "stanislas",
"prio": "haute",
"title": "Envoyer les mails Splendid, Zénith et Opéra"
},
"t07": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-05",
"notes": "Repère ce qui n'est pas clair pour les autres.",
"owner": "chloe",
"prio": "normale",
"title": "Tester le QG de A à Z avant la réu"
},
"t16": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-06",
"notes": "Onglet Terrain.",
"owner": "liam",
"prio": "normale",
"title": "Fixer une date pour le shift Vauban avec Hippolyte"
},
"t04": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-05",
"notes": "7 500 ou 8 000 étudiants ? 15 000 ou 25 000 vues ? 800 ou 1 200 participants ? Aligner mail, plaquette et site.",
"owner": "mathis",
"prio": "normale",
"title": "Trancher les chiffres officiels"
},
"t09": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-08",
"notes": "La vue Bureau montre qui l'a fait.",
"owner": "chloe",
"prio": "normale",
"title": "Relancer chaque dimanche ceux qui n'ont pas fait leur report"
},
"t20": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-09",
"notes": "",
"owner": "gabriel",
"prio": "normale",
"title": "Contacter Team Break et relancer Le Buzz Life"
},
"t14": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-06",
"notes": "Esquermoise@benscookies.com",
"owner": "liam",
"prio": "normale",
"title": "Relancer Ben's Cookies (boîte pour les gagnants)"
},
"t11": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-06",
"notes": "Dans le même fil que le mail de Mathis.",
"owner": "stanislas",
"prio": "normale",
"title": "Relancer l'Aéronef et la Bulle Café"
},
"t03": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-04",
"notes": "Lien de l'application web (Déployer > Gérer les déploiements). Remplir les mails dans l'onglet Membres du Sheet pour les rappels.",
"owner": "mathis",
"prio": "haute",
"title": "Envoyer le lien du site à l'équipe"
},
"t17": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-06",
"notes": "suzie.kempa@fnacdarty.com",
"owner": "hippolyte",
"prio": "normale",
"title": "Relancer Fnac (Suzie Kempa) pour des places de concert"
},
"t15": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-07",
"notes": "En physique, argument : kebab le plus proche de l'IÉSEG.",
"owner": "liam",
"prio": "haute",
"title": "Proposer le deal Ô Délice : 50 étudiants, 25 menus un mercredi"
},
"t12": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-06",
"notes": "Blanche, Romy, Chloé L. ont commencé : il faut qu'ils finalisent.",
"owner": "alban",
"prio": "normale",
"title": "Relancer les Hello Bank en cours"
},
"t05": {
"by": "mathis",
"createdAt": "2026-11-03",
"done": false,
"due": "2026-11-05",
"notes": "Présenter le site, les portefeuilles et les règles. Tout le monde ouvre le site pendant la réu.",
"owner": "mathis",
"prio": "normale",
"title": "Caler la réu de lancement du QG (30 min)"
}
},
"offers": {
"parionssport": {
"code": "Lien dans le sheet PARRAINAGES",
"conditions": "",
"etapes": "1. S'inscrire via le lien (Parions Sport en ligne)\n2. Valider son identité\n3. Nous prévenir pour passer en validé",
"gain": "",
"kind": "paris",
"name": "Parions Sport",
"order": 9,
"pieges": "",
"pitch": ""
},
"boursorama": {
"code": "",
"conditions": "Mettre 50 € sur le compte.",
"etapes": "1. S'inscrire via le lien de parrainage\n2. Envoyer les justificatifs\n3. Verser 50 €\n4. Nous prévenir pour passer en validé",
"gain": "",
"kind": "banque",
"name": "Boursorama",
"order": 3,
"pieges": "Lien à récupérer auprès de Mathis.",
"pitch": "Seulement 50 € à mettre, qui restent sur son compte."
},
"hellobank": {
"code": "https://www.hellobank.fr/fr/parrainage-filleul/?code=PPART140OCTOBRE26&token=746BC399",
"conditions": "Ouvrir le compte avec le lien de parrainage.\nFaire un paiement par mois avec la carte.",
"etapes": "1. Ouvrir le lien sur son téléphone\n2. Remplir le formulaire et envoyer sa pièce d'identité\n3. Activer la carte à réception\n4. Faire au moins un paiement par mois\n5. Nous prévenir pour passer en validé",
"gain": "",
"kind": "banque",
"name": "Hello Bank",
"order": 1,
"pieges": "Tester le lien avant de l'envoyer.\nSans paiement mensuel, le parrainage risque de ne pas être validé.",
"pitch": "Compte en ligne qui s'ouvre en quelques minutes, et ça finance directement les événements de l'asso."
},
"winamax": {
"code": "AIM4B8",
"conditions": "Créer son compte avec le code.",
"etapes": "1. Télécharger l'appli\n2. S'inscrire en entrant le code\n3. Valider son identité\n4. Nous prévenir pour passer en validé",
"gain": "",
"kind": "paris",
"name": "Winamax",
"order": 6,
"pieges": "Entrer le code dès l'inscription : impossible de l'ajouter après.",
"pitch": "Création en 5 minutes, ça aide directement l'asso."
},
"pmu": {
"code": "",
"conditions": "Pas de lien de parrainage (appli orange).",
"etapes": "",
"gain": "",
"kind": "paris",
"name": "PMU",
"order": 10,
"pieges": "",
"pitch": ""
},
"revolut": {
"code": "Lien dans le sheet PARRAINAGES",
"conditions": "",
"etapes": "1. S'inscrire via le lien\n2. Valider son identité dans l'appli\n3. Nous prévenir pour passer en validé",
"gain": "",
"kind": "banque",
"name": "Revolut",
"order": 4,
"pieges": "",
"pitch": ""
},
"n26": {
"code": "",
"conditions": "",
"etapes": "",
"gain": "",
"kind": "banque",
"name": "N26",
"order": 5,
"pieges": "Lien à ajouter.",
"pitch": ""
},
"betclic": {
"code": "Lien dans le sheet PARRAINAGES",
"conditions": "",
"etapes": "1. S'inscrire via le lien\n2. Valider son identité\n3. Nous prévenir pour passer en validé",
"gain": "",
"kind": "paris",
"name": "Betclic",
"order": 8,
"pieges": "",
"pitch": ""
},
"unibet": {
"code": "LHCBV",
"conditions": "Créer son compte avec le code.",
"etapes": "1. Télécharger l'appli\n2. S'inscrire en entrant le code\n3. Valider son identité\n4. Nous prévenir pour passer en validé",
"gain": "",
"kind": "paris",
"name": "Unibet",
"order": 7,
"pieges": "",
"pitch": "Création en 5 minutes, ça aide directement l'asso."
},
"fortuneo": {
"code": "302631335",
"conditions": "Verser 300 € sur le compte.\nRemplir les infos de base.\nFaire un paiement par mois.",
"etapes": "1. S'inscrire en entrant le code parrain\n2. Remplir les infos et envoyer les justificatifs\n3. Verser les 300 €\n4. Faire un paiement par mois\n5. Nous prévenir pour passer en validé",
"gain": "",
"kind": "banque",
"name": "Fortuneo",
"order": 2,
"pieges": "Deux codes circulent (302631335 et 14262224) : vérifier lequel utiliser.",
"pitch": "Les 300 € restent sur son compte : c'est son argent, il ne le perd pas."
}
},
"shifts": {
"s02": {
"date": null,
"duo": [
"hippolyte",
"liam"
],
"status": "prévu",
"zone": "Vauban"
},
"s01": {
"date": "2026-09-17",
"duo": [
"mathis",
"chloe"
],
"status": "fait",
"zone": "Vieux-Lille"
}
},
"members": {
"mathis": {
"helloBank": "2026-11-03"
}
}
};
