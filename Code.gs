// ===== Bruslení škol – rezervace ledu (Google Apps Script) =====
// 1) Změňte PIN správce.  2) Doplňte svůj e-mail.  3) Spusťte jednou funkci setup.
const ADMIN_PIN = "1234";                 // PIN správce ledu – ZMĚŇTE
const NOTIFY_EMAIL = "";                  // váš e-mail pro upozornění (prázdné = neposílat)
const NOTIFY_SCHOOLS = false;             // true = při uvolnění termínu napsat všem školám s e-mailem

const START = "2026-11-05", END = "2027-03-31";
const ID_RE = /^20\d\d-\d\d-\d\d-(am|pm)$/;
const T_HEAD = ["id", "status", "releasedBy", "school", "contact", "reason", "at", "bookedBy"];
const DAYS = ["Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek"];

// ---------- Jednorázové založení listů ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Rozpis")) {
    const sh = ss.insertSheet("Rozpis");
    const odd = [["ZŠ B. Němcové", "ZŠ Žalhostice"], ["ZŠ Na Valech", "Ploskovice"], ["MŠ", "Křešice / dle domluvy Waldorfská škola"],
                 ["ZŠ Ladova", "ZŠ U Stadionu"], ["ZŠ Havlíčkova", "ZŠ Masarykova / dle domluvy Lingua Universal"]];
    const even = [["ZŠ Masarykova", "ZŠ B. Němcové"], ["Speciální školy", "ZŠ Havlíčkova"], ["MŠ", ""],
                  ["ZŠ U Stadionu", "ZŠ Na Valech"], ["Lingua Universal / dle domluvy další ZŠ", "ZŠ Masarykova"]];
    const rows = [["Týden", "Den", "8.00–9.30", "12.00–13.30"]];
    odd.forEach((r, i) => rows.push(["lichý", DAYS[i], r[0], r[1]]));
    even.forEach((r, i) => rows.push(["sudý", DAYS[i], r[0], r[1]]));
    sh.getRange(1, 1, rows.length, 4).setValues(rows);
    sh.getRange(1, 1, 1, 4).setFontWeight("bold"); sh.setFrozenRows(1); sh.autoResizeColumns(1, 4);
  }
  if (!ss.getSheetByName("Skoly")) {
    const sh = ss.insertSheet("Skoly");
    const names = ["Lingua Universal", "MŠ", "Křešice", "Ploskovice", "Speciální školy", "Svobodná škola", "Waldorfská škola",
      "ZŠ B. Němcové", "ZŠ Havlíčkova", "ZŠ Ladova", "ZŠ Masarykova", "ZŠ Na Valech", "ZŠ U Stadionu", "ZŠ Žalhostice"];
    const used = {};
    const rows = [["Škola (musí se shodovat s textem v Rozpisu)", "Kód", "E-mail", "Chodí od"]];
    names.forEach(n => { let c; do { c = String(Math.floor(100000 + Math.random() * 900000)); } while (used[c]); used[c] = 1; rows.push([n, c, "", n === "Lingua Universal" ? new Date(2027, 0, 1) : ""]); });
    sh.getRange(1, 1, rows.length, 4).setValues(rows);
    sh.getRange(2, 4, rows.length - 1, 1).setNumberFormat("d. m. yyyy");
    sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat("@");
    sh.getRange(1, 1, 1, 4).setFontWeight("bold"); sh.setFrozenRows(1); sh.autoResizeColumns(1, 4);
  }
  if (!ss.getSheetByName("Terminy")) {
    const sh = ss.insertSheet("Terminy");
    sh.getRange(1, 1, 1, T_HEAD.length).setValues([T_HEAD]).setFontWeight("bold"); sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName("Historie")) {
    const sh = ss.insertSheet("Historie");
    sh.getRange(1, 1, 1, 5).setValues([["Čas", "Akce", "Termín", "Kdo", "Poznámka"]]).setFontWeight("bold"); sh.setFrozenRows(1);
  }
}

// ---------- Pomocné ----------
const ss_ = () => SpreadsheetApp.getActiveSpreadsheet();
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
const norm_ = s => String(s || "").trim().toLowerCase();

function schedule_() {
  const v = ss_().getSheetByName("Rozpis").getDataRange().getValues().slice(1);
  const out = { odd: [], even: [] };
  v.forEach(r => {
    const w = norm_(r[0]);
    const key = w.startsWith("lich") ? "odd" : w.startsWith("sud") ? "even" : null;
    if (key && out[key].length < 5) out[key].push([String(r[2] || "").trim(), String(r[3] || "").trim()]);
  });
  return out;
}
function schools_() {
  return ss_().getSheetByName("Skoly").getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1]).map(r => ({ name: String(r[0]).trim(), code: String(r[1]).trim(), email: String(r[2] || "").trim(),
      from: r[3] instanceof Date ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : "" }));
}
// Škola, která chodí až od určitého data: její termíny jsou do té doby volné
function startMap_() {
  const m = {}; schools_().forEach(s => { if (s.from) m[s.name] = s.from; }); return m;
}
function primary_(base) { return String(base || "").split("/")[0].trim(); }
function schoolByCode_(code) {
  code = String(code || "").trim();
  if (!code) return null;
  return schools_().find(s => s.code === code) || null;
}
function isoWeek_(d) {
  const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  const n = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - n);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil(((t - y0) / 864e5 + 1) / 7);
}
function baseFor_(id, sched) {
  const [y, m, d] = id.slice(0, 10).split("-").map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  const wd = (dt.getUTCDay() + 6) % 7;
  if (wd > 4 || id.slice(0, 10) < START || id.slice(0, 10) > END) return null;
  const row = sched[isoWeek_(dt) % 2 ? "odd" : "even"][wd];
  const base = row ? row[id.endsWith("am") ? 0 : 1] : "";
  const from = startMap_()[primary_(base)];
  return from && id.slice(0, 10) < from ? "" : base;
}
function readAll_() {
  const sh = ss_().getSheetByName("Terminy");
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, T_HEAD.length).getValues().filter(r => r[0])
    .map(r => Object.fromEntries(T_HEAD.map((h, i) => [h, String(r[i])])));
}
function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}
function write_(id, obj) {
  const sh = ss_().getSheetByName("Terminy");
  if (!sh.getRange(1, T_HEAD.length).getValue()) sh.getRange(1, 1, 1, T_HEAD.length).setValues([T_HEAD]).setFontWeight("bold");
  const row = T_HEAD.map(h => h === "id" ? id : String(obj[h] || "").slice(0, 120));
  const r = findRow_(sh, id);
  if (r > 0) sh.getRange(r, 1, 1, T_HEAD.length).setValues([row]); else sh.appendRow(row);
}
function remove_(id) {
  const sh = ss_().getSheetByName("Terminy");
  const r = findRow_(sh, id);
  if (r > 0) sh.deleteRow(r);
}
function log_(action, id, who, note) {
  ss_().getSheetByName("Historie").appendRow([new Date(), action, id, who, note || ""]);
}
function human_(id) {
  const [y, m, d] = id.slice(0, 10).split("-").map(Number);
  const wd = (new Date(Date.UTC(y, m - 1, d)).getUTCDay() + 6) % 7;
  return DAYS[wd] + " " + d + ". " + m + ". " + y + ", " + (id.endsWith("am") ? "8.00–9.30" : "12.00–13.30");
}
function mail_(to, subject, body) {
  if (!to) return;
  try { MailApp.sendEmail(to, subject, body); } catch (e) { /* kvóta nebo chyba – rezervace i tak platí */ }
}
function knownSchool_(name) {
  const n = norm_(name);
  const s = schools_().find(x => norm_(x.name) === n);
  return s ? s.name : "";
}
function mayActFor_(school, base) {
  return !!school && !!base && norm_(base).includes(norm_(school.name));
}

// ---------- Web ----------
function doGet() {
  return json_({ ok: true, base: schedule_(), schools: schools_().map(s => s.name), startFrom: startMap_(), data: readAll_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: "Systém je vytížený, zkuste to za chvíli." });
  try {
    const p = JSON.parse(e.postData.contents || "{}");
    const admin = !!p.pin && p.pin === ADMIN_PIN;
    const school = schoolByCode_(p.code);
    const now = new Date().toISOString();

    if (p.action === "checkPin") return json_(admin ? { ok: true } : { ok: false, error: "Nesprávný PIN." });
    if (p.action === "login") return json_(school ? { ok: true, school: school.name } : { ok: false, error: "Neznámý kód školy." });

    if (!ID_RE.test(p.id || "")) return json_({ ok: false, error: "Neplatný termín." });
    const sched = schedule_();
    const base = baseFor_(p.id, sched);
    if (base === null) return json_({ ok: false, error: "Termín není v rozpisu." });
    const cur = readAll_().find(r => r.id === p.id);
    const status = cur ? cur.status : (base ? "assigned" : "free");
    const who = admin ? "správce" : school ? school.name : "";
    if (!who) return json_({ ok: false, error: "Přihlaste se kódem školy." });
    const when = human_(p.id);

    switch (p.action) {
      case "release": {
        if (status !== "assigned") return json_({ ok: false, error: "Termín už mezitím někdo změnil." });
        if (!admin && !mayActFor_(school, base)) return json_({ ok: false, error: "Tento termín patří škole " + base + ". Uvolnit ho může jen ona nebo správce." });
        write_(p.id, { status: "free", releasedBy: admin ? base : school.name, at: now });
        log_("uvolněno", p.id, who, base);
        mail_(NOTIFY_EMAIL, "Bruslení: uvolněn termín " + when, who + " uvolnil(a) termín " + when + " (v rozpisu: " + base + ").");
        if (NOTIFY_SCHOOLS) schools_().filter(s => s.email && (!school || s.name !== school.name))
          .forEach(s => mail_(s.email, "Volný termín na ledě: " + when, "Uvolnil se termín " + when + ". Rezervovat ho můžete na stránce rozpisu bruslení."));
        break;
      }
      case "book": {
        if (status !== "free") return json_({ ok: false, error: "Termín už mezitím někdo obsadil." });
        const name = knownSchool_(p.school || (school && school.name));
        if (!name) return json_({ ok: false, error: "Vyberte školu ze seznamu." });
        write_(p.id, { status: "booked", releasedBy: (cur && cur.releasedBy) || base, school: name, contact: p.contact, at: now, bookedBy: who });
        log_("rezervováno", p.id, who, name + (p.contact ? ", " + p.contact : ""));
        mail_(NOTIFY_EMAIL, "Bruslení: rezervace " + when, name + " má zarezervovaný termín " + when + " (zadal(a): " + who + ")" + (p.contact ? ". Kontakt: " + p.contact : "."));
        break;
      }
      case "handover": {
        if (status !== "assigned") return json_({ ok: false, error: "Termín už mezitím někdo změnil." });
        if (!admin && !mayActFor_(school, base)) return json_({ ok: false, error: "Tento termín patří škole " + base + ". Předat ho může jen ona nebo správce." });
        const name = knownSchool_(p.school);
        if (!name) return json_({ ok: false, error: "Vyberte školu ze seznamu." });
        const from = admin ? base : school.name;
        write_(p.id, { status: "booked", releasedBy: from, school: name, contact: p.contact, at: now, bookedBy: who });
        log_("předáno", p.id, who, from + " → " + name + (p.contact ? ", " + p.contact : ""));
        mail_(NOTIFY_EMAIL, "Bruslení: předán termín " + when, from + " předal(a) termín " + when + " škole " + name + ".");
        const target = schools_().find(s => s.name === name);
        if (target && target.email) mail_(target.email, "Bruslení: máte termín " + when, "Škola " + from + " vám předala termín na ledě: " + when + ".");
        break;
      }
      case "unbook": {
        if (status !== "booked") return json_({ ok: false, error: "Rezervace už neexistuje." });
        if (!admin && norm_(cur.school) !== norm_(school.name) && norm_(cur.bookedBy) !== norm_(school.name))
          return json_({ ok: false, error: "Zrušit rezervaci může jen škola " + cur.school + ", kdo ji zadal, nebo správce." });
        write_(p.id, { status: "free", releasedBy: cur.releasedBy || base, at: now });
        log_("rezervace zrušena", p.id, who, cur.school);
        mail_(NOTIFY_EMAIL, "Bruslení: zrušena rezervace " + when, who + " zrušil(a) rezervaci termínu " + when + ". Termín je znovu volný.");
        break;
      }
      case "close":
        if (!admin) return json_({ ok: false, error: "Jen správce." });
        write_(p.id, { status: "closed", reason: p.reason || "Led není k dispozici", at: now });
        log_("termín zrušen", p.id, who, p.reason);
        break;
      case "reset":
        if (!admin) return json_({ ok: false, error: "Jen správce." });
        remove_(p.id);
        log_("vráceno podle rozpisu", p.id, who, "");
        break;
      default:
        return json_({ ok: false, error: "Neznámá akce." });
    }
    SpreadsheetApp.flush();
    return json_({ ok: true, data: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: "Chyba serveru: " + err.message });
  } finally {
    lock.releaseLock();
  }
}
