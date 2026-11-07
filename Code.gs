/**
 * OCDS Monthly Collection System - Google Sheets backend.
 * Setup: open the treasurer's Google Sheet > Extensions > Apps Script > paste this file > Save.
 * Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone > Deploy.
 * Copy the Web app URL (ends in /exec) into the app: Settings & Backup.
 * After editing this script later, use Deploy > Manage deployments > Edit > New version.
 */
const CHUNK = 40000;

function doGet(e) {
  const p = e.parameter || {};
  let out;
  try { out = p.action === "load" ? { ok: true, data: readData_() } : { ok: true, message: "OCDS backend is running" }; }
  catch (err) { out = { ok: false, error: String(err) }; }
  const s = JSON.stringify(out), cb = /^[\w.$]+$/.test(p.callback || "") ? p.callback : "";
  return ContentService.createTextOutput(cb ? cb + "(" + s + ")" : s)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === "save") writeData_(b.data);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

function sheet_(n) { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(n) || ss.insertSheet(n); }

function readData_() {
  const s = SpreadsheetApp.getActive().getSheetByName("AppData");
  if (!s || s.getLastRow() < 1) return null;
  const json = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => r[0]).join("");
  return json ? JSON.parse(json) : null;
}

function put_(name, head, rows, textCol) {
  const s = sheet_(name), all = [head].concat(rows);
  s.clear();
  if (textCol) s.getRange(1, textCol, all.length, 1).setNumberFormat("@");
  s.getRange(1, 1, all.length, head.length).setValues(all);
  s.getRange(1, 1, 1, head.length).setFontWeight("bold").setBackground("#5b3d28").setFontColor("#ffffff");
  s.setFrozenRows(1);
  s.autoResizeColumns(1, head.length);
}

function writeData_(d) {
  const tz = Session.getScriptTimeZone(), fmt = t => Utilities.formatDate(new Date(t), tz, "yyyy-MM-dd HH:mm");
  const members = d.members || [], keys = Object.keys(d.records || {}).sort();
  put_("Members", ["No.", "Member"], members.map((m, i) => [i + 1, m]));
  put_("Payables", ["Payable"], (d.payables || []).map(p => [p]));
  const rec = [], pay = [], sum = {};
  keys.forEach(k => {
    const [mo, id] = k.split("|"), r = d.records[k], name = members[Number(id)] || "Former member";
    const cats = Object.keys(Object.assign({}, r.due, r.paid));
    cats.forEach(c => { const du = Number(r.due[c] || 0), pd = Number(r.paid[c] || 0); rec.push([mo, name, c, du, pd, du - pd]); (sum[mo] = sum[mo] || { due: 0, paid: 0, n: 0 }); sum[mo].due += du; sum[mo].paid += pd; });
    sum[mo] = sum[mo] || { due: 0, paid: 0, n: 0 }; sum[mo].n++;
    (r.transactions || []).forEach(t => pay.push([t.receipt, fmt(t.at), mo, name, Number(t.amount), t.method || "", t.note || "", Number(t.unapplied || 0)]));
  });
  put_("Monthly Records", ["Month", "Member", "Payable", "Amount due", "Paid", "Balance"], rec, 1);
  put_("Payments", ["Receipt", "Date", "Month", "Member", "Amount", "Method", "Note", "Unapplied credit"], pay, 3);
  put_("Audit Trail", ["Date", "Month", "Member", "Activity", "Details", "Amount"],
    (d.audit || []).map(a => [fmt(a.at), a.month || "", members[a.id] || "", a.activity, a.details || "", Number(a.amount || 0)]), 2);
  put_("Dashboard", ["Month", "Members with records", "Dues set this month", "Collected", "Balance of month"],
    Object.keys(sum).sort().map(m => [m, sum[m].n, sum[m].due, sum[m].paid, sum[m].due - sum[m].paid]), 1);
  const json = JSON.stringify(d), parts = [], a = sheet_("AppData");
  for (let i = 0; i < json.length; i += CHUNK) parts.push([json.slice(i, i + CHUNK)]);
  a.clear(); a.getRange(1, 1, parts.length, 1).setNumberFormat("@").setValues(parts);
  a.hideSheet();
}
