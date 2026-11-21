const SHEET_ID = '1zd2bv18MrdMhlUofCRXj71Y3NZ_JC5aEKlhPI_lLDMo';
const ENTRIES = 'Entries', PAYMENTS = 'Payments', VEHICLES = 'Vehicles';
// Pump column aakhir (H) me hai, taaki purani entries na bigden
const E_HEAD = ['Timestamp', 'Date', 'Vehicle No', 'Diesel Liter', 'Rate', 'Amount', 'Slip No', 'Pump'];
const P_HEAD = ['Timestamp', 'Date', 'Pump', 'Amount', 'Mode', 'Reference'];

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === 'init') return json_(getInit());
    if (p.action === 'monthly') return json_(getMonthly(String(p.month || '')));
    return json_({ ok: true, app: 'ATC Diesel API' });
  } catch (err) { return json_({ error: String(err.message || err) }); }
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'addEntry') return json_(addEntry(b.data));
    if (b.action === 'addPayment') return json_(addPayment(b.data));
    return json_({ error: 'Unknown action' });
  } catch (err) { return json_({ error: String(err.message || err) }); }
}

function mk_(ss, name, head, textCols) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
    textCols.forEach(c => sh.getRange(c + ':' + c).setNumberFormat('@'));
  }
  return sh;
}

function setup_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const e = mk_(ss, ENTRIES, E_HEAD, ['B', 'G']);
  if (e.getRange(1, 8).getValue() !== 'Pump') e.getRange(1, 8).setValue('Pump').setFontWeight('bold');
  mk_(ss, PAYMENTS, P_HEAD, ['B', 'F']);
  mk_(ss, VEHICLES, ['Vehicle No'], []);
  return ss;
}

const today_ = () => Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
const num_ = v => Number(v) || 0;
const rnd_ = x => Math.round(x * 100) / 100;
const pad_ = x => ('0' + x).slice(-2);
function dstr_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]);
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
  if (m) return m[3] + '-' + pad_(m[2]) + '-' + pad_(m[1]);
  return s;
}

function rows_(sh, n) {
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, n).getValues();
}

function vehicleList_() {
  return rows_(setup_().getSheetByName(VEHICLES), 1).map(r => String(r[0]).trim()).filter(String);
}

function getInit() { return { vehicles: vehicleList_(), today: today_() }; }

function addVehicle_(no) {
  no = String(no || '').trim().toUpperCase();
  if (!no) throw new Error('Vehicle number is required');
  if (vehicleList_().map(x => x.toUpperCase()).indexOf(no) === -1) {
    setup_().getSheetByName(VEHICLES).appendRow([no]);
  }
  return no;
}

function save_(sheetName, row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { setup_().getSheetByName(sheetName).appendRow(row); } finally { lock.releaseLock(); }
}

function addEntry(d) {
  const liters = num_(d.liters), amount = num_(d.amount), slip = String(d.slip || '').trim();
  if (!d.date || !d.vehicle || !d.pump || !(liters > 0) || !(amount > 0) || !slip) {
    throw new Error('Please fill pump, vehicle, litres, amount and slip number');
  }
  save_(ENTRIES, [new Date(), String(d.date), addVehicle_(d.vehicle), liters, num_(d.rate) || '', amount, slip, String(d.pump)]);
  return { ok: true };
}

function addPayment(d) {
  const amount = num_(d.amount);
  if (!d.date || !d.pump || !(amount > 0) || !d.mode) throw new Error('Please fill pump, amount and payment mode');
  save_(PAYMENTS, [new Date(), String(d.date), String(d.pump), amount, String(d.mode), String(d.ref || '').trim()]);
  return { ok: true };
}

// month = 'YYYY-MM' ya 'all'
function getMonthly(month) {
  const ss = setup_(), tz = ss.getSpreadsheetTimeZone();
  const all = month === 'all';
  const due = {}, list = [];
  let tL = 0, tA = 0, tP = 0, nE = 0, nP = 0;
  const key = n => String(n || '').trim() || '—';

  rows_(ss.getSheetByName(ENTRIES), E_HEAD.length).forEach(r => {
    nE++;
    const d = dstr_(r[1], tz), a = num_(r[5]), p = key(r[7]);
    due[p] = (due[p] || 0) + a;
    if (!all && d.substring(0, 7) !== month) return;
    const l = num_(r[3]);
    tL += l; tA += a;
    list.push({ type: 'diesel', ts: +new Date(r[0]) || 0, date: d, vehicle: String(r[2]), pump: p,
      slip: String(r[6]), liters: l, rate: num_(r[4]), amount: a });
  });

  rows_(ss.getSheetByName(PAYMENTS), P_HEAD.length).forEach(r => {
    nP++;
    const d = dstr_(r[1], tz), a = num_(r[3]), p = key(r[2]);
    due[p] = (due[p] || 0) - a;
    if (!all && d.substring(0, 7) !== month) return;
    tP += a;
    list.push({ type: 'payment', ts: +new Date(r[0]) || 0, date: d, pump: p, mode: String(r[4]),
      ref: String(r[5]), amount: a });
  });

  list.sort((x, y) => y.date.localeCompare(x.date) || y.ts - x.ts);
  return {
    totalLiters: rnd_(tL), totalAmount: rnd_(tA), totalPaid: rnd_(tP),
    totalDue: rnd_(Object.keys(due).reduce((s, k) => s + due[k], 0)),
    count: { diesel: nE, payments: nP },
    entries: list
  };
}
