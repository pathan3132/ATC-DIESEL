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
    if (p.action === 'report') return json_(getReport(String(p.from || ''), String(p.to || '')));
    if (p.action === 'monthly') return json_(getMonthly(String(p.month || '')));
    return json_({ ok: true, app: 'ATC Diesel API' });
  } catch (err) { return json_({ error: String(err.message || err) }); }
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'addEntry') return json_(addEntry(b.data));
    if (b.action === 'addPayment') return json_(addPayment(b.data));
    if (b.action === 'updateEntry') return json_(updateEntry(b.data));
    if (b.action === 'updatePayment') return json_(updatePayment(b.data));
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

// Purani entry badalna: row number + original timestamp match hona zaroori (galat row na badle)
function updateRow_(name, row, ts, values, textCols) {
  row = Number(row);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = setup_().getSheetByName(name);
    if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('Entry not found. Please refresh and try again.');
    if ((+new Date(sh.getRange(row, 1).getValue()) || 0) !== Number(ts)) {
      throw new Error('This entry was changed or moved. Please refresh and try again.');
    }
    textCols.forEach(col => sh.getRange(row, col).setNumberFormat('@'));
    sh.getRange(row, 2, 1, values.length).setValues([values]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function updateEntry(d) {
  const liters = num_(d.liters), amount = num_(d.amount), slip = String(d.slip || '').trim();
  if (!d.date || !d.vehicle || !d.pump || !(liters > 0) || !(amount > 0) || !slip) {
    throw new Error('Please fill pump, vehicle, litres, amount and slip number');
  }
  const vehicle = addVehicle_(d.vehicle);
  return updateRow_(ENTRIES, d.row, d.ts,
    [String(d.date), vehicle, liters, num_(d.rate) || '', amount, slip, String(d.pump)], [2, 7]);
}

function updatePayment(d) {
  const amount = num_(d.amount);
  if (!d.date || !d.pump || !(amount > 0) || !d.mode) throw new Error('Please fill pump, amount and payment mode');
  return updateRow_(PAYMENTS, d.row, d.ts,
    [String(d.date), String(d.pump), amount, String(d.mode), String(d.ref || '').trim()], [2, 6]);
}

function getMonthly(month) {
  return month === 'all' ? getReport('', '') : getReport(month + '-01', month + '-31');
}

// from / to = 'YYYY-MM-DD' (khali ho to koi seema nahi)
// Har transaction ke saath uske baad ka running outstanding (balance) jaata hai
function getReport(from, to) {
  const ss = setup_(), tz = ss.getSpreadsheetTimeZone();
  const key = n => String(n || '').trim() || '—';
  const tx = [];
  let nE = 0, nP = 0;

  rows_(ss.getSheetByName(ENTRIES), E_HEAD.length).forEach((r, i) => {
    nE++;
    const a = num_(r[5]);
    tx.push({ type: 'diesel', row: i + 2, ts: +new Date(r[0]) || 0, date: dstr_(r[1], tz), vehicle: String(r[2]), pump: key(r[7]),
      slip: String(r[6]), liters: num_(r[3]), rate: num_(r[4]), amount: a, delta: a });
  });
  rows_(ss.getSheetByName(PAYMENTS), P_HEAD.length).forEach((r, i) => {
    nP++;
    const a = num_(r[3]);
    tx.push({ type: 'payment', row: i + 2, ts: +new Date(r[0]) || 0, date: dstr_(r[1], tz), pump: key(r[2]),
      mode: String(r[4]), ref: String(r[5]), amount: a, delta: -a });
  });
  tx.sort((x, y) => x.date.localeCompare(y.date) || x.ts - y.ts);

  let bal = 0, opening = 0, closing = 0, tL = 0, tA = 0, tP = 0;
  const list = [];
  tx.forEach(t => {
    bal += t.delta;
    if (from && t.date < from) opening = bal;
    if (!to || t.date <= to) closing = bal;
    if ((from && t.date < from) || (to && t.date > to)) return;
    if (t.type === 'diesel') { tL += t.liters; tA += t.amount; } else { tP += t.amount; }
    t.balance = rnd_(bal);
    delete t.delta;
    list.push(t);
  });

  return {
    totalLiters: rnd_(tL), totalAmount: rnd_(tA), totalPaid: rnd_(tP),
    opening: rnd_(opening), closing: rnd_(closing), totalDue: rnd_(closing),
    count: { diesel: nE, payments: nP },
    entries: list.reverse()   // nayi entry sabse upar
  };
}
