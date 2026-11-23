// ===== Bubur Ayam Koki Gendut – daily sales =====
// Prices live in the 'Master Harga' sheet and are edited from the app's
// "Harga" tab. These defaults are only used to fill that sheet the first time.
const DEFAULT_PRICES = {
  Tengger: { lengkap: 17000, bubur: 9000 },
  Darmo:   { lengkap: 19000, bubur: 11000 },
  Benowo:  { lengkap: 17000, bubur: 9000 }
};

const SHEET_NAME = 'Daily Sales';
// New columns go at the END so existing data keeps its place.
const HEADERS = [
  'Date', 'Branch', 'Bawa Porsi', 'Jam Mulai', 'Jam Selesai', 'Sisa Porsi',
  'Laku Lengkap', 'Laku Bubur Saja', 'Revenue Lengkap', 'Revenue Bubur',
  'Total Uang', 'Note', 'Saved at', 'Emailed at',
  'Revenue Add On', 'Modal Add On'
];
// Column number (1-based) by header name, e.g. COL('Total Uang') -> 11
function COL(name) { return HEADERS.indexOf(name) + 1; }

// One row per add-on sold, linked to the daily row by Date + Branch + Saved at
const ADDON_SHEET_NAME = 'Add On Sales';
const ADDON_HEADERS = [
  'Date', 'Branch', 'Add On', 'Laku', 'Harga Jual', 'Harga Modal',
  'Subtotal', 'Total Modal', 'Untung', 'Saved at'
];

const PRICE_SHEET_NAME = 'Master Harga';
const PRICE_HEADERS = ['Branch', 'Harga Lengkap', 'Harga Bubur Saja', 'Updated at'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Koki Gendut Sales')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Make sure the header row has every column in `headers` (adds new ones to old sheets)
function ensureHeaders(sh, headers) {
  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }
  if (sh.getLastColumn() < headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#fde8d6');
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#fde8d6');
    // Keep times as plain text so Sheets doesn't reformat them
    sh.getRange(2, COL('Jam Mulai'), sh.getMaxRows() - 1, 2).setNumberFormat('@');
  }
  ensureHeaders(sh, HEADERS);
  return sh;
}

function getAddonSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ADDON_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ADDON_SHEET_NAME);
    sh.appendRow(ADDON_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, ADDON_HEADERS.length).setFontWeight('bold').setBackground('#fde8d6');
  }
  return sh;
}

function getPriceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PRICE_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PRICE_SHEET_NAME);
    sh.appendRow(PRICE_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, PRICE_HEADERS.length).setFontWeight('bold').setBackground('#fde8d6');
    const now = new Date();
    const rows = Object.keys(DEFAULT_PRICES).map(b =>
      [b, DEFAULT_PRICES[b].lengkap, DEFAULT_PRICES[b].bubur, now]);
    sh.getRange(2, 1, rows.length, PRICE_HEADERS.length).setValues(rows);
  }
  return sh;
}

// Current prices, in sheet order: { Tengger: { lengkap: 17000, bubur: 9000 }, ... }
function getPrices() {
  const sh = getPriceSheet();
  const prices = {};
  if (sh.getLastRow() < 2) return prices;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    const name = String(r[0]).trim();
    if (name) prices[name] = { lengkap: Number(r[1]) || 0, bubur: Number(r[2]) || 0 };
  });
  return prices;
}

// ----- PIN for the Laporan and Harga tabs -----
// Set it in Project Settings > Script Properties: property PRICE_PIN, value e.g. 4821.
function checkPin_(pin) {
  const real = String(PropertiesService.getScriptProperties().getProperty('PRICE_PIN') || '').trim();
  if (!real) throw new Error('PIN belum diatur. Tambahkan PRICE_PIN di Script Properties.');
  if (String(pin || '').trim() !== real) throw new Error('PIN salah.');
}

// Called by the app to unlock the Laporan and Harga tabs
function verifyPin(pin) {
  checkPin_(pin);
  return true;
}

// Replace the price list. list = [{ branch, lengkap, bubur }, ...]
function savePrices(list, pin) {
  checkPin_(pin);
  if (!Array.isArray(list) || list.length === 0) throw new Error('Minimal 1 cabang.');
  const seen = {};
  const now = new Date();
  const rows = list.map(x => {
    const name = String(x.branch || '').trim();
    const lengkap = Math.round(Number(x.lengkap));
    const bubur = Math.round(Number(x.bubur));
    if (!name) throw new Error('Nama cabang kosong.');
    if (seen[name.toLowerCase()]) throw new Error('Cabang dobel: ' + name);
    seen[name.toLowerCase()] = true;
    if (!(lengkap > 0) || !(bubur > 0)) throw new Error('Harga ' + name + ' harus lebih dari 0.');
    return [name, lengkap, bubur, now];
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getPriceSheet();
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, PRICE_HEADERS.length).clearContent();
    }
    sh.getRange(2, 1, rows.length, PRICE_HEADERS.length).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return getPrices();
}

function saveSale(data) {
  const p = getPrices()[data.branch];
  if (!p) throw new Error('Unknown branch: ' + data.branch);
  const n = v => Math.max(0, Number(v) || 0);

  const lengkap = n(data.lengkap);
  const bubur = n(data.bubur);
  const revL = lengkap * p.lengkap;
  const revB = bubur * p.bubur;

  const addons = (Array.isArray(data.addons) ? data.addons : [])
    .map(a => ({
      name: String(a.name || '').trim(),
      qty: Math.floor(n(a.qty)),
      jual: n(a.jual),
      modal: n(a.modal)
    }))
    .filter(a => a.name || a.qty || a.jual || a.modal);
  addons.forEach(a => { if (!a.name) throw new Error('Add on name is empty'); });

  const revAddon = addons.reduce((s, a) => s + a.qty * a.jual, 0);
  const modalAddon = addons.reduce((s, a) => s + a.qty * a.modal, 0);
  const total = revL + revB + revAddon;
  const savedAt = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    getSheet().appendRow([
      data.date,
      data.branch,
      n(data.bawa),
      String(data.jamMulai || ''),
      String(data.jamSelesai || ''),
      n(data.sisa),
      lengkap,
      bubur,
      revL,
      revB,
      total,
      String(data.note || ''),
      savedAt,
      '', // Emailed at – filled in by Email.gs
      revAddon,
      modalAddon
    ]);
    if (addons.length) {
      const rows = addons.map(a => [
        data.date, data.branch, a.name, a.qty, a.jual, a.modal,
        a.qty * a.jual, a.qty * a.modal, a.qty * (a.jual - a.modal), savedAt
      ]);
      const ash = getAddonSheet();
      ash.getRange(ash.getLastRow() + 1, 1, rows.length, ADDON_HEADERS.length).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
  return total;
}

// ----- Laporan: sales between two dates (inclusive), 'yyyy-MM-dd' -----
// Needs the PIN. Returns plain values only (google.script.run can't send Date objects).
function getSales(from, to, pin) {
  checkPin_(pin);
  const isDay = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s));
  if (!isDay(from) || !isDay(to)) throw new Error('Tanggal tidak valid.');
  if (from > to) { const t = from; from = to; to = t; }

  const tz = Session.getScriptTimeZone();
  const day = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x || '').trim();
  const stamp = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd HH:mm:ss') : String(x || '');
  const time = x => x instanceof Date ? Utilities.formatDate(x, tz, 'HH:mm') : String(x || '');
  const num = x => Number(x) || 0;
  const inRange = d => d >= from && d <= to;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Add-ons, keyed by date | branch | saved at (the same values saveSale wrote to both sheets)
  const addonsByKey = {};
  const ash = ss.getSheetByName(ADDON_SHEET_NAME);
  if (ash && ash.getLastRow() > 1) {
    ash.getRange(2, 1, ash.getLastRow() - 1, ADDON_HEADERS.length).getValues().forEach(r => {
      const d = day(r[0]);
      if (!inRange(d)) return;
      const key = d + '|' + r[1] + '|' + stamp(r[9]);
      (addonsByKey[key] = addonsByKey[key] || []).push({
        name: String(r[2]), qty: num(r[3]), jual: num(r[4]), modal: num(r[5]),
        subtotal: num(r[6]), untung: num(r[8])
      });
    });
  }

  const entries = [];
  const sh = ss.getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 1) {
    ensureHeaders(sh, HEADERS);
    const v = (r, name) => r[COL(name) - 1];
    sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues().forEach(r => {
      const d = day(v(r, 'Date'));
      if (!d || !inRange(d)) return;
      const branch = String(v(r, 'Branch'));
      const saved = stamp(v(r, 'Saved at'));
      entries.push({
        date: d,
        branch: branch,
        bawa: num(v(r, 'Bawa Porsi')),
        mulai: time(v(r, 'Jam Mulai')),
        selesai: time(v(r, 'Jam Selesai')),
        sisa: num(v(r, 'Sisa Porsi')),
        lengkap: num(v(r, 'Laku Lengkap')),
        bubur: num(v(r, 'Laku Bubur Saja')),
        revL: num(v(r, 'Revenue Lengkap')),
        revB: num(v(r, 'Revenue Bubur')),
        revAddon: num(v(r, 'Revenue Add On')),
        modalAddon: num(v(r, 'Modal Add On')),
        total: num(v(r, 'Total Uang')),
        note: String(v(r, 'Note') || ''),
        savedAt: saved,
        addons: addonsByKey[d + '|' + branch + '|' + saved] || []
      });
    });
  }
  entries.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : (a.savedAt < b.savedAt ? -1 : 1));
  return { from: from, to: to, entries: entries };
}
