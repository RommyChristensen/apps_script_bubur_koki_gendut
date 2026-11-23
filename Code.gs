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

// Replace the price list. list = [{ branch, lengkap, bubur }, ...]
function savePrices(list) {
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
