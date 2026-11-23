// ===== Bubur Ayam Koki Gendut – daily sales =====
// Prices per portion. Keep in sync with PRICES in Index.html.
const PRICES = {
  Tengger: { lengkap: 17000, bubur: 9000 },
  Darmo:   { lengkap: 19000, bubur: 11000 },
  Benowo:  { lengkap: 17000, bubur: 9000 }
};

const SHEET_NAME = 'Daily Sales';
const HEADERS = [
  'Date', 'Branch', 'Bawa Porsi', 'Jam Mulai', 'Jam Selesai', 'Sisa Porsi',
  'Laku Lengkap', 'Laku Bubur Saja', 'Revenue Lengkap', 'Revenue Bubur',
  'Total Uang', 'Note', 'Saved at', 'Emailed at'
];
// Column number (1-based) by header name, e.g. COL('Total Uang') -> 11
function COL(name) { return HEADERS.indexOf(name) + 1; }

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Koki Gendut Sales')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
  return sh;
}

function saveSale(data) {
  const p = PRICES[data.branch];
  if (!p) throw new Error('Unknown branch: ' + data.branch);
  const n = v => Math.max(0, Number(v) || 0);

  const lengkap = n(data.lengkap);
  const bubur = n(data.bubur);
  const revL = lengkap * p.lengkap;
  const revB = bubur * p.bubur;
  const total = revL + revB;

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
      new Date(),
      '' // Emailed at – filled in by Email.gs
    ]);
  } finally {
    lock.releaseLock();
  }
  return total;
}
