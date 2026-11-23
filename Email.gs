// ===== Daily sales email for Bubur Ayam Koki Gendut =====
// Uses SHEET_NAME, HEADERS, COL(), ensureHeaders() and getPrices() from Code.gs.
//
// How it works:
// - Runs every night around 11:59 p.m.
// - Finds every row that hasn't been emailed yet (empty "Emailed at"),
//   whatever its date, sends one email per sales date, then stamps those rows.
// - If there are no un-emailed rows, no email is sent.
//
// Setup:
// 1. Put your email address below.
// 2. Project Settings: time zone = (GMT+07:00) Jakarta.
// 3. Run setupDailyEmail() ONCE and approve permissions.

const REPORT_EMAIL = 'your.email@gmail.com';

function sendDailyReport() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // avoid clashing with a sale being saved
  try {
    const tz = Session.getScriptTimeZone();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return;
    ensureHeaders(sh, HEADERS); // older sheets may lack the add-on columns

    const lastRow = sh.getLastRow();
    const data = sh.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const v = (r, name) => r[COL(name) - 1];
    const asTime = x => x instanceof Date ? Utilities.formatDate(x, tz, 'HH:mm') : String(x || '');

    const byDate = {}; // date -> { rows: [], entries: [] }
    data.forEach((r, i) => {
      if (v(r, 'Emailed at') !== '') return;           // already emailed
      if (v(r, 'Date') === '' && v(r, 'Branch') === '') return; // blank row
      const d = v(r, 'Date') instanceof Date
        ? Utilities.formatDate(v(r, 'Date'), tz, 'yyyy-MM-dd')
        : String(v(r, 'Date'));
      if (!byDate[d]) byDate[d] = { rows: [], entries: [] };
      byDate[d].rows.push(i + 2);
      byDate[d].entries.push({
        branch: v(r, 'Branch'),
        bawa: Number(v(r, 'Bawa Porsi')) || 0,
        mulai: asTime(v(r, 'Jam Mulai')),
        selesai: asTime(v(r, 'Jam Selesai')),
        sisa: Number(v(r, 'Sisa Porsi')) || 0,
        lengkap: Number(v(r, 'Laku Lengkap')) || 0,
        bubur: Number(v(r, 'Laku Bubur Saja')) || 0,
        addon: Number(v(r, 'Revenue Add On')) || 0,
        total: Number(v(r, 'Total Uang')) || 0,
        note: v(r, 'Note')
      });
    });

    const dates = Object.keys(byDate).sort();
    if (dates.length === 0) return; // nothing new -> no email

    dates.forEach(d => {
      sendEmailForDate(d, byDate[d].entries, ss.getUrl());
      const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
      byDate[d].rows.forEach(rowNum => sh.getRange(rowNum, COL('Emailed at')).setValue(stamp));
    });
  } finally {
    lock.releaseLock();
  }
}

function sendEmailForDate(date, entries, sheetUrl) {
  const rp = n => 'Rp' + Number(n).toLocaleString('id-ID');
  const branches = Object.keys(getPrices());
  // Include any branch name found in the data that isn't in the price list
  entries.forEach(e => { if (branches.indexOf(e.branch) === -1) branches.push(e.branch); });

  let gL = 0, gB = 0, gA = 0, gT = 0;
  const rowsHtml = branches.map(b => {
    const list = entries.filter(e => e.branch === b);
    if (list.length === 0) {
      return '<tr style="color:#b8a596"><td>' + b + '</td><td colspan="7">no entry</td></tr>';
    }
    return list.map(e => {
      gL += e.lengkap; gB += e.bubur; gA += e.addon; gT += e.total;
      return '<tr><td>' + b + '</td>' +
        '<td align="center">' + e.mulai + '–' + e.selesai + '</td>' +
        '<td align="center">' + e.bawa + '</td>' +
        '<td align="center">' + (e.sisa === 0 ? 'Habis' : e.sisa) + '</td>' +
        '<td align="center">' + e.lengkap + '</td>' +
        '<td align="center">' + e.bubur + '</td>' +
        '<td align="right">' + (e.addon ? rp(e.addon) : '–') + '</td>' +
        '<td align="right">' + rp(e.total) + '</td></tr>' +
        (e.note ? '<tr><td></td><td colspan="7" style="color:#8a7462;font-size:13px">📝 ' + e.note + '</td></tr>' : '');
    }).join('');
  }).join('');

  const html =
    '<div style="font-family:Arial,sans-serif;color:#3a2a1a">' +
    '<h2 style="color:#e8792b">🍲 Bubur Ayam Koki Gendut – ' + date + '</h2>' +
    '<table cellpadding="7" style="border-collapse:collapse;border:1px solid #f0e2d2;font-size:14px">' +
    '<tr style="background:#fde8d6"><th align="left">Cabang</th><th>Jam</th><th>Bawa</th><th>Sisa</th>' +
    '<th>Lengkap</th><th>Bubur</th><th align="right">Add on</th><th align="right">Total</th></tr>' +
    rowsHtml +
    '<tr style="font-weight:bold;background:#fff8ef"><td>Total</td><td></td><td></td><td></td>' +
    '<td align="center">' + gL + '</td><td align="center">' + gB + '</td>' +
    '<td align="right">' + rp(gA) + '</td>' +
    '<td align="right">' + rp(gT) + '</td></tr>' +
    '</table>' +
    '<p><a href="' + sheetUrl + '">Open the full sheet</a></p></div>';

  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: 'Koki Gendut sales ' + date + ': ' + rp(gT),
    htmlBody: html
  });
}

// Run this ONCE to schedule the nightly email (about 11:59 p.m.).
function setupDailyEmail() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(59)
    .create();
}
