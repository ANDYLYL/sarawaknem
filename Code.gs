/**
 * Sarawak NEM – vote & lead database (Google Apps Script)
 * ------------------------------------------------------
 * Paste this into Extensions > Apps Script of a new Google Sheet,
 * then Deploy > New deployment > Web app
 *   Execute as: Me    |    Who has access: Anyone
 * Copy the Web app URL into SHEET_API in index.html.
 *
 * Sheets created automatically:
 *   Votes       – one row per homebuyer vote
 *   Developers  – one row per developer partner sign-up
 * The public stats endpoint returns totals only (no names / emails / phones).
 */

const VOTE_HEADERS = ['Timestamp', 'Name', 'Email', 'Phone', 'Buyer status', 'Area', 'Monthly bill', 'Consent', 'Page'];
const DEV_HEADERS  = ['Timestamp', 'Company', 'Representative', 'Email', 'Phone', 'Project / units', 'Consent', 'Page'];

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0B2545').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v, max) {
  v = String(v == null ? '' : v).trim().slice(0, max || 120);
  // stop spreadsheet formula injection
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    if (p.website) return json_({ ok: true });            // honeypot: bots fill this, people don't
    if (p.consent !== true) return json_({ ok: false, error: 'consent' });

    if (p.type === 'vote') {
      if (!/^\S+@\S+\.\S+$/.test(p.email || '')) return json_({ ok: false, error: 'email' });
      sheet_('Votes', VOTE_HEADERS).appendRow([new Date(), clean_(p.name), clean_(p.email).toLowerCase(), clean_(p.phone, 30),
        clean_(p.buyer_status), clean_(p.area), clean_(p.bill), 'Yes', clean_(p.page, 200)]);
    } else if (p.type === 'developer') {
      sheet_('Developers', DEV_HEADERS).appendRow([new Date(), clean_(p.company), clean_(p.representative), clean_(p.email).toLowerCase(),
        clean_(p.phone, 30), clean_(p.project, 200), 'Yes', clean_(p.page, 200)]);
    } else {
      return json_({ ok: false, error: 'type' });
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/** Public, anonymised statistics for the website charts. */
function doGet() {
  const sh = sheet_('Votes', VOTE_HEADERS);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, VOTE_HEADERS.length).getValues() : [];
  // keep only the latest vote per email
  const latest = {};
  rows.forEach(r => { const k = String(r[2]).toLowerCase(); if (k) latest[k] = r; });
  const uniq = Object.values(latest);

  const count = (idx) => uniq.reduce((m, r) => { const k = r[idx] || 'Not stated'; m[k] = (m[k] || 0) + 1; return m; }, {});
  const byMonth = {};
  uniq.forEach(r => {
    const d = new Date(r[0]); if (isNaN(d)) return;
    const k = Utilities.formatDate(d, 'Asia/Kuching', 'yyyy-MM');
    byMonth[k] = (byMonth[k] || 0) + 1;
  });

  return json_({
    ok: true,
    total: uniq.length,
    status: count(4),
    area: count(5),
    bill: count(6),
    month: byMonth,
    updated: new Date().toISOString()
  });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
