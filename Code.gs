/**
 * ABSENSI BPS - Google Apps Script (anti CORS)
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me | Who has access: Anyone
 */
const SS_ID = '1nE_7ytn89hpCmFbTlS6UO-l7VZuuxA30zQ-sULspETU';
const SHEET_ABSEN = 'Absensi';

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Cari sheet tanpa peduli huruf besar/kecil
function sheet_(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const found = ss.getSheets().find(s => s.getName().toLowerCase() === name.toLowerCase());
  return found || null;
}

function colA_(name) {
  const sh = sheet_(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim()).filter(String);
}

function absenSheet_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sh = sheet_(SHEET_ABSEN);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ABSEN);
    sh.appendRow(['Timestamp', 'Tanggal', 'Nama', 'Jabatan', 'Alamat', 'Keterangan']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#0b5ed7').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// GET: ?action=init  |  ?action=stats
function doGet(e) {
  try {
    const action = (e.parameter.action || 'init');
    if (action === 'init') {
      return json_({ ok: true, nama: colA_('nama'), alamat: colA_('alamat') });
    }
    if (action === 'stats') {
      const sh = absenSheet_();
      const n = sh.getLastRow() - 1;
      const rows = n > 0 ? sh.getRange(2, 1, n, 6).getValues() : [];
      const tz = Session.getScriptTimeZone();
      const data = rows.map(r => ({
        tanggal: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
        nama: r[2], jabatan: r[3], alamat: r[4], keterangan: r[5]
      }));
      return json_({ ok: true, data: data });
    }
    return json_({ ok: false, error: 'Action tidak dikenal' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// POST (Content-Type text/plain agar tidak ada preflight CORS)
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    if (!d.tanggal || !d.nama || !d.jabatan || !d.alamat || !d.keterangan) {
      return json_({ ok: false, error: 'Data belum lengkap' });
    }
    absenSheet_().appendRow([new Date(), d.tanggal, d.nama, d.jabatan, d.alamat, d.keterangan]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
