/**
 * ABSENSI & KEUANGAN BPS - Google Apps Script (anti CORS)
 * Deploy: Deploy > Manage deployments > Edit (ikon pensil) > Version: New version > Deploy
 *   Execute as: Me | Who has access: Anyone
 *   (Pakai "Manage deployments", BUKAN "New deployment", supaya URL /exec tidak berubah)
 */
const SS_ID = '1nE_7ytn89hpCmFbTlS6UO-l7VZuuxA30zQ-sULspETU';
const SHEET_ABSEN = 'Absensi';
const SHEET_KEU = 'Keuangan';

// ===== Login Laporan Keuangan (dicek di server, tidak tampil di halaman web) =====
const LOGIN_USER = 'bpsrw09';
const LOGIN_PASS = 'transparan';
const TOKEN_TTL = 21600;      // 6 jam (batas maksimum CacheService)
const MAX_GAGAL = 5;          // maksimal salah login berturut-turut
const BLOKIR_DETIK = 600;     // lalu diblokir 10 menit

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Cari sheet tanpa peduli huruf besar/kecil/spasi. Cocok persis dulu, lalu yang mengandung nama.
function sheet_(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const key = name.trim().toLowerCase();
  const all = ss.getSheets();
  return all.find(s => s.getName().trim().toLowerCase() === key)
      || all.find(s => s.getName().trim().toLowerCase().indexOf(key) !== -1)
      || null;
}

function sheetNames_() {
  return SpreadsheetApp.openById(SS_ID).getSheets().map(s => s.getName());
}

// Ambil daftar dari sheet: kolom pertama yang berisi data, header (jika ada) dilewati, duplikat dibuang
function colA_(name) {
  const sh = sheet_(name);
  if (!sh || sh.getLastRow() < 1) return [];
  const vals = sh.getDataRange().getValues();
  let col = -1;
  for (let c = 0; c < vals[0].length && col < 0; c++) {
    if (vals.some(r => String(r[c]).trim() !== '')) col = c;
  }
  if (col < 0) return [];
  let list = vals.map(r => String(r[col]).trim()).filter(String);
  const h = list[0].toLowerCase();
  if (h === name.toLowerCase() || h.indexOf(name.toLowerCase()) !== -1 || h === 'no') list = list.slice(1);
  return list.filter((v, i) => list.indexOf(v) === i);
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

// Sheet "Keuangan" (6 kolom):
// A Tanggal Input | B Transaksi Pemasukan | C Jumlah Pemasukan
// | D Transaksi Pengeluaran | E Jumlah Pengeluaran | F Saldo
// Kolom B / D berisi NAMA transaksinya (mis. "Iuran Warga", "Beli ATK").
// Saldo = rumus otomatis: total Jumlah Pemasukan - total Jumlah Pengeluaran sampai baris itu.
const KEU_HEADER = ['Tanggal Input', 'Transaksi Pemasukan', 'Jumlah Pemasukan',
                    'Transaksi Pengeluaran', 'Jumlah Pengeluaran', 'Saldo'];
const KEU_COLS = KEU_HEADER.length;
const FMT_RP = '"Rp" #,##0;-"Rp" #,##0';

function keuHeader_(sh) {
  sh.getRange(1, 1, 1, KEU_COLS).setValues([KEU_HEADER])
    .setFontWeight('bold').setBackground('#0f766e').setFontColor('#fff').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 105);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(5, 130);
  sh.setColumnWidth(6, 130);
}

// Satu baris data. Pemasukan -> kolom B & C terisi; Pengeluaran -> kolom D & E terisi.
function keuRow_(tgl, transaksi, nama, jumlah) {
  return transaksi === 'Pemasukan'
    ? [tgl, nama, jumlah, '', '', '']
    : [tgl, '', '', nama, jumlah, ''];
}

// Tulis baris mulai dari baris `start`, lengkap dengan format & rumus saldo.
function keuWrite_(sh, start, rows) {
  const n = rows.length;
  sh.getRange(start, 1, n, KEU_COLS).setValues(rows);
  sh.getRange(start, 1, n, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(start, 3, n, 1).setNumberFormat(FMT_RP);
  sh.getRange(start, 5, n, 1).setNumberFormat(FMT_RP);
  const f = [];
  for (let r = start; r < start + n; r++) f.push(['=SUM($C$2:C' + r + ')-SUM($E$2:E' + r + ')']);
  sh.getRange(start, 6, n, 1).setFormulas(f).setNumberFormat(FMT_RP).setFontWeight('bold');
}

// Sheet susunan lama otomatis diubah ke susunan baru (6 kolom). Dikenali 2 susunan lama:
//  (1) Tanggal | Transaksi | Nama Transaksi | Jumlah
//  (2) 7 kolom: Tanggal | Trx Pemasukan | Jml Pemasukan | Trx Pengeluaran | Jml Pengeluaran | Nama Transaksi | Saldo
// Salinan cadangan dibuat dulu supaya data aman.
function keuMigrate_(sh) {
  const last = sh.getLastRow();
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 7)).getValues()[0]
    .map(x => String(x).trim().toLowerCase());
  if (last < 1 || h[0] === '') { keuHeader_(sh); return; }          // sheet kosong
  if (h[1] === 'transaksi pemasukan' && h[5] === 'saldo') return;   // sudah susunan baru

  const lama1 = h[1] === 'transaksi';
  const lama2 = h[1] === 'transaksi pemasukan' && h[5] === 'nama transaksi';
  if (!lama1 && !lama2) return;                                     // tidak dikenal, jangan disentuh

  const tz = Session.getScriptTimeZone();
  sh.copyTo(SpreadsheetApp.openById(SS_ID))
    .setName('Backup_Keuangan_Lama_' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm'));

  const rows = [];
  if (lama1) {
    const old = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
    old.forEach(r => {
      if (r[0] === '' || (r[1] !== 'Pemasukan' && r[1] !== 'Pengeluaran')) return;
      rows.push(keuRow_(r[0], r[1], String(r[2]), Number(r[3]) || 0));
    });
  } else {
    const old = last > 1 ? sh.getRange(2, 1, last - 1, 7).getValues() : [];
    old.forEach(r => {
      if (r[0] === '') return;
      const masuk = Number(r[2]) || 0, keluar = Number(r[4]) || 0;
      if (masuk > 0)  rows.push(keuRow_(r[0], 'Pemasukan',   String(r[5] || r[1]), masuk));
      if (keluar > 0) rows.push(keuRow_(r[0], 'Pengeluaran', String(r[5] || r[3]), keluar));
    });
  }

  sh.clear();
  keuHeader_(sh);
  if (rows.length) keuWrite_(sh, 2, rows);
}

function keuSheet_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sh = sheet_(SHEET_KEU);
  if (!sh) {
    sh = ss.insertSheet(SHEET_KEU);
    keuHeader_(sh);
  } else {
    keuMigrate_(sh);
  }
  return sh;
}

// ---------- Auth ----------
function authed_(token) {
  return !!token && CacheService.getScriptCache().get('tok_' + token) === '1';
}

function login_(d) {
  const cache = CacheService.getScriptCache();
  const gagal = Number(cache.get('login_gagal') || 0);
  if (gagal >= MAX_GAGAL) {
    return { ok: false, error: 'Terlalu banyak percobaan salah. Coba lagi dalam 10 menit.' };
  }
  if (String(d.user) === LOGIN_USER && String(d.pass) === LOGIN_PASS) {
    cache.remove('login_gagal');
    const token = Utilities.getUuid() + Utilities.getUuid();
    cache.put('tok_' + token, '1', TOKEN_TTL);
    return { ok: true, token: token };
  }
  cache.put('login_gagal', String(gagal + 1), BLOKIR_DETIK);
  return { ok: false, error: 'Username atau password salah' };
}

// ---------- Keuangan ----------
function keuList_() {
  const sh = keuSheet_();
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, KEU_COLS).getValues() : [];
  const tz = Session.getScriptTimeZone();
  const data = [];
  rows.forEach(r => {
    if (r[0] === '') return;
    const tanggal = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]);
    const masuk = Number(r[2]) || 0;    // kolom C
    const keluar = Number(r[4]) || 0;   // kolom E
    if (masuk > 0)  data.push({ tanggal: tanggal, transaksi: 'Pemasukan',   nama: String(r[1]), jumlah: masuk });
    if (keluar > 0) data.push({ tanggal: tanggal, transaksi: 'Pengeluaran', nama: String(r[3]), jumlah: keluar });
  });
  return { ok: true, data: data };
}

function keuAdd_(d) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.tanggal))) return { ok: false, error: 'Tanggal tidak valid' };
  if (d.transaksi !== 'Pemasukan' && d.transaksi !== 'Pengeluaran') return { ok: false, error: 'Jenis transaksi tidak valid' };
  const items = Array.isArray(d.items) ? d.items : [];
  if (!items.length || items.length > 50) return { ok: false, error: 'Item transaksi kosong atau terlalu banyak' };

  const p = String(d.tanggal).split('-').map(Number);
  const tgl = new Date(p[0], p[1] - 1, p[2]);
  const rows = [];
  for (const it of items) {
    const nama = String(it.nama || '').trim().slice(0, 120);
    const jumlah = Math.round(Number(it.jumlah));
    if (!nama || !isFinite(jumlah) || jumlah <= 0) return { ok: false, error: 'Nama dan jumlah setiap item wajib diisi' };
    rows.push(keuRow_(tgl, d.transaksi, nama, jumlah));
  }
  const sh = keuSheet_();
  keuWrite_(sh, sh.getLastRow() + 1, rows);
  return { ok: true, saved: rows.length };
}

// GET: ?action=init  |  ?action=stats   (absensi, tidak berubah)
function doGet(e) {
  try {
    const action = (e.parameter.action || 'init');
    if (action === 'init') {
      return json_({ ok: true, nama: colA_('nama'), alamat: colA_('alamat'), sheets: sheetNames_() });
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
// d.action: 'login' | 'keu_list' | 'keu_add' | (kosong = absensi, seperti sebelumnya)
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);

    if (d.action === 'login') return json_(login_(d));

    if (d.action === 'keu_list' || d.action === 'keu_add') {
      if (!authed_(d.token)) return json_({ ok: false, auth: false, error: 'Sesi berakhir, silakan login lagi' });
      return json_(d.action === 'keu_list' ? keuList_() : keuAdd_(d));
    }

    // Absensi (default)
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
