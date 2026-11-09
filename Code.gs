/**
 * Latihan CAT CPNS ASM — penyimpanan nilai di Google Sheets.
 * Tempel seluruh isi file ini di Apps Script (Extensions > Apps Script) dari Google Sheets Anda,
 * lalu Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */
const SHEET = 'Skor';
const KATEGORI = ['TWK', 'TIU', 'TKP', 'UJIAN', 'PU', 'WK', 'BING', 'MTK'];
const HEADER = ['Waktu', 'Nama', 'Kategori', 'Nilai', 'Benar', 'Salah', 'Tidak dijawab', 'Durasi (detik)', 'ID'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Mengambil semua nilai (dipakai halaman ranking). */
function doGet(e) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('rows');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);

  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  let rows = [];
  if (n > 0) {
    rows = sh.getRange(2, 1, n, 8).getValues()
      .filter(r => r[1] !== '' && r[2] !== '')
      .map(r => [
        r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        String(r[1]).replace(/^'/, ''), String(r[2]),
        Number(r[3]), Number(r[4]), Number(r[5]), Number(r[6]), Number(r[7])
      ]);
  }
  const text = JSON.stringify({ ok: true, rows: rows });
  if (text.length < 95000) cache.put('rows', text, 20);
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

/** Menyimpan satu nilai baru dari website. */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Data tidak terbaca' }); }

  const int = (v, min, max) => Number.isInteger(v) && v >= min && v <= max;
  const nama = String(d.nama || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  const id = String(d.id || '');
  if (nama.length < 2) return json_({ ok: false, error: 'Nama tidak valid' });
  if (KATEGORI.indexOf(d.kat) < 0) return json_({ ok: false, error: 'Kategori tidak valid' });
  const total = d.benar + d.salah + d.kosong;
  if (!int(d.nilai, 0, 500) || !int(d.benar, 0, 100) || !int(d.salah, 0, 100) || !int(d.kosong, 0, 100)
      || (total !== 50 && total !== 100) || !int(d.durasi, 0, 86400)) {
    return json_({ ok: false, error: 'Nilai tidak valid' });
  }
  if (!/^[a-z0-9]{6,40}$/.test(id)) return json_({ ok: false, error: 'ID tidak valid' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    // Cegah nilai yang sama tersimpan dua kali (misalnya saat kirim ulang setelah koneksi putus)
    const dup = sh.getRange('I:I').createTextFinder(id).matchEntireCell(true).findNext();
    if (!dup) {
      const aman = /^[=+\-@]/.test(nama) ? "'" + nama : nama;
      sh.appendRow([new Date(), aman, d.kat, d.nilai, d.benar, d.salah, d.kosong, d.durasi, id]);
    }
    CacheService.getScriptCache().remove('rows');
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
