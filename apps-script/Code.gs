/**
 * Belanjawan Kelab Guru - API untuk PWA (GitHub Pages).
 * Tampal dalam Google Sheet anda: Extensions > Apps Script. Tiada paparan HTML; app ada di folder docs/.
 *
 * Tab Transaksi (baris 1 = tajuk, data dari baris 2):
 *   A Tarikh | B Butiran | C Ahli | D Kategori | E Bulan mula | F Bulan akhir |
 *   G Kelab | H Lawatan | I Masuk lain | J Keluar | K Baki (formula) | L Nota | M..R pembantu (formula)
 * App hanya menulis A..J dan L. Lajur formula tidak disentuh.
 */
const DEFAULT_SSID = '1cI6EJXsmrhBIYflDYl12aw5uBT7BiHP-b4Korb-we-A'; // Google Sheet anda
const SHEET_TX = 'Transaksi';
const SHEET_SET = 'Tetapan';
const SHEET_YURAN = 'Yuran';
const NVAL = 10;               // lajur nilai A..J
const NSLOT = 20;              // slot ahli dalam Tetapan / Yuran
const N_MONTHS = 19;           // bilangan bulan dalam grid Yuran
const Y_SUM_ROW = 5;           // Yuran: ringkasan ahli bermula baris 5
const Y_K_HEAD = 28;           // Yuran: baris tajuk bulan jadual Kelab (data 29..48)
const Y_L_HEAD = 52;           // Yuran: baris tajuk bulan jadual Lawatan (data 53..72)

// ------------------------------------------------------------------ titik masuk
function doGet() {
  return json_({ ok: true, data: 'Belanjawan Kelab Guru API aktif.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case 'getData': return json_({ ok: true, data: getData(req.pin) });
      case 'add': return json_({ ok: true, data: addTransactions(req.pin, req.rows) });
      case 'delete': return json_({ ok: true, data: deleteTransaction(req.pin, req.row, req.fp) });
      default: throw new Error('Tindakan tidak dikenali');
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------------ persediaan
/**
 * LANGKAH PERTAMA: tukar 1234 kepada PIN anda (4-8 digit), pilih fungsi "aturPin" di bar atas,
 * tekan Run dan beri kebenaran (Authorize). Wajib sebelum Deploy.
 */
function aturPin() { setPin('1234'); }

function setPin(pin) {
  pin = String(pin);
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN mesti 4-8 digit');
  const ss = ss_();
  [SHEET_TX, SHEET_SET, SHEET_YURAN].forEach(function (n) {
    if (!ss.getSheetByName(n)) throw new Error('Tab "' + n + '" tiada dalam helaian ini');
  });
  const props = PropertiesService.getScriptProperties();
  props.setProperty('PIN', pin);
  props.setProperty('SSID', ss.getId());
  LockService.getScriptLock();
  Logger.log('Siap. Helaian: ' + ss.getName() + '. Seterusnya: Deploy > New deployment > Web app.');
}

/** Run dari editor: mencetak URL web app, akaun pemilik dan sama ada data boleh dibaca. */
function diagnos() {
  let url = '(belum di-deploy)';
  try { url = ScriptApp.getService().getUrl() || url; } catch (e) { url = 'ralat: ' + e; }
  const ss = ss_();
  Logger.log('URL API       : ' + url);
  Logger.log('Pemilik       : ' + Session.getEffectiveUser().getEmail());
  Logger.log('Helaian       : ' + ss.getName());
  Logger.log('PIN ditetapkan: ' + !!PropertiesService.getScriptProperties().getProperty('PIN'));
  const d = getData(PropertiesService.getScriptProperties().getProperty('PIN'));
  Logger.log('OK: ' + d.items.length + ' transaksi, ' + d.ahli.length + ' ahli, ' + d.yuran.rows.length + ' baris yuran');
}

function checkPin_(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!saved) throw new Error('PIN belum ditetapkan. Jalankan aturPin di editor Apps Script.');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 5) throw new Error('Terlalu banyak cubaan salah. Cuba lagi selepas 10 minit.');
  if (String(pin) !== saved) {
    cache.put('fails', String(fails + 1), 600);
    throw new Error('PIN salah');
  }
  if (fails) cache.remove('fails');
}

function ss_() {
  try { const a = SpreadsheetApp.getActive(); if (a) return a; } catch (e) { /* web app: tiada helaian aktif */ }
  const id = PropertiesService.getScriptProperties().getProperty('SSID') || DEFAULT_SSID;
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Tab "' + name + '" tiada');
  return s;
}

// ------------------------------------------------------------------ baca
function getData(pin) {
  checkPin_(pin);
  const ss = ss_(), tz = ss.getSpreadsheetTimeZone();
  const tx = sheet_(SHEET_TX), last = lastRow_(tx);
  const rows = last < 2 ? [] : tx.getRange(2, 1, last - 1, 12).getValues();
  const items = rows.map(function (r, i) {
    const it = {
      row: i + 2, tarikh: fmt_(r[0], tz, 'yyyy-MM-dd'), butiran: String(r[1] || ''), ahli: String(r[2] || ''),
      kategori: String(r[3] || ''), mula: fmt_(r[4], tz, 'yyyy-MM'), akhir: fmt_(r[5], tz, 'yyyy-MM'),
      kelab: num_(r[6]), lawatan: num_(r[7]), lain: num_(r[8]), keluar: num_(r[9]), nota: String(r[11] || '')
    };
    it.fp = fp_(it.tarikh, it.butiran, it.kelab, it.lawatan, it.lain, it.keluar);
    return it;
  });

  const set = sheet_(SHEET_SET);
  const ahli = set.getRange(5, 1, NSLOT, 6).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { nama: String(r[0]), mula: fmt_(r[1], tz, 'yyyy-MM'), tamat: fmt_(r[2], tz, 'yyyy-MM'),
               kelab: num_(r[3]), lawatan: num_(r[4]), nota: String(r[5] || '') };
    });
  const kat = set.getRange(5, 8, 11, 1).getValues().map(function (r) { return String(r[0]); });

  return {
    items: items, ahli: ahli,
    lawatanAkhir: fmt_(set.getRange(27, 2).getValue(), tz, 'yyyy-MM'),
    mulaRekod: fmt_(set.getRange(28, 2).getValue(), tz, 'yyyy-MM'),
    kategori: { Masuk: kat.slice(1, 4), Keluar: kat.slice(4) },
    yuran: yuran_(tz),
    sheetName: ss.getName()
  };
}

/** Status yuran dibaca terus daripada tab Yuran (formula sheet) supaya app sentiasa sama dengan sheet. */
function yuran_(tz) {
  const y = sheet_(SHEET_YURAN);
  if (String(y.getRange(Y_K_HEAD - 1, 1).getValue()) !== 'Kelab' || String(y.getRange(Y_L_HEAD - 1, 1).getValue()) !== 'Lawatan') {
    throw new Error('Susun atur tab Yuran berubah (A27 mesti "Kelab", A51 mesti "Lawatan"). Jangan alih baris.');
  }
  const bulan = y.getRange(Y_K_HEAD, 2, 1, N_MONTHS).getValues()[0].map(function (d) { return fmt_(d, tz, 'yyyy-MM'); });
  const k = y.getRange(Y_K_HEAD + 1, 1, NSLOT, N_MONTHS + 1).getValues();
  const l = y.getRange(Y_L_HEAD + 1, 1, NSLOT, N_MONTHS + 1).getValues();
  const rows = [];
  for (let i = 0; i < NSLOT; i++) {
    if (!k[i][0]) continue;
    rows.push({ nama: String(k[i][0]),
                k: k[i].slice(1).map(String), l: l[i].slice(1).map(String) });
  }
  return { bulan: bulan, rows: rows };
}

// ------------------------------------------------------------------ tulis
/**
 * rows: [{tarikh:'YYYY-MM-DD', butiran, ahli, kategori, mula:'YYYY-MM'|'', akhir:'YYYY-MM'|'',
 *         kelab, lawatan, lain, keluar, nota}]
 */
function addTransactions(pin, rows) {
  checkPin_(pin);
  if (!rows || !rows.length || rows.length > 50) throw new Error('Data tidak sah');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const tx = sheet_(SHEET_TX), tz = ss_().getSpreadsheetTimeZone();
    const last = lastRow_(tx);
    const cur = readBlock_(tx, last);
    const add = rows.map(function (x) { return toRow_(x, tz); });
    if (last - 1 + add.length > 1999) throw new Error('Tab Transaksi penuh (maks 2000 baris)');
    const all = cur.concat(add);
    // susun ikut tarikh; JS sort stabil, jadi susunan asal dalam hari yang sama kekal
    all.sort(function (a, b) { return (+a.v[0] || 0) - (+b.v[0] || 0); });
    writeBlock_(tx, all);
    SpreadsheetApp.flush();
    return add.length;
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction(pin, row, fp) {
  checkPin_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const tx = sheet_(SHEET_TX), tz = ss_().getSpreadsheetTimeZone();
    const last = lastRow_(tx);
    row = Number(row);
    if (!(row >= 2 && row <= last)) throw new Error('Baris tidak dijumpai. Muat semula data.');
    const cur = readBlock_(tx, last), r = cur[row - 2].v;
    if (fp_(fmt_(r[0], tz, 'yyyy-MM-dd'), String(r[1] || ''), num_(r[6]), num_(r[7]), num_(r[8]), num_(r[9])) !== fp) {
      throw new Error('Baris telah berubah dalam helaian. Muat semula data dan cuba lagi.');
    }
    cur.splice(row - 2, 1);
    writeBlock_(tx, cur, last);
    SpreadsheetApp.flush();
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------ pembantu
/** Baca A..J dan L sebagai senarai {v:[10 nilai], nota}. */
function readBlock_(tx, last) {
  if (last < 2) return [];
  const v = tx.getRange(2, 1, last - 1, NVAL).getValues();
  const n = tx.getRange(2, 12, last - 1, 1).getValues();
  return v.map(function (row, i) { return { v: row, nota: n[i][0] }; });
}

/** Tulis semula blok. prevLast = bilangan baris lama (untuk kosongkan baris lebihan selepas padam). */
function writeBlock_(tx, items, prevLast) {
  const n = items.length;
  if (n) {
    tx.getRange(2, 1, n, NVAL).setValues(items.map(function (it) { return it.v; }));
    tx.getRange(2, 12, n, 1).setValues(items.map(function (it) { return [it.nota]; }));
  }
  if (prevLast && prevLast - 1 > n) {
    const extra = prevLast - 1 - n;
    tx.getRange(n + 2, 1, extra, NVAL).clearContent();
    tx.getRange(n + 2, 12, extra, 1).clearContent();
  }
}

function toRow_(x, tz) {
  const kelab = num_(x.kelab), lawatan = num_(x.lawatan), lain = num_(x.lain), keluar = num_(x.keluar);
  if (!(kelab + lawatan + lain + keluar > 0)) throw new Error('Jumlah tidak sah');
  if (keluar > 0 && kelab + lawatan + lain > 0) throw new Error('Satu baris tidak boleh masuk dan keluar serentak');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(x.tarikh || '')) throw new Error('Tarikh tidak sah');
  const mula = x.mula || '', akhir = x.akhir || '';
  if ((mula && !/^\d{4}-\d{2}$/.test(mula)) || (akhir && !/^\d{4}-\d{2}$/.test(akhir))) throw new Error('Format bulan tidak sah');
  if (akhir && !mula) throw new Error('Bulan akhir perlu bulan mula');
  if (mula && akhir && akhir < mula) throw new Error('Bulan akhir sebelum bulan mula');
  const d = Utilities.parseDate(x.tarikh, tz, 'yyyy-MM-dd');
  return {
    v: [d, clean_(x.butiran), clean_(x.ahli), clean_(x.kategori),
        mula ? Utilities.parseDate(mula + '-01', tz, 'yyyy-MM-dd') : '',
        akhir && akhir !== mula ? Utilities.parseDate(akhir + '-01', tz, 'yyyy-MM-dd') : '',
        kelab || '', lawatan || '', lain || '', keluar || ''],
    nota: clean_(x.nota)
  };
}

function lastRow_(tx) {
  // baris terakhir yang ada tarikh di lajur A (formula Baki dll wujud sehingga baris 2001)
  const v = tx.getRange(1, 1, tx.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][0] !== '') return i + 1;
  return 1;
}

function fp_(tarikh, butiran, kelab, lawatan, lain, keluar) {
  return [tarikh, butiran, kelab, lawatan, lain, keluar].join('|');
}

function num_(v) {
  const n = Number(v);
  return isFinite(n) && n > 0 && n < 1000000 ? Math.round(n * 100) / 100 : 0;
}

function fmt_(v, tz, pattern) {
  return v instanceof Date ? Utilities.formatDate(v, tz, pattern) : '';
}

function clean_(s) {
  s = String(s == null ? '' : s).slice(0, 200);
  return /^[=+\-@]/.test(s) ? ' ' + s : s; // elak formula injection
}
