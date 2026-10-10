/**
 * Belanjawan Kelab Guru - bahagian pelayan (Google Apps Script).
 * Skrip ini "terikat" pada Google Sheet: Extensions > Apps Script. Hanya API; app ada di folder docs/ (GitHub Pages).
 *
 * Lajur helaian Transaksi:
 *  A ID | B Tarikh | C Ahli/Pihak | D Butiran | E Kategori | F Bulan Yuran |
 *  G Kelab | H Lawatan | I Masuk Lain | J Keluar | (K Baki, L Hari# = formula, jangan disentuh)
 */
const SHEET_TX = 'Transaksi';
const SHEET_SET = 'Tetapan';
const TZ = 'Asia/Kuala_Lumpur';
const NCOL = 10;

/**
 * API JSON untuk app (GitHub Pages / PWA). Semua permintaan ialah POST dengan badan teks JSON
 * {action, pin, ...}. Jenis kandungan text/plain mengelakkan preflight CORS.
 */
function doGet() {
  return json_({ ok: true, data: 'Belanjawan Kelab Guru API aktif. Gunakan app untuk membaca data.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case 'getData': return json_({ ok: true, data: getData(req.pin) });
      case 'add': return json_({ ok: true, data: addTransactions(req.pin, req.rows) });
      case 'delete': return json_({ ok: true, data: deleteTransaction(req.pin, req.id) });
      default: throw new Error('Tindakan tidak dikenali');
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * LANGKAH PERTAMA: tukar 1234 di bawah kepada PIN pilihan anda (4-8 digit),
 * pilih fungsi "aturPin" di bar atas editor, tekan Run, dan beri kebenaran (Authorize).
 * Kebenaran ini WAJIB sebelum Deploy; tanpanya pautan web app tidak akan terbuka.
 */
function aturPin() { setPin('1234'); }

function setPin(pin) {
  pin = String(pin);
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN mesti 4-8 digit');
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Skrip ini mesti dibuka melalui Extensions > Apps Script dalam Google Sheet anda.');
  if (!ss.getSheetByName(SHEET_TX) || !ss.getSheetByName(SHEET_SET)) {
    throw new Error('Tab "Transaksi" / "Tetapan" tiada. Import helaian baharu dahulu.');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('PIN', pin);
  props.setProperty('SSID', ss.getId());
  LockService.getScriptLock(); // sentuh LockService supaya kebenarannya diminta sekarang
  Logger.log('Siap. PIN ditetapkan. Seterusnya: Deploy > New deployment > Web app.');
}

/**
 * Diagnosis pautan: pilih "diagnos" -> Run -> buka View > Logs (Execution log).
 * Mencetak URL web app yang SEBENAR (di akaun sekolah ia berbentuk /a/macros/<domain>/s/.../exec),
 * akaun pemilik, dan sama ada helaian boleh dibaca.
 */
function diagnos() {
  let url = '(belum di-deploy sebagai Web app)';
  try { url = ScriptApp.getService().getUrl() || url; } catch (e) { url = 'ralat: ' + e; }
  Logger.log('URL web app : ' + url);
  Logger.log('Akaun pemilik: ' + Session.getEffectiveUser().getEmail());
  Logger.log('PIN ditetapkan: ' + !!PropertiesService.getScriptProperties().getProperty('PIN'));
  Logger.log('Helaian      : ' + ss_().getName() + ' | Transaksi=' + !!ss_().getSheetByName(SHEET_TX) +
    ' | Tetapan=' + !!ss_().getSheetByName(SHEET_SET));
}

/** Ujian dari editor: pilih testApp dan tekan Run; log menunjukkan sama ada data boleh dibaca. */
function testApp() {
  const d = getData(PropertiesService.getScriptProperties().getProperty('PIN'));
  Logger.log('OK: ' + d.items.length + ' transaksi, ' + d.ahli.length + ' ahli');
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
  const id = PropertiesService.getScriptProperties().getProperty('SSID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Tab "' + name + '" tiada');
  return s;
}

/** Semua data untuk aplikasi. */
function getData(pin) {
  checkPin_(pin);
  const tx = sheet_(SHEET_TX);
  const last = lastRow_(tx);
  const rows = last < 2 ? [] : tx.getRange(2, 1, last - 1, NCOL).getValues();
  const items = rows.map(function (r) {
    return {
      id: r[0], tarikh: fmtDate_(r[1]), ahli: r[2], butiran: r[3], kategori: r[4], bulan: String(r[5] || ''),
      kelab: Number(r[6]) || 0, lawatan: Number(r[7]) || 0, lain: Number(r[8]) || 0, keluar: Number(r[9]) || 0
    };
  });

  const set = sheet_(SHEET_SET);
  const ahli = set.getRange(5, 1, 15, 6).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { nama: r[0], mula: fmtMonth_(r[1]), tamat: fmtMonth_(r[2]), kelab: Number(r[3]) || 0,
               lawatan: Number(r[4]) || 0, aktif: r[5] !== 'Tidak' };
    });
  const kat = set.getRange(5, 9, 11, 1).getValues().map(function (r) { return String(r[0]); });
  return {
    items: items, ahli: ahli,
    lawatanAkhir: fmtMonth_(set.getRange(21, 2).getValue()),
    kategori: { Masuk: kat.slice(1, 4), Keluar: kat.slice(4) }
  };
}

/**
 * Tambah transaksi. rows: [{tarikh:'YYYY-MM-DD', ahli, butiran, kategori, bulan, kelab, lawatan, lain, keluar}]
 */
function addTransactions(pin, rows) {
  checkPin_(pin);
  if (!rows || !rows.length || rows.length > 50) throw new Error('Data tidak sah');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const tx = sheet_(SHEET_TX);
    let n = lastRow_(tx);
    let seq = maxId_(tx, n);
    const ids = [];
    rows.forEach(function (x) {
      const kelab = num_(x.kelab), lawatan = num_(x.lawatan), lain = num_(x.lain), keluar = num_(x.keluar);
      if (!(kelab + lawatan + lain + keluar > 0)) throw new Error('Jumlah tidak sah');
      if (keluar > 0 && kelab + lawatan + lain > 0) throw new Error('Satu baris tidak boleh masuk dan keluar serentak');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(x.tarikh)) throw new Error('Tarikh tidak sah');
      if (x.bulan && !/^\d{4}-\d{2}(:\d+(\.\d+)?)?(, \d{4}-\d{2}(:\d+(\.\d+)?)?)*$/.test(x.bulan)) throw new Error('Format bulan tidak sah');
      const id = 'M' + ('0000' + (++seq)).slice(-4);
      n++;
      const d = x.tarikh.split('-');
      tx.getRange(n, 1, 1, NCOL).setValues([[
        id, new Date(+d[0], +d[1] - 1, +d[2]), clean_(x.ahli), clean_(x.butiran), clean_(x.kategori),
        x.bulan || '', kelab || '', lawatan || '', lain || '', keluar || ''
      ]]);
      ids.push(id);
    });
    // susun ikut tarikh (kemudian ID) supaya pengasingan harian & baki sentiasa betul
    tx.getRange(2, 1, n - 1, NCOL).sort([{ column: 2, ascending: true }, { column: 1, ascending: true }]);
    SpreadsheetApp.flush();
    return ids;
  } finally {
    lock.releaseLock();
  }
}

/** Padam transaksi mengikut ID. */
function deleteTransaction(pin, id) {
  checkPin_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const tx = sheet_(SHEET_TX);
    const last = lastRow_(tx);
    const ids = tx.getRange(2, 1, Math.max(last - 1, 1), 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === id) {
        const start = i + 2;
        const rest = start < last ? tx.getRange(start + 1, 1, last - start, NCOL).getValues() : [];
        tx.getRange(start, 1, last - start + 1, NCOL).clearContent(); // kekalkan formula & format
        if (rest.length) tx.getRange(start, 1, rest.length, NCOL).setValues(rest);
        return true;
      }
    }
    throw new Error('Transaksi tidak dijumpai');
  } finally {
    lock.releaseLock();
  }
}

// ---- pembantu ----
function lastRow_(tx) {
  // baris terakhir yang ada ID di lajur A (formula Baki/Hari# wujud sehingga baris 3000)
  const v = tx.getRange(1, 1, tx.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '') return i + 1;
  return 1;
}

function maxId_(tx, last) {
  // ID tertinggi (baris disusun ikut tarikh, jadi baris terakhir bukan semestinya ID terbesar)
  if (last < 2) return 0;
  return tx.getRange(2, 1, last - 1, 1).getValues().reduce(function (m, r) {
    return Math.max(m, parseInt(String(r[0]).replace(/\D/g, ''), 10) || 0);
  }, 0);
}

function num_(v) {
  const n = Number(v);
  return isFinite(n) && n > 0 && n < 1000000 ? Math.round(n * 100) / 100 : 0;
}

function fmtDate_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v);
}

function fmtMonth_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM') : '';
}

function clean_(s) {
  s = String(s == null ? '' : s).slice(0, 200);
  return /^[=+\-@]/.test(s) ? ' ' + s : s; // elak formula injection
}
