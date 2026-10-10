/**
 * Belanjawan Kelab Guru - bahagian pelayan (Google Apps Script).
 * Skrip ini "terikat" pada Google Sheet: Extensions > Apps Script.
 *
 * Lajur helaian Transaksi: A ID | B Tarikh | C Jenis | D Kategori | E Ahli/Pihak | F Butiran | G Jumlah | H Baki
 */
const SHEET_TX = 'Transaksi';
const SHEET_SET = 'Tetapan';
const TZ = 'Asia/Kuala_Lumpur';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Belanjawan Kelab Guru')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Jalankan SEKALI dari editor untuk menetapkan PIN: setPin('1234'). */
function setPin(pin) {
  if (!/^\d{4,8}$/.test(String(pin))) throw new Error('PIN mesti 4-8 digit');
  PropertiesService.getScriptProperties().setProperty('PIN', String(pin));
}

function checkPin_(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!saved) throw new Error('PIN belum ditetapkan. Jalankan setPin("1234") di editor Apps Script.');
  if (String(pin) !== saved) throw new Error('PIN salah');
}

function sheet_(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s) throw new Error('Helaian "' + name + '" tiada');
  return s;
}

/** Semua data untuk aplikasi. */
function getData(pin) {
  checkPin_(pin);
  const tx = sheet_(SHEET_TX);
  const last = lastRow_(tx);
  const rows = last < 2 ? [] : tx.getRange(2, 1, last - 1, 7).getValues();
  const items = rows.map(function (r) {
    return {
      id: r[0], tarikh: fmtDate_(r[1]), jenis: r[2], kategori: r[3],
      ahli: r[4], butiran: r[5], jumlah: Number(r[6]) || 0
    };
  });

  const set = sheet_(SHEET_SET);
  const ahli = set.getRange(5, 1, 15, 6).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { nama: r[0], kelab: Number(r[3]) || 0, lawatan: Number(r[4]) || 0, aktif: r[5] !== 'Tidak' }; });
  const masuk = set.getRange(5, 9, 7, 1).getValues().map(function (r) { return r[0]; }).filter(String);
  const keluar = set.getRange(5, 10, 7, 1).getValues().map(function (r) { return r[0]; }).filter(String);
  return { items: items, ahli: ahli, kategori: { Masuk: masuk, Keluar: keluar } };
}

/** Tambah satu atau lebih transaksi. rows: [{tarikh:'YYYY-MM-DD', jenis, kategori, ahli, butiran, jumlah}] */
function addTransactions(pin, rows) {
  checkPin_(pin);
  if (!rows || !rows.length || rows.length > 50) throw new Error('Data tidak sah');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const tx = sheet_(SHEET_TX);
    let n = lastRow_(tx);
    const out = [];
    rows.forEach(function (x) {
      const jumlah = Number(x.jumlah);
      if (!(jumlah > 0) || jumlah > 1000000) throw new Error('Jumlah tidak sah');
      if (x.jenis !== 'Masuk' && x.jenis !== 'Keluar') throw new Error('Jenis tidak sah');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(x.tarikh)) throw new Error('Tarikh tidak sah');
      n++;
      const id = nextId_(tx, n);
      const d = x.tarikh.split('-');
      tx.getRange(n, 1, 1, 7).setValues([[
        id, new Date(+d[0], +d[1] - 1, +d[2]), x.jenis, clean_(x.kategori), clean_(x.ahli), clean_(x.butiran), jumlah
      ]]);
      out.push(id);
    });
    SpreadsheetApp.flush();
    return out;
  } finally {
    lock.releaseLock();
  }
}

/** Padam transaksi mengikut ID (baris dibuang, baki dikira semula oleh formula). */
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
        const rest = start < last ? tx.getRange(start + 1, 1, last - start, 7).getValues() : [];
        tx.getRange(start, 1, last - start + 1, 7).clearContent(); // kekalkan formula Baki & format
        if (rest.length) tx.getRange(start, 1, rest.length, 7).setValues(rest);
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
  // baris terakhir yang ada ID di lajur A (helaian ada formula Baki sehingga baris 3000)
  const v = tx.getRange(1, 1, tx.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '') return i + 1;
  return 1;
}

function nextId_(tx, row) {
  const prev = row > 2 ? String(tx.getRange(row - 1, 1).getValue()) : 'M0000';
  const num = parseInt(prev.replace(/\D/g, ''), 10) || row - 1;
  return 'M' + ('0000' + (num + 1)).slice(-4);
}

function fmtDate_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v);
}

function clean_(s) {
  s = String(s == null ? '' : s).slice(0, 200);
  return /^[=+\-@]/.test(s) ? ' ' + s : s; // elak formula injection
}
