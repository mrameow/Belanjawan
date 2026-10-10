# Belanjawan Kelab Guru

Aplikasi telefon (Android) untuk rekod wang masuk/keluar kelab, yang menulis terus ke Google Sheet anda.

```
Telefon Android (Chrome)  ──►  Web App (Apps Script)  ──►  Google Sheet anda
   borang + senarai yuran       semak PIN, tulis baris       Ringkasan · Harian · Transaksi · Yuran Ahli · Tetapan
```

## Fail

| Fail | Fungsi |
|---|---|
| `Belanjawan_Kelab_Guru_Baru.xlsx` | Tab baharu untuk dimasukkan ke Google Sheet anda (212 baris 2025–2026 sudah dipindahkan) |
| `apps-script/Code.gs`, `Index.html`, `appsscript.json` | Aplikasi |
| `tools/migrate.py`, `tools/build_sheet.py` | Skrip yang menukar helaian lama kepada format baharu |

Baki selepas migrasi = **RM887.02**, sama dengan helaian lama.

## Struktur data (penting)

Satu baris dalam **Transaksi** = satu bayaran, seperti helaian lama:

| Tarikh | Ahli | Butiran | Kategori | **Bulan Yuran** | Kelab | Lawatan | Masuk Lain | Keluar |
|---|---|---|---|---|---|---|---|---|
| 10 Okt 2026 | Amir | Amir (Mac 26, Apr 26) | Yuran | `2026-03, 2026-04` | 70 | 100 | | |

* Nombor dalam kurungan (cth. `Amir (3, 4)`) ialah **bulan yang dibayar**. Ia kini disimpan dalam lajur **Bulan Yuran** supaya sistem tahu bulan mana sudah/belum dibayar.
* Bayaran separa: `2025-12:30` (bulan 12 dibayar RM30). Dipaparkan sebagai ½.
* **Harian**: satu baris = satu hari (terbaru di atas) dengan jumlah masuk, keluar dan baki hujung hari. Dalam **Transaksi**, setiap hari berselang-seli warna dan tarikh hanya jelas pada baris pertama hari itu.
* **Yuran Ahli**: jadual ahli × bulan (✓ dibayar, ✗ belum, ½ separa) dan senarai bulan belum bayar bagi setiap ahli.

## Pemasangan (sekali sahaja)

### A. Masukkan tab baharu ke Google Sheet anda
1. Muat naik `Belanjawan_Kelab_Guru_Baru.xlsx` ke Google Drive (atau terus dari komputer).
2. Buka Google Sheet anda → **File → Import → Upload** → pilih fail → **Import location: _Insert new sheet(s)_** → Import data.
   Lima tab baharu muncul; tab lama (2025, 2026, dsb.) tidak disentuh.

### B. Pasang aplikasi
1. Dalam Google Sheet itu: **Extensions → Apps Script**.
2. Padam kandungan `Code.gs`, tampal isi `apps-script/Code.gs`.
3. Klik **+ → HTML**, namakan tepat `Index`, tampal isi `apps-script/Index.html`.
4. **Project Settings (⚙) → tandakan "Show appsscript.json manifest file"**, buka fail itu dan tampal isi `apps-script/appsscript.json`.
5. Dalam `Code.gs`, tukar `1234` dalam `aturPin()` kepada PIN anda. Pilih fungsi **aturPin** di bar atas → **Run** → **Review permissions → pilih akaun → Advanced → Go to … (unsafe) → Allow**.
   *Langkah ini wajib. Tanpa kebenaran, pautan web app tidak akan terbuka.*
6. (Pilihan) pilih **testApp** → Run. Log mesti menunjukkan `OK: 212 transaksi, 15 ahli`.
7. **Deploy → New deployment → ⚙ Web app** → *Execute as: Me*, *Who has access: Anyone* → Deploy. Salin URL yang berakhir `/exec`.
8. Di Android: buka URL dalam Chrome → menu ⋮ → **Add to Home screen**. Masukkan PIN sekali.

> Bila anda ubah kod kemudian, **Deploy → Manage deployments → ✏ → Version: New version → Deploy**. Jika tidak, URL lama masih menjalankan kod lama.

## Pautan tidak boleh dibuka?

### Jika yang keluar ialah "Sorry, unable to open the file at this time. Please check the address and try again." (halaman Google Drive)
Ini **bukan** ralat kod (ralat kod memaparkan halaman Apps Script). Ia bermaksud Google tidak boleh memadankan pautan dengan akaun yang sedang log masuk. Cuba mengikut turutan:

1. **Tetingkap Incognito/Private.** Buka pautan di situ. Jika ia terbuka (atau minta log masuk), puncanya ialah beberapa akaun Google log masuk serentak. Di telefon: log keluar akaun lain dalam Chrome, atau guna tab Incognito.
2. **Guna URL yang sebenar.** Dalam editor Apps Script pilih fungsi **diagnos** → Run → lihat *Execution log*. Baris `URL web app` ialah pautan yang betul. Pada akaun sekolah ia biasanya berbentuk `https://script.google.com/a/macros/moe-dl.edu.my/s/…/exec` (ada `/a/macros/<domain>/`). Pautan lama daripada deployment yang dipadam juga memberi mesej ini.
3. **Semak tetapan deployment.** *Deploy → Manage deployments → ✏*: *Execute as* = **Me**, *Who has access* = **Anyone**. Jika pilihan *Anyone* tiada atau pautan masih gagal, pentadbir domain sekolah mungkin menyekatnya.
4. **Jalan keluar paling pasti: guna akaun Gmail peribadi.** Muat naik `.xlsx` ke Drive Gmail itu, import ke Google Sheet baharu (Import → Upload), ulang langkah B1–B8. Akaun peribadi tiada sekatan pentadbir.

### Masalah lain
- **"Authorization is required"** → belum beri kebenaran; ulang langkah B5.
- **URL berakhir `/dev`** → mesti `/exec`.
- **Skrin kosong/ralat selepas ubah kod** → *Manage deployments → ✏ → Version: New version → Deploy*. Pastikan fail HTML dinamakan tepat `Index`.

## Cara guna

- **+ → Masuk → Yuran ahli**: pilih ahli, ketik bulan yang dibayar (bulan belum bayar bertanda merah, sudah bayar hijau), tandakan Kelab/Lawatan. Satu baris ditulis dengan bulan tersimpan.
- **+ → Keluar** atau **Lain-lain**: kategori, butiran, jumlah.
- **Rekod**: dikumpul mengikut hari; cari/tapis; ketik satu baris untuk padam (perlu ketik dua kali).
- **Yuran**: bulan belum bayar bagi setiap ahli.
- Boleh juga edit terus dalam Google Sheet. Selepas tambah baris baru secara manual, pastikan lajur *Tarikh* diisi dan baris disusun ikut tarikh (Data → Sort) supaya pengasingan harian betul. App menyusun automatik.

## Keselamatan

URL web app boleh dibuka sesiapa yang tahu pautannya, tetapi **semua data dilindungi PIN** pada pelayar dan pelayan. Jangan kongsi PIN atau URL dalam kumpulan awam. Tukar PIN dengan mengubah `aturPin()` dan Run semula.

## Had & perkara untuk disemak

- Ini web app yang dipasang sebagai ikon, bukan fail `.apk`.
- **Tarikh Mula/Tamat ahli** (tab Tetapan) diteka daripada data lama (cth. Wafi tamat Mac 2026, Suhaimi Ogos 2025, Boss Khairil Sep 2025). Semak, kerana ia menentukan bulan mana dikira "belum bayar".
- **3 bayaran lama tiada bulan ditulis** ("Ustaz" lawatan RM100, "Hasnol" RM225, "Kuchai" RM315). Bulan diisi automatik pada bulan kosong terawal; semak lajur *Bulan Yuran* untuk baris ini.
- Kuchai dan Wafi ditetapkan Lawatan = 0 (kadar 0 = dikecualikan) berdasarkan data lama; ubah di Tetapan jika salah.
- Tarikh helaian lama bercampur (sebahagian bulan/hari tertukar oleh Excel); skrip migrasi membetulkannya. Semak tab Harian untuk tarikh janggal.
- Kategori kos diteka daripada perkataan dalam butiran lama; boleh diubah dalam helaian.
