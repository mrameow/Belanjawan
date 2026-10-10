# Belanjawan Kelab Guru

> **Status:** `Belanjawan_Kelab_Guru_Baru.xlsx` telah direka semula (Dashboard, Harian, Yuran, Transaksi, Tetapan, Panduan) dan ialah fokus semasa.
> Folder `docs/` dan `apps-script/` di bawah masih untuk struktur lajur **lama**; ia akan disesuaikan dengan Excel baharu selepas Excel diluluskan.

PWA (app yang boleh dipasang di Android) untuk rekod wang masuk/keluar kelab, menulis terus ke Google Sheet anda.

```
Android / mana-mana pelayar ──► docs/ (GitHub Pages, PWA) ──► Apps Script (API JSON) ──► Google Sheet anda
        buka & pasang app          borang + senarai yuran         semak PIN, baca/tulis        Ringkasan · Harian · Transaksi · Yuran Ahli · Tetapan
```

## Fail

| Fail | Fungsi |
|---|---|
| `docs/` | **App (PWA)**: `index.html`, `manifest.webmanifest`, `sw.js`, `icons/`. Dihidang oleh GitHub Pages |
| `apps-script/Code.gs` | API kecil yang ditampal dalam Google Sheet (tiada paparan HTML lagi) |
| `Belanjawan_Kelab_Guru_Baru.xlsx` | Tab baharu untuk dimasukkan ke Google Sheet anda |
| `tools/` | Skrip yang menukar helaian lama kepada format baharu |

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
1. Buka Google Sheet anda → **File → Import → Upload** → pilih `Belanjawan_Kelab_Guru_Baru.xlsx` → **Import location: _Insert new sheet(s)_** → Import data. Lima tab baharu muncul; tab lama tidak disentuh.

### B. Pasang API (Apps Script)
1. Dalam Google Sheet itu: **Extensions → Apps Script**. Padam kandungan `Code.gs`, tampal isi `apps-script/Code.gs`. (Tiada fail HTML diperlukan.)
2. Tukar `1234` dalam `aturPin()` kepada PIN anda. Pilih fungsi **aturPin** di bar atas → **Run** → **Review permissions → akaun → Advanced → Go to … → Allow**. *Wajib.*
3. (Pilihan) Run **testApp**: log mesti `OK: … transaksi, 15 ahli`.
4. **Deploy → New deployment → ⚙ Web app** → *Execute as: Me*, *Who has access: Anyone* → Deploy. Salin URL yang berakhir `/exec`.
   Bila `Code.gs` diubah kemudian: **Manage deployments → ✏ → Version: New version → Deploy**.

### C. Hidangkan app di GitHub Pages
1. Repo GitHub → **Settings → Pages → Build and deployment → Source: Deploy from a branch** → Branch: cawangan yang ada folder `docs` (`claude/fervent-einstein-y0czm1`), folder **/docs** → Save.
2. Tunggu 1–2 minit; alamat muncul di bahagian atas halaman itu, biasanya `https://mrameow.github.io/Belanjawan/`.
3. **Penting:** GitHub Pages untuk repo *private* memerlukan pelan berbayar (Pro/Team). Pada pelan percuma repo mesti *public*. Jangan jadikan repo ini public kerana ia mengandungi fail `.xlsx` dengan nama dan wang kelab. Lihat "Tanpa GitHub Pages" di bawah.

### D. Pasang di Android
1. Buka alamat Pages dalam **Chrome**. Pada kali pertama, tampal **URL API** (dari B4) dan **PIN**, tekan *Sambung*.
2. Menu ⋮ Chrome → **Install app** / **Add to Home screen** (atau ⚙ dalam app → *Pasang app*). Ikon "RM" muncul di skrin utama dan app dibuka tanpa bar pelayar.
3. URL API dan PIN disimpan dalam peranti itu sahaja. Boleh buat yang sama di telefon/komputer lain.

### Tanpa GitHub Pages
Folder `docs/` ialah laman statik biasa. Muat naik ke mana-mana hos statik HTTPS: **Cloudflare Pages**, **Netlify Drop** (seret folder `docs`), atau repo GitHub *public* baharu yang hanya mengandungi `docs/`. Tiada data dalam folder itu; data hanya dibaca daripada Google Sheet dengan PIN.

## Masalah

- **"Jawapan tidak sah daripada API"** → URL API salah, atau belum Deploy sebagai Web app. Mesti berakhir `/exec`.
- **"PIN salah"** → mesti sama dengan `aturPin()`. Lima cubaan salah mengunci 10 minit.
- **Luar talian** → app memaparkan data terakhir yang disimpan; simpan & padam perlu internet.
- **Mengubah `Code.gs`** → buat *New version* (B4), jika tidak URL lama masih menjalankan kod lama.
- **Membuka URL API terus dalam Chrome menunjukkan "Sorry, unable to open the file"** → kesan beberapa akaun Google log masuk; *tidak menjejaskan app* kerana app memanggil API tanpa log masuk. Untuk menguji URL, buka dalam Incognito: ia memaparkan `{"ok":true,…}`.
- **Pilihan *Anyone* tiada** (akaun sekolah) → minta pentadbir domain benarkan, atau guna akaun Gmail peribadi.

## Cara guna

- **+ → Masuk → Yuran ahli**: pilih ahli, ketik bulan yang dibayar (bulan belum bayar bertanda merah, sudah bayar hijau), tandakan Kelab/Lawatan. Satu baris ditulis dengan bulan tersimpan.
- **+ → Keluar** atau **Lain-lain**: kategori, butiran, jumlah.
- **Rekod**: dikumpul mengikut hari; cari/tapis; ketik satu baris untuk padam (perlu ketik dua kali).
- **Yuran**: bulan belum bayar bagi setiap ahli.
- Boleh juga edit terus dalam Google Sheet. Selepas tambah baris baru secara manual, pastikan lajur *Tarikh* diisi dan baris disusun ikut tarikh (Data → Sort) supaya pengasingan harian betul. App menyusun automatik.

## Keselamatan

URL API boleh dipanggil sesiapa yang tahu pautannya, tetapi **semua data dilindungi PIN** pada pelayan (lima cubaan salah = kunci 10 minit). Jangan kongsi URL API atau PIN dalam kumpulan awam. Tukar PIN dengan mengubah `aturPin()` dan Run semula. App di GitHub Pages tidak mengandungi data atau PIN.

## Had & perkara untuk disemak

- PWA dipasang melalui Chrome (bukan fail `.apk` dalam Play Store); ia berkelakuan seperti app biasa.
- **Tarikh Mula/Tamat ahli** (tab Tetapan) diteka daripada data lama (cth. Wafi tamat Mac 2026, Suhaimi Ogos 2025, Boss Khairil Sep 2025). Semak, kerana ia menentukan bulan mana dikira "belum bayar".
- **3 bayaran lama tiada bulan ditulis** ("Ustaz" lawatan RM100, "Hasnol" RM225, "Kuchai" RM315). Bulan diisi automatik pada bulan kosong terawal; semak lajur *Bulan Yuran* untuk baris ini.
- Kuchai dan Wafi ditetapkan Lawatan = 0 (kadar 0 = dikecualikan) berdasarkan data lama; ubah di Tetapan jika salah.
- Tarikh helaian lama bercampur (sebahagian bulan/hari tertukar oleh Excel); skrip migrasi membetulkannya. Semak tab Harian untuk tarikh janggal.
- Kategori kos diteka daripada perkataan dalam butiran lama; boleh diubah dalam helaian.
