# Belanjawan Kelab Guru

PWA (app yang boleh dipasang di Android) untuk rekod wang masuk/keluar kelab. Ia menulis terus ke Google Sheet anda dan menunjukkan siapa yang belum bayar yuran.

```
Android (Chrome)  ──►  GitHub Pages (docs/, PWA)  ──►  Apps Script (apps-script/Code.gs)  ──►  Google Sheet anda
 buka / pasang app       borang, rekod, Yuran           semak PIN, baca & tulis baris              Dashboard · Harian · Yuran · Transaksi · Tetapan
```

| Fail | Fungsi |
|---|---|
| `docs/` | App (PWA): `index.html`, `manifest.webmanifest`, `sw.js`, `icons/` |
| `apps-script/Code.gs` | API kecil yang ditampal dalam Google Sheet (Extensions → Apps Script) |
| `.github/workflows/pages.yml` | Menerbit `docs/` ke GitHub Pages secara automatik |
| `Belanjawan_Kelab_Guru_Baru.xlsx`, `tools/` | Excel baharu dan skrip yang menjana/memindahkan data |

## Cara app berfungsi
- **Rekod baru** ditulis sebagai baris dalam tab **Transaksi** (lajur Tarikh, Butiran, Ahli, Kategori, Bulan mula/akhir, Kelab, Lawatan, Masuk lain, Keluar). Baris disusun ikut tarikh supaya tab **Harian** betul. Formula Baki dan lajur pembantu tidak disentuh.
- **Yuran belum bayar**: app membaca terus tab **Yuran** dalam sheet (✓ ✗ ½), jadi apa yang anda nampak dalam app sentiasa sama dengan sheet. Ketik nama ahli untuk lihat 19 bulan, atau *Catat bayaran* untuk menandakan bulan yang dipilih.
- Pilih beberapa bulan berturutan dalam borang yuran → satu baris (Bulan mula–akhir). Bulan tidak berturutan → baris berasingan, sama seperti petua dalam tab Panduan.
- **Padam**: ketik rekod → padam (dua ketukan). Pelayan menyemak baris masih sama sebelum memadam.
- **Luar talian**: app dan data terakhir masih boleh dilihat; simpan dan padam perlu internet.

## Pemasangan (sekali sahaja)

### 1. Pasang API dalam Google Sheet anda
1. Buka sheet anda → **Extensions → Apps Script**. Padam kandungan `Code.gs`, tampal isi `apps-script/Code.gs`. (ID sheet sudah ada dalam kod; jika anda guna sheet lain, tukar `DEFAULT_SSID`.)
2. Tukar `1234` dalam fungsi `aturPin()` kepada PIN anda (4–8 digit). Pilih fungsi **aturPin** di bar atas → **Run** → **Review permissions** → akaun → **Advanced → Go to … → Allow**. *Wajib, jika tidak API tidak akan berfungsi.*
3. Pilih fungsi **diagnos** → Run → buka **Execution log**: mesti tertera `OK: … transaksi, … ahli, … baris yuran`.
4. **Deploy → New deployment → ⚙ Web app** → *Execute as: Me*, *Who has access: Anyone* → Deploy. Salin URL yang berakhir `/exec`.
   Bila `Code.gs` diubah kemudian: **Manage deployments → ✏ → Version: New version → Deploy**.

### 2. Terbitkan app di GitHub Pages
1. Repo GitHub → **Settings → Pages → Build and deployment → Source: _GitHub Actions_**.
2. Push apa-apa ke `docs/` (atau Actions → *Terbit PWA ke GitHub Pages* → *Run workflow*). Alamat app muncul dalam Settings → Pages, biasanya `https://mrameow.github.io/Belanjawan/`.
3. **Penting:** GitHub Pages untuk repo *private* memerlukan pelan berbayar (Pro/Team). Pada pelan percuma repo mesti *public*. **Jangan jadikan repo ini public**: ia mengandungi fail `.xlsx` dengan nama dan wang kelab (juga dalam sejarah commit). Alternatif: muat naik folder `docs/` ke hos statik HTTPS lain (Cloudflare Pages, Netlify Drop) atau repo public baharu yang hanya mengandungi `docs/`. Folder itu tiada data.

### 3. Pasang di Android
1. Buka alamat app dalam **Chrome**. Kali pertama: tampal **URL API** (langkah 1.4) dan **PIN**, tekan *Sambung*.
2. Menu ⋮ → **Install app** (atau ⚙ dalam app → *Pasang app*). Ikon "RM" muncul di skrin utama dan app dibuka tanpa bar pelayar.
3. URL API dan PIN disimpan dalam peranti itu sahaja. Ulang pada peranti lain.

## Masalah
- **"Jawapan tidak sah daripada API"** → URL salah atau belum Deploy sebagai Web app. Mesti berakhir `/exec`.
- **"PIN salah"** → mesti sama dengan `aturPin()`. Lima cubaan salah mengunci 10 minit.
- **"Susun atur tab Yuran berubah"** → jangan alih/padam baris dalam tab Yuran. `A27` mesti "Kelab", `A51` mesti "Lawatan".
- **"Tab … tiada"** → nama tab Transaksi, Tetapan, Yuran mesti tepat.
- **Selepas mengubah `Code.gs`** → buat *New version* (1.4).
- **Membuka URL API terus dalam Chrome menunjukkan "Sorry, unable to open the file"** → kesan beberapa akaun Google log masuk; *tidak menjejaskan app*, kerana app memanggil API tanpa log masuk. Untuk menguji URL, buka dalam Incognito: ia memaparkan `{"ok":true,…}`.
- **Pilihan *Anyone* tiada** (akaun sekolah) → minta pentadbir domain benarkan, atau guna akaun Gmail peribadi.
- **App tidak dikemas kini selepas push** → tutup dan buka semula app; versi baharu diambil bila ada talian.

## Keselamatan
URL API boleh dipanggil sesiapa yang tahu pautannya, tetapi **semua data dilindungi PIN** pada pelayan (lima cubaan salah = kunci 10 minit). Jangan kongsi URL API atau PIN dalam kumpulan awam. Tukar PIN dengan mengubah `aturPin()` dan Run semula. App di GitHub Pages tidak mengandungi data atau PIN.

## Had
- Bayaran yuran separa: catat jumlah kurang melalui sheet (Transaksi) kerana app menganggap kadar penuh setiap bulan.
- Sehingga 2,000 baris Transaksi (cukup untuk beberapa tahun).
- Anggaran "belum bayar (RM)" mengira bulan ✗ × kadar; bulan ½ tidak dikira.
- Tarikh Mula/Tamat ahli (tab Tetapan) menentukan bulan mana dikira belum bayar; semak dan betulkan di situ.
