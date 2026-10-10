# Belanjawan Kelab Guru

Aplikasi telefon (Android) untuk rekod wang masuk/keluar kelab, yang menulis terus ke satu Google Sheet.

```
Telefon Android (Chrome)  ──►  Web App (Apps Script)  ──►  Google Sheet "Belanjawan Kelab Guru"
   borang + ringkasan          semak PIN, tulis baris        Transaksi · Ringkasan · Yuran Ahli · Tetapan
```

## Apa dalam repo ini

| Fail | Fungsi |
|---|---|
| `Belanjawan_Kelab_Guru_Baru.xlsx` | Helaian baharu: semua 282 transaksi 2025–2026 sudah dipindahkan, dengan Ringkasan, Yuran Ahli dan Tetapan |
| `apps-script/Code.gs`, `Index.html`, `appsscript.json` | Aplikasi (pelayan + paparan telefon) |
| `tools/migrate.py`, `tools/build_sheet.py` | Skrip yang menukar helaian lama kepada format baharu (boleh dijalankan semula) |

Baki selepas migrasi = **RM887.02**, sama dengan baki akhir helaian lama. Dec 2025 berakhir RM4,540 = "Baki daripada 2025" dalam helaian 2026.

## Pemasangan (sekali sahaja, ±10 minit)

1. **Import helaian.** Buka <https://sheets.new>, *File → Import → Upload* `Belanjawan_Kelab_Guru_Baru.xlsx` → *Replace spreadsheet*. Namakan "Belanjawan Kelab Guru".
2. **Pasang aplikasi.** Dalam helaian itu: *Extensions → Apps Script*.
   - Padam kandungan `Code.gs`, tampal isi `apps-script/Code.gs`.
   - Klik **+ → HTML**, namakan `Index`, tampal isi `apps-script/Index.html`.
   - *Project Settings* → tandakan "Show appsscript.json", tampal isi `apps-script/appsscript.json`.
3. **Tetapkan PIN.** Dalam editor, tambah baris sementara `function p(){ setPin('1234'); }` (tukar 1234), pilih `p` → **Run**, beri kebenaran. Kemudian padam fungsi `p`.
4. **Deploy.** *Deploy → New deployment → Web app*; *Execute as: Me*; *Who has access: Anyone*. Salin URL `https://script.google.com/macros/s/…/exec`.
   - Jika akaun sekolah (moe-dl.edu.my) menyekat "Anyone", minta pentadbir benarkan, atau buat helaian & skrip di akaun Gmail peribadi.
5. **Pasang di Android.** Buka URL dalam Chrome → menu ⋮ → **Add to Home screen** (Tambah ke skrin utama). Masukkan PIN sekali; telefon akan mengingatinya.

## Cara guna

- **+** → *Masuk → Yuran ahli*: pilih ahli, tandakan bulan, Kelab/Lawatan. Dua baris (Kelab & Lawatan) ditulis dengan butiran `Nama (3, 4)` seperti format lama.
- **+** → *Keluar* atau *Lain-lain*: pilih kategori, butiran, jumlah.
- **Rekod**: cari/tapis, ketik satu baris untuk padam.
- Edit terus di Google Sheet pun boleh; ketik *Muat semula* (tarik penyemak imbas) untuk lihat dalam app.

## Keselamatan

URL web app boleh dibuka sesiapa yang tahu pautannya, tetapi **semua data dilindungi PIN** pada pelayan (tanpa PIN tiada data dihantar). Jangan kongsi PIN dalam kumpulan awam. PIN 4–8 digit; tukar dengan `setPin()` lagi. Untuk keselamatan lebih tinggi, tukar "Who has access" kepada akaun Google tertentu (Workspace).

## Had & perkara untuk disemak

- Ini web app yang dipasang sebagai ikon (bukan fail `.apk`). Jika mahu `.apk`, ia boleh dibungkus (TWA) kemudian.
- Tab **Yuran Ahli** ialah *anggaran*: tarikh Mula/Tamat ahli di tab *Tetapan* diteka daripada data lama (cth. Wafi tamat Mac 2026, Suhaimi Ogos 2025). Semak dan betulkan. Kuchai & Wafi ditetapkan Lawatan = 0 kerana data lama menunjukkan pengecualian/refund.
- Tarikh helaian lama bercampur (sebahagian "bulan/hari" tertukar oleh Excel). Skrip migrasi membetulkannya dengan menyemak urutan baris; semak tab *Transaksi* untuk sebarang tarikh janggal.
- Kategori ditetapkan secara automatik daripada kata kunci dalam butiran lama — boleh diubah terus dalam helaian.
