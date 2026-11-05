# GIT Asset Online (Aktivasi & Mutasi)
Aplikasi web statis untuk memantau aktivasi dan mutasi asset online yang belum selesai.
Data dibaca dari file `REPORT_GIT.xlsx` yang diunggah di halaman (diproses di perangkat, tidak dikirim ke server).

## Struktur
| File | Isi |
|---|---|
| `index.html` | Kerangka halaman dan header |
| `css/style.css` | Tampilan (responsif, mode gelap) |
| `js/app.js` | Seluruh logika: baca Excel, filter bulan, cari toko, histori, share WA |
| `vendor/xlsx.full.min.js` | Library pembaca Excel (SheetJS 0.18.5), jangan diedit |
| `build.js` | Menggabungkan semuanya menjadi satu file mandiri |

## Cara pakai di VS Code
1. **File > Open Folder**, pilih folder ini. Pasang ekstensi yang disarankan (Live Server, Prettier).
2. Klik kanan `index.html` > **Open with Live Server**. Halaman terbuka di browser dan otomatis refresh setiap file disimpan.
3. Unggah `REPORT_GIT.xlsx` di halaman untuk menguji perubahan. Jangan simpan file Excel di folder ini (sudah diabaikan oleh `.gitignore`).
4. Selesai mengedit: **Terminal > Run Build Task** (Ctrl+Shift+B), atau `node build.js`.
   Hasil: `dist/GIT_Asset_Online.html`, satu file yang bisa dikirim lewat WhatsApp atau dibuka tanpa internet.

## Bagian yang paling sering diubah (di `js/app.js`)
- Status yang dianggap butuh aksi toko: konstanta `TK` (kata kunci `pejabat toko`, `user penerima`, ...).
- Pesan WhatsApp: fungsi `msg` di tab Share.
- Batas KRITIS / TINGGI: kolom isian di halaman (default 60 dan 30 hari), atau nilai awal di `index.html`.
- Sheet Excel yang dibaca: `UPDATE AKTIVASI`, `UPDATE MUTASI`, `UPDATE STRUKTUR`, `PARAMETER`, `PEJABAT TOKO`.

## Git (opsional)
`git init` > `git add .` > `git commit -m "versi awal"`. Setiap perubahan bisa dicatat dan dikembalikan.

## Publikasi
Lihat `PANDUAN_GITHUB_PAGES.md` untuk link github.io dan domain gratis.

## Jaringan lokal
`node server.js` menyajikan aplikasi ke perangkat lain di Wi-Fi/LAN yang sama. Lihat `PANDUAN_GITHUB_PAGES.md` bagian E.
