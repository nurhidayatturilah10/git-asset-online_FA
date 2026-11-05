# Publikasi ke GitHub Pages (gratis) + domain gratis

## A. Buat link di github.io
1. Buat akun di github.com (gratis), lalu **New repository**. Nama contoh: `git-asset-online`. Pilih **Public** (GitHub Pages gratis untuk repositori publik).
2. Klik **uploading an existing file**, seret SEMUA isi folder ini (kecuali folder `dist` bila ada), lalu **Commit changes**.
   Pastikan `index.html`, `sw.js`, `manifest.webmanifest`, folder `css`, `js`, `vendor`, `icons` berada di tingkat paling atas repositori.
3. Buka **Settings > Pages**. Di *Build and deployment* pilih **Deploy from a branch**, branch `main`, folder `/ (root)`, klik **Save**.
4. Tunggu 1 sampai 5 menit. Link Anda: `https://NAMAPENGGUNA.github.io/git-asset-online/`
5. Bagikan link itu. Di HP, buka link lalu pilih **Tambahkan ke layar utama** (Chrome) atau **Bagikan > Tambah ke Layar Utama** (Safari). Setelah dibuka sekali, aplikasi bisa jalan tanpa internet.

JANGAN unggah file Excel data ke repositori (repositori publik bisa dilihat siapa saja). File `.gitignore` sudah memblokir `*.xlsx` bila memakai Git.

## B. Domain gratis (opsional)
Link github.io sudah gratis dan permanen. Bila ingin nama lebih pendek, pilihan gratis:
| Pilihan | Contoh | Cara | Catatan |
|---|---|---|---|
| **is-a.dev** | `gitpalu.is-a.dev` | Fork repo `is-a-dev/register`, tambah `domains/gitpalu.json`, kirim pull request | Untuk situs developer. Disetujui manual, bisa ditolak. Efek sampai 1 hari |
| **js.org** | `gitpalu.js.org` | Pull request ke repo `js-org/js.org` | Untuk proyek JavaScript, ada syarat isi proyek |
| **eu.org** | `gitpalu.eu.org` | Daftar di nic.eu.org | Persetujuan bisa berminggu-minggu |

Isi berkas JSON untuk is-a.dev (ganti `NAMAPENGGUNA` dan `email`):
```json
{
  "owner": { "username": "NAMAPENGGUNA", "email": "email@anda.com" },
  "records": { "CNAME": "NAMAPENGGUNA.github.io" }
}
```
Setelah domain disetujui: **Settings > Pages > Custom domain**, isi `gitpalu.is-a.dev`, Save, lalu centang **Enforce HTTPS**.
Syarat dan format bisa berubah, selalu ikuti petunjuk terbaru di repo penyedia domain.

## C. Memperbarui aplikasi
Ubah file, unggah ulang ke repositori (commit). Naikkan angka `VERSI` di `sw.js` (mis. `v2`) agar semua perangkat memuat versi terbaru.

## D. Link untuk anak toko (hanya lihat + unduh Excel)
Halaman yang sama berubah otomatis menjadi **mode lihat** bila ada file `data.js` di repositori.
1. Buka halaman admin dengan menambahkan `?admin` di akhir link (contoh `https://NAMAPENGGUNA.github.io/git-asset-online/?admin`), unggah `REPORT_GIT.xlsx`.
2. Buka tab **Link Anak Toko**, isi kata sandi (minimal 10 karakter, dua kali), klik **Buat data.js**, simpan file.
3. Unggah `data.js` ke repositori (timpa file lama) lalu Commit. Link biasa (tanpa `?admin`) kini meminta kata sandi.
4. Bagikan link ke anak toko, dan kata sandi lewat jalur terpisah (pesan pribadi).
5. Setiap update data: ulangi langkah 1 sampai 3. Link tidak berubah.

Anak toko bisa: mencari toko, melihat AC, AM, Pejabat Toko, item belum selesai dan histori, serta **mengunduh Excel** (1 file, sheet Aktivasi dan Mutasi, hanya item belum selesai sesuai bulan dan toko yang tampil). Mereka tidak bisa mengubah apa pun dan tidak melihat tab admin.

Catatan keamanan:
- Data di `data.js` dienkripsi (AES-256-GCM, kunci dari kata sandi). Siapa pun yang punya link dapat mengunduh file terenkripsi itu, jadi **kekuatan sandi menentukan keamanan**. Pakai kalimat sandi panjang dan jangan dipakai di tempat lain.
- Riwayat Git menyimpan versi `data.js` lama. Bila sandi bocor, ganti sandi dan buat `data.js` baru; versi lama tetap terbuka dengan sandi lama.
- Siapa pun yang tahu sandi bisa melihat seluruh toko dan mengunduh Excel. Untuk pembatasan per toko diperlukan sistem login terpisah.

## E. Jaringan lokal (Wi-Fi kantor / satu LAN)
1. Di komputer yang menyimpan folder ini jalankan `node server.js` (atau VS Code: Terminal > Run Task > "Jalankan server jaringan lokal").
2. Terminal menampilkan alamat seperti `http://192.168.1.10:8080/`. Buka alamat itu dari HP atau komputer lain di jaringan yang sama.
3. Bila tidak bisa dibuka dari perangkat lain, izinkan Node.js di Windows Firewall (jaringan Private). Komputer server harus tetap menyala.
4. Alamat `http://IP-lokal` bukan HTTPS, jadi fitur enkripsi bawaan browser nonaktif. Aplikasi ini memakai library enkripsi sendiri sehingga sandi tetap berfungsi.
5. Tanpa server sama sekali: kirim file tunggal `GIT_anak_toko.html` (atau `dist/GIT_Asset_Online.html` dengan `data.js`) lewat WhatsApp atau flashdisk. Bisa dibuka tanpa internet.

## F. Sandi tidak bisa dipakai? Periksa
- Gunakan tombol "Tampilkan sandi" untuk memastikan ejaan. Keyboard HP sering mengubah huruf pertama menjadi kapital; kolom sandi sudah dimatikan autokapitalnya, tetapi sandi yang disalin dari chat bisa membawa spasi (spasi di ujung diabaikan otomatis).
- Pastikan `data.js` yang diunggah adalah file terbaru dan dibuat dengan sandi yang sama. File lama tetap memakai sandi lama.
- Tombol "Buat data.js" kini menguji sendiri bahwa sandi bisa membuka data sebelum file disimpan.
