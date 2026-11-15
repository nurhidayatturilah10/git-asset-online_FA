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

## D. Link anak toko yang otomatis berubah (hanya lihat + unduh Excel)
Link anak toko = link github.io biasa (tanpa login akun apa pun). Anak toko cukup memasukkan sandi.

Pengaturan sekali saja (halaman admin = link yang sama + `?admin` di belakang, atau file lokal / server lokal):
1. Buka tab **Link Anak Toko**. Isi **sandi** (atau klik Buat sandi acak). Sandi tampil jelas di kolom dan bisa disalin untuk dibagikan.
2. Buat token GitHub: GitHub > Settings > Developer settings > Personal access tokens > Fine-grained tokens. Pilih hanya repositori ini, izin **Contents: Read and write**. Salin token.
3. Isi nama pengguna GitHub, nama repositori, cabang (main) dan token. Biarkan **Terbitkan otomatis** tercentang. Klik **Simpan dan terbitkan sekarang**. Pengaturan tersimpan di browser perangkat admin.

Setiap update berikutnya:
- Cukup unggah `REPORT_GIT.xlsx` terbaru di halaman admin. Data dienkripsi dengan sandi yang sama dan dikirim sendiri ke repositori. Tidak perlu membuat atau mengunggah `data.js`.
- Link anak toko memuat data baru dalam 1 sampai 2 menit (GitHub Pages membangun ulang). Halaman anak toko yang sedang terbuka memeriksa data baru tiap 3 menit dan saat tab dibuka kembali, lalu memperbarui diri tanpa meminta sandi lagi.
- Anak toko dapat mencentang **Ingat sandi di perangkat ini** agar tidak mengetik sandi lagi.

Catatan keamanan:
- Hanya data terenkripsi (AES-256-GCM) yang masuk ke repositori. Kekuatan sandi menentukan keamanan; pakai sandi panjang. Siapa pun yang tahu sandi bisa melihat semua toko.
- Token memberi akses tulis ke repositori. Simpan hanya di komputer admin, batasi ke satu repositori, dan cabut bila perangkat hilang. Jika ingin tanpa token, gunakan tombol "Unduh data.js (cara manual)" lalu unggah file itu ke repositori.
- Mengganti sandi: ubah sandi di tab, lalu terbitkan lagi. Perangkat anak toko yang menyimpan sandi lama akan diminta sandi baru.
- Riwayat Git menyimpan versi `data.js` lama yang terbuka dengan sandi lamanya.
- Link claude.ai tidak cocok untuk anak toko: data bersama di sana hanya dapat dibaca orang yang login ke akun claude.ai.

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

## G. Master DAKAR, Pejabat Toko otomatis, dan GIT SOZ1 (Office & Warehouse)
Semua data ada di satu file `REPORT_GIT.xlsx`. Aplikasi hanya meminta unggahan file ini.
1. Sheet **DATA KARYAWAN**: tempel sheet `EXIST` (karyawan aktif) dari file DAKAR apa adanya (header di baris 2, data mulai baris 3; kolom NIK, NAME, POSITION CODE DESC, DEPARTMENT, STORE CODE, LOCATION, POSITION GROUPING). Kolom rumus bantu di sebelah kanan jangan dihapus. Sheet OUT tidak dipakai.
2. Sheet **PEJABAT TOKO** terisi otomatis: karyawan CHIEF OF STORE (COS); bila tidak ada, ASISTANT CHIEF OF STORE (ACOS); bila beberapa orang, NIK terkecil (tertua). Toko yang tidak punya COS/ACOS (mis. franchise) diisi manual di kolom G:H (kuning).
3. Aplikasi menghitung Pejabat Toko langsung dari DATA KARYAWAN dengan aturan yang sama (tidak bergantung pada hasil rumus Excel).
4. **GIT SOZ1**: aktivasi dan mutasi dengan UU tujuan SOZ1 yang punya user penerima (NIK tujuan). Lokasi kerja penerima dibaca dari master DAKAR: DC = Warehouse, OFFICE = Office. Item tanpa NIK tujuan tidak ditampilkan. Daftar disusun per penerima: klik nama untuk menampilkan atau menyembunyikan item (awalnya tersembunyi), tombol **Share WA** per penerima, dan **Share semua** untuk pesan seluruh penerima.
5. **Office SOZ1** dan **Warehouse SOZ1** berfungsi seperti toko: bisa dicari di Cari Toko, punya profil (belum selesai dan histori dengan nama penerima, NIK, lokasi kerja), bisa dishare lewat Share ke Toko, dan bisa diunduh ke Excel.
   Unduhan Excel SOZ1 (menu GIT SOZ1 dan profil Office/Warehouse) berisi sheet Ringkasan Penerima, Mutasi, dan Aktivasi: item belum selesai dan riwayat selesai dengan kolom KONDISI, tanggal asli, NIK, nama, lokasi kerja, departemen, posisi, dan filter di judul kolom.
6. Bila pengaturan GitHub (bagian D) aktif, link anak toko ikut diperbarui otomatis setiap kali `REPORT_GIT.xlsx` diunggah.
