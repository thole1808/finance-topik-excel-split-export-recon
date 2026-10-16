# Panduan Finance — Recon Builder

Panduan ini menjelaskan cara menyiapkan file, menjalankan aplikasi, dan
memeriksa hasil export. Aplikasi hanya memproses workbook Excel `.xlsx`.

## Gambaran alur

![Diagram alur Recon Builder untuk Finance](flow-recon-finance.png)

Versi JPG tersedia di [flow-recon-finance.jpg](flow-recon-finance.jpg).
File sumber diagram: [flow-recon-finance.svg](flow-recon-finance.svg).

Alur singkat: **Operator → upload LB/RB/recon dan Backdate opsional → validasi →
preview per sheet/koreksi → proses dengan terminal → export bertimestamp →
riwayat aktivitas → periksa Summary, Records, LB, dan RB**.
Jumlah baris, tanggal, nominal, dan nilai lain diambil dari file yang diupload
setiap kali proses berjalan. Data pada sheet Summary dan Records dari file recon
referensi ditampilkan untuk ditinjau; Summary dan Records pada output dihitung
ulang dari upload LB/RB. Jika Backdate disertakan, sheet RECON Backdate
divalidasi, tetapi tidak menggantikan nilai dari sumber LB/RB.

## Menjalankan aplikasi

1. Di macOS, buka DMG dan jalankan installer PKG untuk memasang **Recon Builder**
   ke folder Applications, lalu buka dari Launchpad. Di Windows, jalankan
   `ReconBuilder-Setup.exe`; aplikasi dipasang ke Program Files dan shortcut
   tersedia di Start Menu (shortcut Desktop juga dibuat secara default).
   (Untuk source development, jalankan `python3 app.py`.)
   Aplikasi otomatis memakai nama akun PC/laptop yang sedang login untuk
   mencatat riwayat; tidak ada prompt operator. Jendela dimaksimalkan dan jam
   lokal tampil di header.
2. Pada **Upload LB**, pilih workbook Loket Bayar berformat `.xlsx`; tunggu
   preview selesai dimuat.
3. Pada **Upload RB**, pilih workbook Rajabiller berformat `.xlsx`; tunggu
   preview selesai dimuat.
4. Pada **Upload Recon**, pilih file recon `.xlsx` yang memiliki sheet
   **Summary** dan **Records**. Kedua sheet dimuat sebagai referensi dan dapat
   ditinjau melalui tab **Recon · Summary** dan **Recon · Records**.
5. Jika ada data Backdate, pilih satu workbook `.xlsx` pada **File Backdate
   (opsional)**. Workbook harus memiliki sheet `LB`, `RB`, dan `RECON` dengan
   header yang cocok. Jika tidak ada Backdate, biarkan field ini kosong.
6. Tinjau tab **Preview LB** dan **Preview RB**. Preview dapat menelusuri
   seluruh baris file: pilih jumlah baris per halaman atau **Semua**, lalu gunakan
   **Sebelumnya/Berikutnya**. Gunakan **Cari seluruh data** untuk menemukan
   nilai di semua kolom, termasuk transaksi pada halaman lain.
   Sheet Backdate tampil sebagai tab **Backdate · LB**, **Backdate · RB**, dan
   **Backdate · RECON**, sekalipun semuanya berasal dari satu workbook.
7. Untuk koreksi LB/RB, pilih sel lalu klik **Edit sel terpilih**, atau klik
   ganda sel. Ubah nilai, lalu tekan **Enter** atau pindah fokus untuk
   menerapkan; **Esc** membatalkan. Perubahan pada baris preview ikut
   diterapkan ke export. Baris lain tetap berasal dari file upload.
   Preview recon bersifat referensi dan tidak dapat diedit.
8. Folder output otomatis memakai folder data aplikasi pada komputer. Nama file
   dibuat dari tanggal dan waktu lokal, misalnya
   `ReconBuilder_20261017_143025_123456.xlsx`.
9. Klik **Generate Recon**. Terminal proses menampilkan tahapan penulisan dan
   jumlah baris Records yang diproses. Tunggu pesan sukses.
10. Buka file output. Tanpa Backdate, hasilnya sama seperti proses reguler.
    Dengan Backdate, transaksi yang dibentuk dari sumber `LB`/`RB` Backdate
    muncul di bagian awal `Records` berdasarkan Datetime, lalu diikuti hasil
    End Date. Nilai `Ref` dan `Vendor Ref` tetap bersumber dari kolom asli
    LB/RB. Data
    sumber Backdate digabung ke sheet `LB`/`RB` yang sama dan ditampilkan
    sebelum data End Date pada kelompok status masing-masing. `Summary`
    menghitung gabungan Backdate dan End Date. Transaksi End Date pukul
    23:00–23:59 ditampilkan di bagian **For tomorrow because different time**;
    transaksi Backdate pukul 23:00 atau lebih tetap terlihat di blok utama.
11. **Buka Folder Output** akan memilih file terakhir yang berhasil dibuat
    langsung di foldernya.
12. Buka tab **File Hasil** untuk melihat workbook yang tersimpan. Pilih file
    lalu tekan **Buka File** atau **Unduh / Simpan Salinan…** untuk menyalinnya
    ke lokasi yang mudah dibagikan.

## Update aplikasi dan notifikasi suara

Tombol **Periksa Update** ada di kanan atas header aplikasi, di samping jam.
Header menampilkan versi aplikasi dan versi release. Saat aplikasi dibuka,
aplikasi memeriksa GitHub Release. Jika versi baru tersedia, tombol berubah
menjadi **Update v...** dan header menampilkan versi release yang tersedia.
Klik tombol tersebut untuk mengunduh installer sesuai sistem operasi; aplikasi memeriksa
SHA-256 sebelum memasang dan membuka kembali aplikasi versi baru. macOS dapat meminta
otorisasi administrator bila aplikasi berada di folder sistem. Windows dapat
meminta konfirmasi UAC. Koneksi internet dan release GitHub yang berhasil
diterbitkan diperlukan. Update otomatis tersedia untuk macOS dan Windows.
Nomor versi pada aplikasi berasal dari `src/version.py`; release memakai tag
`v<versi>` yang sama. Workflow menolak tag yang tidak cocok dengan versi source.

Aplikasi membacakan “Mohon tunggu. Proses generate file recon sedang berjalan.”
saat Generate dimulai dan membacakan pemberitahuan sukses setelah file selesai
dibuat. Tombol **Tes Suara** di bagian bawah aplikasi tersedia satu kali saja;
setelah dipakai, tombol disembunyikan pada sesi berikutnya. Aplikasi memilih
voice Indonesia yang terpasang di sistem, lalu memakai voice bawaan jika tidak
ada. Voice sistem adalah text-to-speech, bukan rekaman manusia.

Jam menggunakan tanggal dan waktu lokal dari komputer, termasuk zona waktu
yang diatur pada sistem operasi, dengan nama hari dan bulan berbahasa Indonesia.
Nama file output dan waktu riwayat tetap memakai waktu lokal komputer. Kolom
Waktu Riwayat dan CSV menggunakan nama hari/bulan Indonesia dan zona waktu PC.

Jendela dimaksimalkan saat dibuka, mengisi area yang tersedia pada laptop atau
layar lebar, dan kontrol preview menyesuaikan ukuran jendela. Lebar minimum
jendela 960 piksel; perkecil atau maksimalkan jendela sesuai ruang kerja.

Ikon Recon Builder disertakan dalam aplikasi macOS dan installer Windows, lalu
ikut diperbarui bersama versi aplikasi. Diagram alur tersedia sebagai PNG dan
JPG pada dokumen ini dan sebagai aset pada GitHub Release.

Setiap push ke `main` dengan versi baru di `src/version.py` membangun installer
Windows dan mencatat Git tag `v<versi>` serta GitHub Release yang memuat DMG,
installer PKG macOS, installer Windows, diagram PNG/JPG, dan checksum. Siapkan
DMG dan PKG hasil build macOS di root repository sebelum push versi baru.

## Riwayat penggunaan dan audit

Tab **Riwayat** menampilkan aktivitas operator dengan waktu lokal, nama akun PC
yang sedang login, perangkat, versi aplikasi, IP lokal tersamarkan, MAC
tersamarkan, aksi, dan rinciannya. IP disimpan sebagai tiga oktet pertama
(misalnya `192.168.1.*`); MAC hanya menampilkan tiga oktet awal
(misalnya `AA:BB:CC:**:**:**`). Alamat mentah tidak disimpan. Aplikasi
mencatat pembukaan/penutupan, tab yang dibuka, file yang dipilih/dimuat, pencarian,
perubahan sel (nilai lama dan baru), navigasi preview, proses generate (hasil,
durasi, ukuran atau kegagalan), dan pembukaan lokasi output. Gunakan
**Periksa integritas** untuk memvalidasi rantai hash, atau **Ekspor CSV** untuk
menyimpan seluruh log.

Database SQLite tersimpan per akun komputer: macOS
`~/Library/Application Support/Recon Builder/usage-history.sqlite3`, Windows
`%APPDATA%/Recon Builder/usage-history.sqlite3`, Linux
`~/.local/share/recon-builder/usage-history.sqlite3` (kecuali XDG_DATA_HOME
ditetapkan). Simpan/backup file tersebut sesuai kebijakan Finance. Rantai hash
mendeteksi perubahan pada rekaman yang tersisa, tetapi log lokal bukan bukti
non-repudiation: akun PC tidak membuktikan siapa orang yang memakai perangkat,
dan pemilik perangkat dapat menghapus atau mengganti database. Untuk pembuktian lintas operator yang
kuat, aplikasi perlu login terverifikasi (misalnya SSO) dan mengirim event ke
server audit terpusat dengan hak tulis terbatas serta retensi yang dikelola.
Rincian audit menyimpan nilai sel yang diedit dan query pencarian; batasi akses
ke database/CSV ekspor sesuai kebijakan data Finance.

Dialog pemilihan file dibatasi ke workbook `.xlsx`. Format `.xls`, `.xlsm`,
`.csv`, PDF, dan file selain Excel tidak didukung. Sheet pertama pada setiap
workbook yang dipilih akan dibaca.

## Format sumber yang diperlukan

Gunakan template/export standar dengan nama header berikut pada sheet pertama.
Header harus cocok; jangan mengganti nama kolom wajib.

**LB — Loket Bayar**

`Tanggal`, `Operator`, `Nama`, `Idpel`, `Produk`, `Periode`, `Status`,
`No. Reff (SN)`, `Lembar`, `Tagihan`, `Denda`, `Admin`, `Harga`, `Fee Loket`,
`Fee Bulan`, `Total`, `Ref ID`.

**RB — Rajabiller**

`No`, `ID Trx`, `Tanggal`, `Kode Produk`, `No Pelanggan-1`,
`No Pelanggan-2`, `No Pelanggan-3`, `No Pelanggan-4`, `Via`, `SN`, `Ref`,
`Status`, `Nominal`.

Baris footer/ringkasan RB yang kolom `No`-nya kosong atau bukan angka
dilewati. Baris transaksi dengan nomor valid tetap diproses. Data tidak perlu
memiliki jumlah baris atau tanggal tertentu.

## Isi file hasil

| Sheet | Isi | Cara menggunakan |
|---|---|---|
| `Summary` | Ringkasan kategori/status. Tanpa Backdate dihitung dari upload LB/RB; dengan Backdate merangkum kedua periode. | Gunakan untuk pemeriksaan awal, lalu telusuri transaksi di `Records`. |
| `Records` | Daftar transaksi gabungan. Jika Backdate ada, transaksi yang dibentuk dari sumber LB/RB ditempatkan lebih dulu berdasarkan Datetime, kemudian baris End Date. | Gunakan filter header untuk mencari produk, referensi, status, atau biller. |
| `LB` | Data sumber LB dan kolom recon/formula sesuai layout acuan; data Backdate digabung lebih dulu bila diupload. Waktu Backdate pukul 23:00 ke atas tetap terlihat di blok utama. | Bandingkan kolom hasil recon dengan data sumber pada bagian LB. |
| `RB` | Data sumber RB dan kolom recon/formula sesuai layout acuan; data Backdate digabung lebih dulu bila diupload. Jika sheet RB Backdate sumber kosong, tidak ada transaksi Backdate yang dapat ditambahkan. | Bandingkan kolom hasil recon dengan data sumber pada bagian RB. |

Header dibekukan ketika scroll; `Records` memiliki filter, dan format angka,
tanggal, warna, serta lebar kolom mengikuti workbook acuan agar nilai mudah
dibaca. Nominal transaksi tetap berasal dari input, sedangkan fee menggunakan
profil kode produk dan nominal bila tersedia. Untuk kombinasi baru, aplikasi
menggunakan tabel fee fallback, lalu field fee LB jika tersedia.
Untuk transaksi RB pada acuan 15 Oktober 2026, AXD10H nominal 66.750,
AXD2H nominal 25.650, dan AXBR8H nominal 72.105 menggunakan fee
1.500 / 750 / 750 / 0 (Total / Customer / LST / Margin).
Penyebab fee transaksi tersebut sebelumnya kosong adalah kombinasi kode produk
dan nominalnya belum memiliki profil fee. File RB mentah hanya memiliki kolom
nominal/status, bukan kolom fee; sementara workbook recon referensi hanya
ditampilkan sebagai preview dan tidak disalin menjadi sumber data output.
Karena itu nilai fee dari sheet `Records` referensi tidak otomatis masuk ke
export baru. Sejak versi 1.1.11 ketiga kombinasi tersebut sudah tercatat di
profil aplikasi. Untuk produk/nominal baru yang belum punya profil, jangan
menganggap sel kosong berarti fee nol; minta Finance mengonfirmasi nilai dan
tambahkan profil sebelum memakai hasil sebagai laporan final.

Pada sheet `LB` dan `RB`, baris total ditempatkan langsung setelah kelompok
`SUCCESS` dan setelah kelompok `FAILED`, sejajar dengan kolom angka pada
contoh recon. Tiga baris kosong memisahkan total SUCCESS dari transaksi
FAILED. Amount dan empat kolom fee dijumlahkan masing-masing untuk statusnya.
Jika Backdate disertakan, subtotal SUCCESS/FAILED menggabungkan kedua periode;
transaksi For tomorrow juga ikut dihitung. Baris total mengikuti contoh:
kolom pertama menunjukkan Amount + Margin Biller, kemudian Amount, Total Fee,
Customer Fee, LST Fee, dan Margin Biller. Nilai transaksi per baris tidak
berubah.

Pada sheet `LB` dan `RB`, transaksi otomatis diurutkan: semua `SUCCESS` tampil
lebih dahulu, diikuti semua `FAILED`, lalu status lain bila ada. Dalam setiap
kelompok status, tanggal/waktu diurutkan dari yang paling awal. Total SUCCESS
dan FAILED muncul setelah masing-masing kelompok, dengan tiga baris kosong
sebagai jarak setelah total SUCCESS. Baris kedua tetap menampilkan jumlah
transaksi `SUCCESS` dan `FAILED` untuk tabel Current Day; dropdown filter di
header `Transaction Status` atau `Recon Result` dapat digunakan untuk
pemeriksaan lebih lanjut. Warna hijau untuk `SUCCESS` dan merah untuk
`FAILED` hanya diterapkan pada nilai di kolom `Transaction Status` pada blok
recon Current Day dan For tomorrow; kolom status mentah sumber serta sel lain
tidak diberi warna.

Semua tanggal yang ada di file LB dan RB diproses otomatis; aplikasi tidak
membatasi jumlah hari. Transaksi End Date pukul 23:00–23:59 ditampilkan di
blok recon dan data sumber **For tomorrow because different time**, serta tetap
ada di `Records`. Transaksi Backdate pukul 23:00 atau lebih tetap di blok
utama. Pada LB, tabel recon For tomorrow memuat kolom Datetime sampai
`Recon Result`, mulai tepat di bawah header, dan data sumber tampil di
sampingnya. Baris total per status mencakup semua transaksi sumber.

Kolom `Datetime` pada `Records`, `LB`, dan `RB` disimpan sebagai nilai tanggal/
waktu Excel, bukan teks. Untuk mengambil transaksi hari ini pukul 23.00 sampai
23.59, filter tanggal hari ini dan batas waktu **23:00:00–23:59:59**. Nilai
timestamp tetap menyertakan jam, menit, dan detik.

Bagian **For tomorrow because different time** pada sheet `LB`/`RB` memuat
transaksi End Date mulai pukul **23:00:00**. Transaksi Backdate pukul 23 atau
lebih tetap ditampilkan di blok utama.

File recon referensi hanya dibaca untuk memuat dan meninjau sheet `Summary`
serta `Records`. Isinya tidak menggantikan hasil; sheet `Summary` dan `Records`
pada output selalu dihitung ulang menggunakan upload LB/RB yang sedang diproses.

Preview menampilkan data dalam halaman agar file besar tetap mudah ditinjau;
opsi **Semua** menampilkan seluruh baris pada satu halaman. Setiap sheet
Backdate dapat ditinjau lewat tabnya sendiri.
Pencarian mencakup semua baris, bukan hanya halaman aktif. Perubahan yang
disimpan di preview memengaruhi baris yang sama pada sheet input dan hasil
gabungan; nilai perhitungan terkait akan dibuat ulang oleh exporter. File
sumber tidak dimodifikasi.

## Pemeriksaan sebelum memakai hasil

1. Pastikan keempat sheet tersedia dan jumlah transaksi sesuai file sumber
   (footer RB yang dibuang tidak dihitung).
2. Cocokkan beberapa referensi, tanggal, status, dan nominal antara sumber dan
   sheet `LB`/`RB`.
3. Periksa fee untuk kode produk/nominal yang baru atau belum ada di profil.
4. Periksa nilai Summary, terutama transaksi berstatus gagal/refund.
5. Simpan file sumber dan hasil sesuai kebijakan penyimpanan Finance.

## Batasan penting klasifikasi recon

Versi saat ini **belum melakukan pencocokan pasangan transaksi LB terhadap RB**
berdasarkan `Ref`, `Ref ID`, nominal, atau kunci gabungan lain. Status transaksi
dinormalisasi menjadi `SUCCESS` atau `FAILED`; status sukses menghasilkan
`MATCHED`, sedangkan status selain sukses menghasilkan `INTERNAL`. Karena itu,
label `MATCHED`/`INTERNAL` pada versi ini bukan bukti bahwa transaksi telah
dicocokkan silang antar-biller. Minta aturan matching Finance (kunci transaksi,
toleransi nominal/waktu, dan perlakuan refund) sebelum menggunakan output untuk
keputusan settlement.

## Jika proses gagal

- **Header wajib tidak ditemukan:** pastikan memilih file sumber yang benar dan
  kembalikan nama header ke format standar.
- **File tidak terbaca:** pastikan file benar-benar `.xlsx`, tidak rusak, dan
  dapat dibuka di Excel/LibreOffice.
- **Output gagal dibuat:** tutup file output jika sedang terbuka, lalu pilih
  lokasi yang dapat ditulis dan nama berakhiran `.xlsx`.
