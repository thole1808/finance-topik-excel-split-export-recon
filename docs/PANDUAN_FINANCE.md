# Panduan Finance — Recon Builder

Panduan ini menjelaskan cara menyiapkan file, menjalankan aplikasi, dan
memeriksa hasil export. Aplikasi hanya memproses workbook Excel `.xlsx`.

## Gambaran alur

![Diagram alur Recon Builder untuk Finance](flow-recon-finance.png)

Versi JPG tersedia di [flow-recon-finance.jpg](flow-recon-finance.jpg).

Alur singkat: **Upload LB + RB + file recon referensi → validasi → preview/
koreksi → proses data → export Excel → periksa Summary, Records, LB, dan RB**.
Jumlah baris, tanggal, nominal, dan nilai lain diambil dari file yang diupload
setiap kali proses berjalan. Data pada sheet Summary dan Records dari file recon
referensi ditampilkan untuk ditinjau; Summary dan Records pada output dihitung
ulang dari upload LB/RB.

## Menjalankan aplikasi

1. Buka **ReconBuilder.app** (atau jalankan `python3 app.py` dari source).
2. Pada **Upload LB**, pilih workbook Loket Bayar berformat `.xlsx`; tunggu
   preview selesai dimuat.
3. Pada **Upload RB**, pilih workbook Rajabiller berformat `.xlsx`; tunggu
   preview selesai dimuat.
4. Pada **Upload Recon**, pilih file recon `.xlsx` yang memiliki sheet
   **Summary** dan **Records**. Kedua sheet dimuat sebagai referensi dan dapat
   ditinjau melalui tab **Recon · Summary** dan **Recon · Records**.
5. Tinjau tab **Preview LB** dan **Preview RB**. Preview dapat menelusuri
   seluruh baris file: pilih jumlah baris per halaman lalu gunakan
   **Sebelumnya/Berikutnya**. Gunakan **Cari seluruh data** untuk menemukan
   nilai di semua kolom, termasuk transaksi pada halaman lain.
6. Untuk koreksi LB/RB, klik ganda sel. Ubah nilai, lalu tekan **Enter** atau pindah
   fokus untuk menerapkan; **Esc** membatalkan. Perubahan pada baris preview
   ikut diterapkan ke export. Baris lain tetap berasal dari file upload.
   Preview recon bersifat referensi dan tidak dapat diedit.
7. Tentukan lokasi dan nama file output yang berakhiran `.xlsx`.
8. Klik **Generate Recon** dan tunggu pesan sukses.
9. Buka file output. Periksa sheet **Summary** dan **Records** yang dihitung
   dari upload LB/RB, lalu telusuri data sumber dan formula recon pada sheet
   **LB** dan **RB**.

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
| `Summary` | Ringkasan jumlah dan nominal transaksi per kategori/status, dihitung ulang dari upload LB/RB. | Gunakan untuk pemeriksaan awal, lalu telusuri transaksi di `Records`. |
| `Records` | Daftar gabungan baris LB dan RB, diurutkan berdasarkan tanggal, dihitung ulang dari upload LB/RB. | Gunakan filter header untuk mencari produk, referensi, status, atau biller. |
| `LB` | Data sumber LB dan kolom recon/formula sesuai layout acuan. | Bandingkan kolom hasil recon dengan data sumber pada bagian LB. |
| `RB` | Data sumber RB dan kolom recon/formula sesuai layout acuan. | Bandingkan kolom hasil recon dengan data sumber pada bagian RB. |

Header dibekukan ketika scroll; `Records` memiliki filter, dan format angka,
tanggal, warna, serta lebar kolom mengikuti workbook acuan agar nilai mudah
dibaca. Nominal transaksi tetap berasal dari input, sedangkan fee menggunakan
profil kode produk dan nominal bila tersedia. Untuk kombinasi baru, aplikasi
menggunakan tabel fee fallback, lalu field fee LB jika tersedia.

Pada sheet `LB` dan `RB`, transaksi otomatis diurutkan: semua `SUCCESS` tampil
lebih dahulu, diikuti semua `FAILED`, lalu status lain bila ada. Dalam setiap
kelompok status, tanggal/waktu diurutkan dari yang paling awal. Baris total
diletakkan setelah seluruh transaksi dan menjumlahkan semua status. Baris kedua
tetap menampilkan jumlah `SUCCESS` dan `FAILED`; dropdown filter di header
`Transaction Status` atau `Recon Result` dapat digunakan untuk pemeriksaan
lebih lanjut. Warna hijau untuk `SUCCESS` dan merah untuk `FAILED` hanya
diterapkan pada sel status, bukan seluruh baris.

Semua tanggal yang ada di file LB dan RB diproses otomatis; aplikasi tidak
membatasi jumlah hari. Transaksi pukul 23:00–23:59 tetap tercatat di bagian
utama/current day sesuai tanggal sumbernya dan juga disalin ke bagian **For
tomorrow because different time**. Bagian For tomorrow adalah salinan, bukan
pemindahan atau penghapusan dari data utama.

Kolom `Datetime` pada `Records`, `LB`, dan `RB` disimpan sebagai nilai tanggal/
waktu Excel, bukan teks. Untuk mengambil transaksi hari ini pukul 23.00 sampai
23.59, filter tanggal hari ini dan batas waktu **23:00:00–23:59:59**. Nilai
timestamp tetap menyertakan jam, menit, dan detik.

Bagian **For tomorrow because different time** pada sheet `LB`/`RB` menampilkan
salinan transaksi dari tanggal mana pun yang waktunya **23:00:00 atau lebih**.
Data sebelum pukul 23.00, termasuk pukul 00.00, tidak dimasukkan ke bagian
tersebut.

File recon referensi hanya dibaca untuk memuat dan meninjau sheet `Summary`
serta `Records`. Isinya tidak menggantikan hasil; sheet `Summary` dan `Records`
pada output selalu dihitung ulang menggunakan upload LB/RB yang sedang diproses.

Preview menampilkan data dalam halaman agar file besar tetap mudah ditinjau.
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
