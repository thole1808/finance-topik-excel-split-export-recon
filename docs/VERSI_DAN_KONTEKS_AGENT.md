# Versi aplikasi dan konteks untuk agent

**Versi saat ini: 1.1.13**

Sumber versi tunggal: [`src/version.py`](../src/version.py).

## Sebelum mengubah project

Agent/pengembang berikutnya sebaiknya membaca dokumen berikut sebelum
memodifikasi kode:

1. [`README.md`](../README.md) — setup, format input, build macOS/Windows,
   dan keterbatasan produk.
2. [`PANDUAN_FINANCE.md`](PANDUAN_FINANCE.md) — alur kerja Finance, kolom
   wajib, edit preview, dan pemeriksaan output.
3. [`src/app.py`](../src/app.py) — UI desktop, pencarian, paging, serta edit.
4. [`src/recon_exporter.py`](../src/recon_exporter.py) — aturan data dan
   pembuatan workbook.
5. Workflow dan konfigurasi installer di
   [`.github/workflows/windows-installer.yml`](../.github/workflows/windows-installer.yml)
   dan [`scripts/windows/ReconBuilder.iss`](../scripts/windows/ReconBuilder.iss).

## Serah-terima sesi — 16 Oktober 2026

Gunakan catatan ini saat melanjutkan pekerjaan pada sesi atau agent baru.

### Permintaan terakhir dan keputusan tata letak

- Pengguna mengoreksi ringkasan yang sebelumnya diletakkan di sisi kanan:
  ringkasan harus menjadi baris subtotal di tabel LB dan RB, tepat setelah
  kelompok `SUCCESS` dan tepat setelah kelompok `FAILED`.
- Mengacu pada `recon-2026-10-14-v1.xlsx` dan cuplikan yang dikirim:
  kolom F subtotal adalah `Amount + Margin Biller`; kolom G–K adalah
  `Amount`, `Total Fee`, `Customer Fee`, `LST Fee`, dan `Margin Biller`.
- Sisakan tiga baris kosong setelah subtotal SUCCESS sebelum baris transaksi
  FAILED. Subtotal FAILED juga diletakkan segera setelah transaksi FAILED.
- Subtotal status menjumlahkan transaksi Current Day dan For tomorrow.
  Nilai transaksi sumber/recon per baris tidak diubah. Footer SUM lama di
  bawah kolom nilai bukan lokasi subtotal.

### Implementasi yang sudah ada

- Logika berada di [`src/recon_exporter.py`](../src/recon_exporter.py):
  output mengurutkan SUCCESS lalu FAILED; menulis subtotal di baris kelompok;
  memakai formula Excel beserta cached values; dan memberi tiga baris sela
  setelah subtotal SUCCESS. Data Current Day dan For tomorrow dipasangkan
  berdasarkan status/urutan tanggal di baris yang sama.
- Versi aplikasi dinaikkan menjadi **1.1.13** di
  [`src/version.py`](../src/version.py). README dan panduan Finance menjelaskan
  lokasi serta pemetaan kolom subtotal.
- Validasi pada input lokal LB/RB menunjukkan subtotal formula/cached values
  di kedua sheet sama dengan perhitungan dari input untuk kelima nilai dan
  kedua status. Posisi subtotal SUCCESS, tiga baris kosong, dan subtotal FAILED
  juga diperiksa.
- Syntax check dan diagnostic Pylance untuk exporter lulus tanpa error.
- DMG 1.1.13 dibangun di `dist/ReconBuilder.dmg`; signature aplikasi serta
  checksum DMG valid. File dist tidak dilacak Git.
- Windows installer build workflow sukses; artifact tersedia di
  [GitHub Actions run 38882283710](https://github.com/thole1808/finance-topik-excel-split-export-recon/actions/runs/38882283710).

### Status Git dan artefak

- Branch `main` sudah dipush ke repo pengguna:
  [finance-topik-excel-split-export-recon](https://github.com/thole1808/finance-topik-excel-split-export-recon).
- Commit terbaru saat catatan ini dibuat: `89c8ee6` (`Document subtotal
  column layout`), setelah `e43dfa4` (`Place status totals below LB and RB
  groups`). Belum ada perubahan source yang tertinggal untuk dipush.
- DMG versi 1.1.13 tersedia di `dist/ReconBuilder.dmg` dan root
  `ReconBuilder.dmg`; keduanya identik. Root DMG disimpan untuk distribusi repo.
- Saat sesi serah-terima pertama ditulis, GitHub Release masih sampai
  `v1.1.10`. Perbaikan workflow terbaru menerbitkan Windows installer dan DMG
  bersama-sama pada release tag.
- `Archive.zip` dihapus setelah dipastikan hanya berisi salinan DMG lama dan
  flow PNG lama. Flow diagram yang digunakan dokumentasi berada di `docs/`.
- Cache Python dan direktori PyInstaller `build/` yang bisa dibuat ulang
  dibersihkan. ZIP installer lokal versi 1.1.2 dan 1.1.6 yang sudah usang
  dihapus; DMG dan ZIP versi 1.1.13 tetap tersedia di `dist/`.
- Workbook LB/RB, workbook referensi/output, dan `.venv` dipertahankan; file
  tersebut merupakan data kerja/environment lokal dan tidak dimasukkan Git.

### Konteks versi sebelumnya

- Versi 1.1.11 memperbaiki profil fee RB untuk tiga pasangan produk/nominal
  AXD10H|66750, AXD2H|25650, dan AXBR8H|72105; file acuan hanya preview dan
  bukan sumber data export.
- Versi 1.1.12 menambahkan blok total di sisi kanan untuk Amount saja. Pengguna
  mengoreksi tata letak dan cakupan tersebut; jangan kembali memakai lokasi
  ringkasan samping. Versi 1.1.13 menggantinya dengan subtotal di bawah grup.
- Jika pengguna meminta perubahan lanjutan, baca permintaan terbaru dan
  workbook acuan sebelum mengubah posisi atau cakupan subtotal. Jangan
  mengasumsikan total seluruh status menggantikan subtotal SUCCESS/FAILED.

## Fitur versi 1.1.0

- Preview menjelajahi semua baris LB/RB dengan pagination 100, 250, 500, atau
  1.000 baris per halaman.
- Pencarian literal tidak peka huruf besar/kecil di seluruh kolom dan semua
  baris pada file aktif.
- Edit nilai bertipe pada baris mana pun yang ditemukan; perubahan valid masuk
  ke export, dan file sumber tidak ditimpa.
- UI menampilkan nomor versi; versi yang sama tercatat di properti workbook.
- Workflow Windows membangun installer Inno Setup; tag versi menghasilkan
  GitHub Release berisi installer.
- Upload file recon referensi membaca sheet `Summary` dan `Records` untuk
  ditinjau; sheet dengan nama tersebut pada hasil tetap dihitung ulang dari
  data LB/RB. Hasil export memuat sheet `Summary`, `Records`, `LB`, dan `RB`.

## Perbaikan versi 1.1.1

- Memperbaiki path file referensi dalam Windows build agar PyInstaller dapat
  menemukannya ketika file spec berada di direktori build.

## Fitur versi 1.1.2

- Label pemilihan file di desktop menggunakan istilah **Upload**.
- Menambahkan upload workbook recon referensi; sheet `Summary` dan `Records`
  dimuat untuk preview read-only. Output tetap menghitung ulang sheet tersebut
  dari upload LB/RB dan menyertakan sheet `LB` serta `RB`.
- Diagram alur PNG/JPG menampilkan ketiga upload serta keempat sheet output.

## Perbaikan versi 1.1.3

- Sheet `LB` dan `RB` menampilkan hitungan `SUCCESS`/`FAILED` dan filter
  AutoFilter untuk `Transaction Status` serta `Recon Result`.
- Timestamp pada `Records`, `LB`, dan `RB` ditulis sebagai serial datetime
  Excel dengan format yang menyertakan jam, menit, dan detik; filter jam tidak
  lagi beroperasi pada nilai teks/tanggal tengah malam.

## Perbaikan versi 1.1.4

- Bagian `For tomorrow because different time` pada sheet LB/RB kini berisi
  hanya transaksi dengan jam mulai pukul 23:00:00; jam 00:00 tidak ikut.

## Perbaikan versi 1.1.5

- Baris transaksi pada sheet LB/RB otomatis dikelompokkan dengan `SUCCESS`
  terlebih dahulu dan `FAILED` di bawahnya; total keseluruhan tetap berada
  setelah semua transaksi.

## Perbaikan versi 1.1.6

- Transaksi `SUCCESS` dan `FAILED` dikelompokkan dengan urutan SUCCESS dahulu,
  FAILED di bawahnya, dengan warna hijau/merah hanya pada sel status.

## Perbaikan versi 1.1.7

- Pukul 23:00–23:59 disembunyikan dari tabel recon Current Day dan ditampilkan
  di bagian `For tomorrow`; baris asli tetap ada di blok data sumber LB/RB.
- Semua tanggal dalam file LB/RB diproses, diurutkan kronologis di dalam
  kelompok SUCCESS/FAILED, dan tidak ada lagi batas tetap 50.000 baris pada
  formula pencarian refund.
- Warna hijau/merah dibatasi ke sel `Transaction Status` SUCCESS/FAILED saja;
  sel lain dan seluruh baris tidak diberi warna status.

## Perbaikan versi 1.1.8

- Memastikan transaksi pukul 23:00–23:59 tidak muncul pada tabel recon
  `Current Day`; transaksi tersebut hanya muncul pada data sumber dan tabel
  `For tomorrow`.
- Menjaga referensi formula Current Day tetap menunjuk ke baris data sumber
  yang tepat setelah baris terlambat dipisahkan.

## Perbaikan versi 1.1.9

- Latar hijau SUCCESS dan merah FAILED hanya diterapkan pada sel nilai kolom
  `Transaction Status` di blok recon Current Day/For tomorrow; sel status
  sumber dan kolom lainnya tidak diberi warna.

## Perbaikan versi 1.1.10

- Pada sheet RB, blok data sumber Current Day mengecualikan transaksi mulai
  pukul 23:00 dan blok sumber For tomorrow hanya memuat transaksi mulai pukul
  23:00, sama seperti blok recon.
- Formula LB/RB Current Day kini mereferensikan baris terkompaksi yang sesuai
  setelah transaksi terlambat dipisahkan.

## Perbaikan versi 1.1.11

- Menambahkan profil fee RB dari workbook acuan 15 Oktober 2026 untuk
  AXD10H|66750, AXD2H|25650, dan AXBR8H|72105:
  Total 1.500, Customer 750, LST 750, Margin 0.
- Fee di recon referensi hanya ditampilkan sebagai preview; output tetap dibuat
  ulang dari upload LB/RB. Sebelum profil tersebut ditambahkan, tiga kombinasi
  belum terdaftar dan input RB tidak memiliki kolom fee, sehingga sel fee
  output menjadi kosong.
- Memperbarui diagram alur Finance untuk memuat jalur lookup fee, pemisahan
  status/jam, empat sheet output, serta pemeriksaan data Finance.

## Perbaikan versi 1.1.13

- Menempatkan subtotal SUCCESS dan FAILED setelah kelompok transaksinya di
  sheet LB/RB, mengikuti baris total pada workbook recon contoh.
- Menjumlahkan Amount dan empat kolom fee, termasuk transaksi For tomorrow;
  menyisakan tiga baris pemisah setelah subtotal SUCCESS.

## Perbaikan versi 1.1.12

- Menghapus baris footer auto-SUM dari kolom Amount, Total Fee, Customer Fee,
  LST Fee, dan Margin Biller pada sheet LB/RB.
- Menambahkan blok SUCCESS/FAILED/TOTAL AUTO SUM untuk Amount, mencakup
  transaksi Current Day serta For tomorrow.
- Menambahkan tombol **Edit sel terpilih** pada preview LB/RB; double-click
  cell tetap tersedia sebagai cara cepat.
- Memperbarui gaya UI desktop dan memastikan metadata versi DMG/Inno Setup
  mengikuti `src/version.py`.

## Gambar yang tersedia

Gambar berikut adalah diagram alur aplikasi versi 1.1.11, disediakan untuk
kebutuhan dokumentasi/presentasi. Sumber vector ada di
[`flow-recon-finance.svg`](flow-recon-finance.svg):

- [`flow-recon-finance.png`](flow-recon-finance.png) — PNG 1800×1800.
- [`flow-recon-finance.jpg`](flow-recon-finance.jpg) — JPEG 1800×1800.

Diagram menerangkan upload LB/RB/recon referensi → validasi → preview/edit →
proses Polars → export XlsxWriter → sheet Summary/Records/LB/RB → pemeriksaan
Finance. Gambar ini bukan screenshot interaktif aplikasi.

## Riwayat versi ringkas

| Versi | Perubahan utama |
|---|---|
| 1.1.13 | Menempatkan subtotal Amount dan empat fee setelah kelompok SUCCESS dan FAILED di sheet LB/RB. |
| 1.1.12 | Memindahkan total Amount ke ringkasan SUCCESS/FAILED/Total, memperjelas edit preview, dan menyelaraskan versi installer. |
| 1.1.11 | Mengisi profil fee RB untuk tiga kombinasi kode/nominal dan memperjelas alur data/fee. |
| 1.1.10 | Pisahkan data sumber RB Current Day dan For tomorrow berdasarkan jam 23:00. |
| 1.1.9 | Warna SUCCESS/FAILED hanya pada nilai kolom Transaction Status di blok recon. |
| 1.1.8 | Pindahkan transaksi 23:00–23:59 dari tampilan Current Day ke For tomorrow tanpa menghapus data sumber. |
| 1.1.7 | Pertahankan transaksi jam 23 di data utama, dukung semua tanggal otomatis, dan warnai sel status saja. |
| 1.1.6 | Mengurutkan SUCCESS di atas FAILED dan memberi warna hijau/merah pada baris transaksi. |
| 1.1.5 | Mengurutkan transaksi SUCCESS di atas FAILED pada sheet LB/RB. |
| 1.1.4 | Bagian For tomorrow hanya memuat transaksi mulai jam 23:00. |
| 1.1.3 | Filter dan hitungan SUCCESS/FAILED pada sheet LB/RB; mempertahankan timestamp sebagai datetime Excel. |
| 1.1.2 | Upload recon referensi, preview Summary/Records, dan flow diagram yang diperbarui. |
| 1.1.1 | Memperbaiki path resource pada build installer Windows. |
| 1.1.0 | Preview/pencarian seluruh data, highlight status hijau, versioning, dan installer Windows otomatis. |
| 1.0.0 | Export workbook LB/RB menggunakan Polars dan XlsxWriter; UI dan build macOS. |

## Aturan rilis dan data

- Ubah `__version__` di `src/version.py` pada setiap rilis; gunakan format
  semantic versioning `MAJOR.MINOR.PATCH`.
- Tag Git harus sama dengan versi dan diawali `v` (mis. `v1.1.13`). Tag
  menjalankan Windows build dan menerbitkan installer Windows serta DMG ke
  GitHub Release.
- DMG dibangun di macOS melalui `scripts/build_dmg.sh`; installer Windows
  dibuat di runner Windows melalui GitHub Actions.
- Jangan menambahkan workbook transaksi, data pelanggan, output, `.venv`,
  `build/`, atau `dist/` ke Git.
- Matching silang settlement LB/RB belum didefinisikan oleh aturan bisnis;
  jangan mengubah label matching sebelum kunci/toleransi disepakati Finance.
