# Versi aplikasi dan konteks untuk agent

**Versi saat ini: 1.1.6**

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

## Fitur versi 1.1.0

- Preview menjelajahi semua baris LB/RB dengan pagination 100, 250, 500, atau
  1.000 baris per halaman.
- Pencarian literal tidak peka huruf besar/kecil di seluruh kolom dan semua
  baris pada file aktif.
- Edit nilai bertipe pada baris mana pun yang ditemukan; perubahan valid masuk
  ke export, dan file sumber tidak ditimpa.
- `Recon Result = MATCHED` dan formula `Status = TRUE` memakai highlight hijau.
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
  FAILED di bawahnya, dan total akhir setelah transaksi.
- Baris SUCCESS diberi latar hijau; baris FAILED diberi latar merah pada sheet
  `Records`, `LB`, dan `RB`.

## Gambar yang tersedia

Gambar berikut adalah diagram alur aplikasi yang sama, disediakan untuk
kebutuhan dokumentasi/presentasi:

- [`flow-recon-finance.png`](flow-recon-finance.png) — PNG 1800×1200.
- [`flow-recon-finance.jpg`](flow-recon-finance.jpg) — JPEG 1800×1200.

Diagram menerangkan upload LB/RB/recon referensi → validasi → preview/edit →
proses Polars → export XlsxWriter → sheet Summary/Records/LB/RB → pemeriksaan
Finance. Gambar ini bukan screenshot interaktif aplikasi.

## Riwayat versi ringkas

| Versi | Perubahan utama |
|---|---|
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
- Tag Git harus sama dengan versi dan diawali `v` (mis. `v1.1.6`). Tag
  menjalankan Windows build dan menerbitkan installer ke GitHub Release.
- DMG dibangun di macOS melalui `scripts/build_dmg.sh`; installer Windows
  dibuat di runner Windows melalui GitHub Actions.
- Jangan menambahkan workbook transaksi, data pelanggan, output, `.venv`,
  `build/`, atau `dist/` ke Git.
- Matching silang settlement LB/RB belum didefinisikan oleh aturan bisnis;
  jangan mengubah label matching sebelum kunci/toleransi disepakati Finance.
