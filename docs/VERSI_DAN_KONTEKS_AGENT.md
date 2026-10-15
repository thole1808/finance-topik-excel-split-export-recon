# Versi aplikasi dan konteks untuk agent

**Versi saat ini: 1.1.0**

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

## Gambar yang tersedia

Gambar berikut adalah diagram alur aplikasi yang sama, disediakan untuk
kebutuhan dokumentasi/presentasi:

- [`flow-recon-finance.png`](flow-recon-finance.png) — PNG 1800×1080.
- [`flow-recon-finance.jpg`](flow-recon-finance.jpg) — JPEG 1800×1080.

Diagram menerangkan input LB/RB → validasi → preview/edit → proses Polars →
export XlsxWriter → sheet Summary/Records/LB/RB → pemeriksaan Finance. Gambar
ini bukan screenshot interaktif aplikasi.

## Riwayat versi ringkas

| Versi | Perubahan utama |
|---|---|
| 1.1.0 | Preview/pencarian seluruh data, highlight status hijau, versioning, dan installer Windows otomatis. |
| 1.0.0 | Export workbook LB/RB menggunakan Polars dan XlsxWriter; UI dan build macOS. |

## Aturan rilis dan data

- Ubah `__version__` di `src/version.py` pada setiap rilis; gunakan format
  semantic versioning `MAJOR.MINOR.PATCH`.
- Tag Git harus sama dengan versi dan diawali `v` (mis. `v1.1.0`). Tag
  menjalankan Windows build dan menerbitkan installer ke GitHub Release.
- DMG dibangun di macOS melalui `scripts/build_dmg.sh`; installer Windows
  dibuat di runner Windows melalui GitHub Actions.
- Jangan menambahkan workbook transaksi, data pelanggan, output, `.venv`,
  `build/`, atau `dist/` ke Git.
- Matching silang settlement LB/RB belum didefinisikan oleh aturan bisnis;
  jangan mengubah label matching sebelum kunci/toleransi disepakati Finance.
