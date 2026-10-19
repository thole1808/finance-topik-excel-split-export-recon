# Flow Logic — Recon Builder

Dokumen ini menjelaskan logika alur data secara terperinci dari upload hingga output,
termasuk sinkronisasi referensi, lookup fee, dan mekanisme update.

Diagram singkat empat sumber upload tersedia di `alur-recon/alur-recon.svg`,
`alur-recon/alur-recon.png`, dan `alur-recon/alur-recon.jpg`.
Flowchart lengkap sistem saat ini tersedia di `flowchart-sistem/` dalam format
SVG, PNG, dan JPG.
Peta teknologi runtime dan installer tersedia di `technology-stack/`.

---

## 1. Ringkasan alur keseluruhan

```
 Upload LB/RB  ─┐
 Upload Recon   ├────→ Validasi ──→ Preview (read-only) ──┐
 (opsional)     │                                        │
                └────────────────────────────────────────┤
                                 ↓                       │
                            Proses Data (Polars)           │
                    ┌─ Sinkronisasi Ref/Vendor Ref         │
                    ├─ Lookup Fee                         │
                    ├─ Sorting (SUCCESS > FAILED)        │
                    └─ Pemisahan jam 23:00 (For tomorrow) │
                                 ↓                       │
                           Export (XlsxWriter)            │
                                 ↓                       │
            ┌────────────────────────────────────────────┘
            ↓
    Summary │ Records │ LB │ RB  →  Pemeriksaan Finance
            ↓
    Riwayat lokal (SQLite + hash chain)
    Pemeriksaan update (GitHub Release)
```

---

## 2. Input

### 2.1 Upload wajib

| File | Sheet | Deskripsi |
|------|-------|-----------|
| **LB** (Loket Bayar) | Sheet pertama | Data transaksi dari Loket Bayap |
| **RB** (Rajabiller) | Sheet pertama | Data transaksi dari Rajabiller |
| **Recon referensi** | `Summary`, `Records` | Hanya untuk ditinjau; tidak dipakai sebagai sumber data output |

### 2.2 Upload opsional

| File | Sheet | Deskripsi |
|------|-------|-----------|
| **Backdate** | `LB`, `RB`, `RECON` | Data transaksi periode sebelumnya; digabung ke output saat Generate |

### 2.3 Mapping kolom

- **LB** (`src/recon_exporter.py`, `LB_COL_MAP`): `Tanggal`, `Operator`, `Nama`, `Idpel`,
  `Produk`, `Periode`, `Status`, `No. Reff (SN)`, `Lembar`, `Tagihan`, `Denda`, `Admin`,
  `Harga`, `Fee Loket`, `Fee Bulan`, `Total`, `Ref ID`
- **RB** (`src/recon_exporter.py`, `RB_COL_MAP`): `No`, `ID Trx`, `Tanggal`, `Kode Produk`,
  `No Pelanggan-1`–`4`, `Via`, `SN`, `Ref`, `Status`, `Nominal`

---

## 3. Validasi

1. Ekstensi file harus `.xlsx`
2. Header wajib pada sheet pertama LB/RB harus sesuai
3. Footer RB (kolom `No` kosong/non-numerik) dibuang
4. Jika Backdate diunggah, ketiga sheet (`LB`, `RB`, `RECON`) wajib ada
5. Sheet `RECON` Backdate harus memiliki kolom sesuai `RECORDS_HEADERS`

---

## 4. Preview (read-only)

- Semua tab preview bersifat **read-only** (editable=False)
- Tombol **Edit** dan **Hapus** tidak tersedia
- Fitur: pagination (100/250/500/1000/Semua), pencarian literal di semua kolom, navigasi halaman
- Tab Backdate LB/RB/RECON menampilkan masing-masing sheet dari satu workbook
- Data preview tidak dapat dimodifikasi; perubahan tidak akan masuk ke output

---

## 5. Proses Data (Polars)

### 5.1 Penggabungan periode

```
LB rows = [Backdate LB rows (_period_order=0)] + [End Date LB rows (_period_order=1)]
RB rows = [Backdate RB rows (_period_order=0)] + [End Date RB rows (_period_order=1)]
```

- Jika Backdate tidak diunggah, hanya End Date rows yang diproses
- Jika Backdate diunggah, ketiga sheet harus lengkap; jika tidak, raise ValueError

### 5.2 Records

```
records = []
if backdate_recon_df is not None:
    records += records_from_recon_sheet(backdate_recon_df)  # sorted by datetime
if recon_records_df is not None:
    records += records_from_recon_sheet(recon_records_df)
if not records:  # fallback
    records += records_from_lb_rb(lb_df, rb_df)
```

**Aturan Ref dan Vendor Ref:**
- Backdate Records: Ref dan Vendor Ref diambil dari **sheet RECON** workbook Backdate (kode asli BO BILLER)
- End Date Records: Ref dan Vendor Ref diambil dari **sheet Records** file recon referensi
- Fallback: jika tidak ada sumber RECON/Records, gunakan nilai dari LB/RB sumber
- Nilai asli LB/RB tidak pernah menimpa Ref/Vendor Ref pada output

### 5.3 Sinkronisasi Ref/Vendor Ref ke sheet LB/RB

Setelah records terbentuk, Ref dan Vendor Ref di sheet LB/RB disinkronkan dari records
menggunakan 6-level matching:

| Level | Kunci pencocokan | Ketentuan |
|-------|-----------------|-----------|
| 1 | **billing_number** (dinormalisasi) | Hapus non-digit, hapus leading zero |
| 2 | **datetime + amount** | Pasangan tepat |
| 3 | **amount-only** | Hanya jika unik |
| 4 | **datetime-only** | Hanya jika unik |
| 5 | **same-day fallback** | Tanggal sama, pilih amount terdekat |
| 6 | **fuzzy billing_number** | Edit distance ≤ 2 digit, amount terdekat sebagai tiebreaker |

Setelah sinkronisasi, kolom `ref` dan `ref_id` pada setiap baris LB/RB diganti dengan
nilai dari records yang cocok.

### 5.4 Sorting

```
rows.sort(key=(status_sort_key, datetime))
```
- `SUCCESS` (sort key 0) di atas `FAILED` (sort key 1)
- Di dalam kelompok status, urutkan berdasarkan datetime (ascending)

### 5.5 Pemisahan jam 23:00

- **End Date pukul 23:00+**: dipindahkan ke blok "For tomorrow because different time".
  Untuk periode tiga hari atau lebih, hanya transaksi di tanggal terakhir yang dipindahkan.
  - Tidak muncul di blok Current Day (kolom A dan kolom recon utama)
  - Masih ada di data sumber (kolom B)
  - Tetap ada di Records
- Untuk periode tiga hari atau lebih, tanggal sebelum tanggal terakhir tetap berada di Current Day, meskipun transaksinya pukul 23:00 atau lebih.
- **Backdate pukul 23:00+**: tetap di blok utama (tidak dipindah ke For tomorrow)

### 5.6 Lookup fee

Untuk setiap transaksi, fee ditentukan oleh prioritas:

1. **`reference_data.json`**: lookup berdasarkan `kode_produk|nominal` (RB) atau
   `produk|tagihan` (LB)
2. **`fee_tables.json`**: lookup berdasarkan kode produk jika tidak ditemukan di level 1
3. **Field sumber**: gunakan `fee_loket`, `fee_bulan`, `harga` dari LB jika tersedia

Jika tidak ada profil fee dan input tidak memiliki kolom fee, sel fee output tetap kosong
(bukan angka 0).

---

## 6. Export (XlsxWriter)

### 6.1 Sheet Summary

- 7 kategori: Internal, Rajabiller, Loket Bayar, Matched, Unmatched (masing-masing biller)
- Kolom: Category, Count, Amount, Amount Success, Amount Failed
- Formula Excel (COUNTIF/SUMIF/SUMIFS) dengan cached values

### 6.2 Sheet Records

- Kolom: Datetime, Product, Billing Number, Billing Name, Ref, Vendor Ref, Amount,
  Total Fee, Customer Fee, LST Fee, Margin Biller, Transaction Status, Vendor, Recon Result
- Urutan: Backdate Records (sorted by datetime) → End Date Records
- Warna hijau untuk SUCCESS, merah untuk FAILED pada kolom Transaction Status

### 6.3 Sheet LB / RB

- Layout mengikuti `LB_SHEET_LAYOUT` / `RB_SHEET_LAYOUT`
- Tiga blok: Current Day from bo biller recon, Current Day from bo biller apps,
  For tomorrow because different time
- Subtotal SUCCESS dan FAILED setelah kelompok transaksi
- 3 baris kosong setelah subtotal SUCCESS
- Subtotal mencakup Current Day + For tomorrow + Backdate
- Formula Excel dengan cached values

---

## 7. Pemeriksaan Finance

Sebelum memakai hasil, Finance harus memverifikasi:
1. Jumlah transaksi di keempat sheet sesuai file sumber
2. Ref dan Vendor Ref sesuai dengan BO BILLER (dari RECON/Records, bukan LB/RB asli)
3. Fee untuk kode produk/nominal baru
4. Status transaksi dan Recon Result
5. Transaksi jam 23:00+ berada di blok For tomorrow
6. Subtotal SUCCESS/FAILED dan 3 baris pemisah

---

## 8. Riwayat lokal (SQLite)

- Lokasi: `~/Library/Application Support/Recon Builder/usage-history.sqlite3` (macOS)
- Event Generate mencatat ID mulai/selesai, akun OS, nama/path dan hash SHA-256
  setiap sumber, jumlah baris, hasil, ukuran, lokasi, serta SHA-256 workbook.
- ID event, akun, waktu, nama sumber, dan hash sumber ikut dicatat sebagai
  metadata workbook.
- Event lain meliputi app_started, upload, preview/edit, update, ekspor, dan
  pembukaan hasil.
- Setiap event memiliki `previous_hash` dan `event_hash` (SHA-256 hash chain)
- Layar menampilkan IP/MAC tersamarkan; ekspor CSV mencantumkan nilai lokal
  penuh untuk event yang menyimpannya. Event versi lama hanya memiliki nilai
  tersamarkan dan tidak dapat dipulihkan menjadi nilai asli.
- Bisa diekspor ke CSV dan diperiksa integritas rantainya.
- Log tersimpan per komputer; tidak dikirim ke server karena belum ada backend.

---

## 9. Pemeriksaan & Update Otomatis

### 9.1 Alur update

```
Startup:
  → _check_updates_silently() (silent=True)
    → latest_update(current_version)
      → _open_url(RELEASE_API)  [HTTPS + SSL context + retry]
      → parse tag, find platform asset (DMG untuk macOS, EXE untuk Windows)
      → return UpdateRelease atau None

User klik "Update v...":
  → _on_update_button()
    → _check_for_updates(silent=False) jika belum cek
    → _download_and_start_update()
      → download_update(release)
        → download SHA256SUMS.txt
        → download installer
        → verify SHA-256 checksum
      → start_installer(installer_path, os.getpid())
        macOS: mount DMG → install PKG → launch app baru
        Windows: run Setup.exe dengan UAC → launch app baru
```

### 9.2 SSL & jaringan

- SSL context: certifi (bundled) → system CA (macOS/Linux) → default Python (Windows)
- Retry: 3x untuk HTTP 403/422/500/502/503/504
- Rate limit: GITHUB_TOKEN jika tersedia (5000 req/jam), unauthenticated (60 req/jam)

### 9.3 Keamanan

- SHA-256 checksum diverifikasi sebelum pemasangan
- Installer hanya dijalankan dari release GitHub yang sah
- Aplikasi harus terpasang di `/Applications` (macOS) atau `Program Files` (Windows)

---

## 10. Build & Release

### 10.1 Build lokal macOS

```bash
./scripts/build_dmg.sh          # build DMG + PKG
./scripts/build_dmg.sh --clean  # hapus build lama, lalu build
```

- Python deps dari `.venv` atau sistem
- PyInstaller bundling semua dependensi
- `codesign --force --deep --sign -` (ad-hoc)
- Output: `dist/ReconBuilder-v<versi>-<timestamp>.dmg` dan `.pkg`

### 10.2 Release GitHub Actions

Workflow file: `.github/workflows/release.yml`

Trigger: push ke `main` atau tag `v*`, atau workflow_dispatch

Job flow:
1. **build-macos**: PyInstaller → DMG + PKG → upload sebagai artifact
2. **build-windows**: PyInstaller → EXE → Inno Setup → PKG → upload sebagai artifact
3. **publish-release**:
   - Download semua artifact
   - Generate `SHA256SUMS.txt` (termasuk PNG/JPG)
   - Buat tag `v<versi>` jika push ke main (bukan tag)
   - Buat GitHub Release (jika belum ada)
   - Upload semua aset (EXE, DMG, PKG, PNG, JPG, SHA256SUMS.txt)

### 10.3 Versi

- Versi sumber tunggal: `src/version.py` (`__version__`)
- Tag Git: `v<versi>` harus cocmpan dengan `__version__`
- Workflow menolak tag yang tidak cocok dengan versi source
