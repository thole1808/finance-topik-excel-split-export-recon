# Recon Builder — LB & RB

Aplikasi desktop untuk membaca workbook Loket Bayar (LB) dan Rajabiller (RB),
lalu menghasilkan file recon dengan sheet **Summary**, **Records**, **LB**, dan
**RB**. Layout, lebar kolom, format angka, warna, serta rumus mengacu pada
`recon-2026-10-14-v1.xlsx`.

Panduan pengguna Finance (alur, format input, isi output, dan pengecekan):
[docs/PANDUAN_FINANCE.md](docs/PANDUAN_FINANCE.md).

Gambar alur: [PNG](docs/flow-recon-finance.png) ·
[JPG](docs/flow-recon-finance.jpg).

## Menjalankan dari source

Python 3.10+ direkomendasikan.

```bash
python3 -m pip install -r requirements.txt
python3 app.py
```

Pilih file LB dan RB berformat Excel `.xlsx`, tentukan nama output, lalu klik
**Generate Recon**. Workbook ditulis menggunakan **Polars + XlsxWriter**.
Nilai dan jumlah transaksi dibaca ulang dari file yang dipilih setiap kali
aplikasi dijalankan; tidak ada tanggal atau jumlah baris transaksi yang
ditetapkan tetap.

Setiap tab Preview LB/RB menampilkan maksimal 100 baris pertama. Klik ganda
sebuah sel untuk mengeditnya; tekan Enter atau pindah sel untuk menerapkan,
atau Esc untuk membatalkan. Perubahan tersebut ikut masuk ke export.

Untuk command line:

```bash
python3 recon_exporter.py LB.xlsx RB.xlsx -o recon-output.xlsx
```

## Membuat aplikasi macOS dan DMG

Jalankan perintah ini di macOS setelah memasang dependencies:

```bash
./scripts/build_dmg.sh
```

Hasilnya:

- `dist/ReconBuilder.app`
- `dist/ReconBuilder.dmg`

DMG berisi aplikasi desktop mandiri; spreadsheet LB/RB dipilih saat aplikasi
dijalankan dan tidak dibundel bersama aplikasi. Engine Excel Polars
(`fastexcel`/Calamine) dikemas eksplisit ke dalam aplikasi, sehingga pengguna
DMG tidak perlu memasang paket Python secara terpisah. Script build memakai
`.venv` proyek bila tersedia; interpreter lain dapat ditentukan lewat variabel
`PYTHON`.

## Struktur source

```text
app.py                         # Launcher aplikasi desktop
recon_exporter.py              # Entry point CLI kompatibel
src/
  app.py                       # Antarmuka Tkinter
  recon_exporter.py            # Baca, transformasi Polars, ekspor XlsxWriter
  fee_tables.json              # Fee fallback untuk kode produk baru
  reference_data.json          # Nama produk dan fee dari file acuan
  ref_styles.json              # Warna, format angka, font, dan lebar kolom acuan
scripts/
  build_dmg.sh                 # Build .app dan .dmg di macOS
```

## Catatan

- Header input harus sesuai format workbook LB/RB yang disediakan.
- Aplikasi menerima file `.xlsx` saja dan membaca sheet pertama dari setiap
  file. File CSV, PDF, `.xls`, dan `.xlsm` tidak didukung.
- Baris ringkasan/footer RB non-numerik di kolom `No` otomatis dilewati.
- Fee diambil dari profil yang diekstrak dari workbook acuan per kode produk dan
  nominal; untuk kombinasi yang belum ada, aplikasi memakai `fee_tables.json`
  dan kemudian nilai fee mentah pada LB bila tersedia.
- Rumus Summary, selisih fee, status sukses, serta refund dicantumkan sebagai
  rumus Excel agar tetap bisa dihitung ulang saat dibuka.
- Klasifikasi recon saat ini menormalkan status transaksi: status sukses menjadi
  `MATCHED`, sedangkan status selain sukses menjadi `INTERNAL`. Aplikasi belum
  mencocokkan transaksi LB dengan RB berdasarkan nomor referensi; baca bagian
  batasan pada panduan sebelum menggunakan kategori Summary sebagai hasil
  pencocokan antar-biller.

## Berkas yang masuk ke Git

Workbook input/output, file `.venv`, hasil build, dan DMG dikecualikan dari Git.
Jangan commit workbook transaksi karena dapat berisi data Finance/pelanggan.
