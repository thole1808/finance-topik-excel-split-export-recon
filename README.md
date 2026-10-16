# Recon Builder — LB & RB

**Versi aplikasi: 1.1.16** — catatan versi dan panduan agent:
[docs/VERSI_DAN_KONTEKS_AGENT.md](docs/VERSI_DAN_KONTEKS_AGENT.md).

Aplikasi desktop untuk membaca workbook Loket Bayar (LB) dan Rajabiller (RB),
meninjau sheet **Summary** dan **Records** dari file recon referensi, lalu
menghasilkan recon dengan sheet **Summary**, **Records**, **LB**, dan **RB**.
Layout, lebar kolom, format angka, warna, serta rumus mengacu pada
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

Pada aplikasi desktop, upload file LB, RB, dan file recon referensi berformat
Excel `.xlsx`; tersedia juga field upload opsional untuk satu workbook
Backdate `.xlsx` dengan sheet `LB`, `RB`, dan `RECON`. Saat Backdate dipilih,
klik **Generate Recon** setelah semua file wajib diupload. Baris RECON Backdate
digabung di awal sheet `Records` berdasarkan Datetime, lalu diikuti hasil
End Date; data sumber Backdate digabung ke sheet `LB` dan `RB` yang sesuai,
dengan kelompok Backdate lebih dulu. Summary menghitung gabungan kedua periode.
Tanpa Backdate, output reguler tetap seperti sebelumnya. Workbook
ditulis menggunakan **Polars + XlsxWriter**. File recon referensi harus
memiliki sheet `Summary` dan `Records`; kedua sheet dapat ditinjau, sedangkan
Summary/Records output dihitung dari upload LB/RB dan, bila tersedia, sheet
`RECON` workbook Backdate. Command line tetap menerima dua file LB/RB seperti
contoh di bawah.
Nilai dan jumlah transaksi dibaca ulang dari file yang dipilih setiap kali
aplikasi dijalankan; tidak ada tanggal atau jumlah baris transaksi yang
ditetapkan tetap.

Preview dapat menelusuri semua baris: pilih ukuran halaman 100/250/500/1000,
gunakan Sebelumnya/Berikutnya, atau cari di semua kolom seluruh file. Pilih
sel lalu klik **Edit sel terpilih**, atau klik ganda sel untuk mengedit; Enter
atau pindah sel menerapkan, Esc membatalkan. Perubahan valid ikut masuk ke
export.

Pada sheet LB/RB, subtotal Amount dan empat kolom fee muncul setelah kelompok
SUCCESS dan FAILED. Ada tiga baris kosong setelah subtotal SUCCESS; transaksi
For tomorrow juga ikut dijumlahkan. Tata letak mengikuti baris total pada file
acuan: kolom F adalah Amount + Margin Biller, lalu kolom G-K berisi Amount,
Total Fee, Customer Fee, LST Fee, dan Margin Biller.

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

## Installer Windows

Build Windows dijalankan di runner Windows melalui GitHub Actions:

- Push ke `main` atau jalankan workflow **Windows Installer** secara manual;
  `ReconBuilder-Setup.exe` tersedia sebagai artifact Actions.
- Untuk installer rilis, push tag versi seperti `v1.1.16`. Workflow membuat
  GitHub Release berisi `ReconBuilder-Setup.exe` dan `ReconBuilder.dmg`.
- Installer menggunakan Inno Setup dan menyertakan Polars, `fastexcel`/
  Calamine, serta file aplikasi agar pengguna tidak perlu memasang Python.

## Struktur source

```text
app.py                         # Launcher aplikasi desktop
recon_exporter.py              # Entry point CLI kompatibel
src/
  app.py                       # Antarmuka Tkinter
  recon_exporter.py            # Baca, transformasi Polars, ekspor XlsxWriter
  version.py                   # Versi aplikasi (sumber tunggal)
  fee_tables.json              # Fee fallback untuk kode produk baru
  reference_data.json          # Nama produk dan fee dari file acuan
  ref_styles.json              # Warna, format angka, font, dan lebar kolom acuan
scripts/
  build_dmg.sh                 # Build .app dan .dmg di macOS
  windows/
    ReconBuilder.iss           # Konfigurasi installer Inno Setup
.github/workflows/
  windows-installer.yml        # Build Windows artifact/release
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
