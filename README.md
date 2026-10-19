# Recon Builder — LB & RB

**Versi aplikasi: 1.2.47** — catatan versi dan panduan agent:
[docs/VERSI_DAN_KONTEKS_AGENT.md](docs/VERSI_DAN_KONTEKS_AGENT.md).

Aplikasi desktop untuk membaca workbook Vendor Loket Bayar (LB) dan Rajabiller
(RB), mencocokkannya dengan data BO Biller dari file recon, lalu menghasilkan
workbook dengan sheet **LB** dan **RB**. Kolom kiri berisi BO Biller dan kolom
kanan berisi transaksi Vendor.
Layout, lebar kolom, format angka, warna, serta rumus mengacu pada
`recon-2026-10-14-v1.xlsx`.

Panduan pengguna Finance (alur, format input, isi output, dan pengecekan):
[docs/PANDUAN_FINANCE.md](docs/PANDUAN_FINANCE.md).

Gambar alur sederhana: [PNG](docs/alur-recon/alur-recon.png) ·
[JPG](docs/alur-recon/alur-recon.jpg).
Flowchart lengkap cara kerja aplikasi: [PNG](docs/flowchart-sistem/flowchart-sistem.png) ·
[JPG](docs/flowchart-sistem/flowchart-sistem.jpg).
Teknologi dan framework yang benar-benar digunakan: [PNG](docs/technology-stack/technology-stack.png) ·
[JPG](docs/technology-stack/technology-stack.jpg).
Logika alur lengkap: [docs/FLOW_LOGIC.md](docs/FLOW_LOGIC.md).

Update aplikasi tersedia melalui tombol di kanan atas jendela. Aplikasi
menampilkan versi terpasang dan versi release yang tersedia; update diunduh
dari GitHub Releases dan checksum diperiksa sebelum pemasangan.
Rencana aplikasi mobile Android/iOS ada di
[docs/ROADMAP_MOBILE.md](docs/ROADMAP_MOBILE.md).

## Menjalankan dari source

Python 3.10+ direkomendasikan.

```bash
python3 -m pip install -r requirements.txt
python3 app.py
```

Pada aplikasi desktop, upload file LB, RB, dan file recon referensi berformat
Excel `.xlsx`; tersedia juga field upload opsional untuk satu workbook
Backdate `.xlsx` dengan sheet `LB`, `RB`, dan `RECON`. Saat Backdate dipilih,
klik **Generate Recon** setelah semua file diupload. Data BO Backdate berasal
dari sheet `RECON`, data Vendor Backdate berasal dari sheet `LB`/`RB`; untuk
Enddate, data BO berasal dari sheet `Records` recon dan data Vendor berasal dari
file LB/RB. Pencocokan mengutamakan nomor referensi transaksi dan memeriksa
nomor pelanggan, tanggal, serta status. Nominal pemadanan mengambil nilai
Vendor `Total` dan membandingkannya dengan BO `Amount` (RB standar tanpa `Total`
menggunakan `Nominal`). Nominal harus cocok untuk pencocokan cadangan; jika
referensi kuat menunjukkan transaksi sama tetapi nominal berbeda, keduanya
tetap dipasangkan dan ditandai `AMOUNT MISMATCH`. Baris ambigu atau tanpa
pasangan ditandai di hasil. Baris Backdate ditempatkan sebelum Enddate. Subtotal
SUCCESS/FAILED di sheet `LB` dan `RB` menggabungkan Backdate dan End Date;
transaksi End Date pukul 23:00 ditampilkan di bagian For tomorrow.
Workbook ditulis menggunakan **Polars + XlsxWriter**. File recon input harus
memiliki sheet `Summary` dan `Records`; file Backdate harus memiliki sheet `LB`,
`RB`, dan `RECON`. Command line tetap menerima dua file LB/RB seperti contoh
di bawah.
Nilai dan jumlah transaksi dibaca ulang dari file yang dipilih setiap kali
aplikasi dijalankan; tidak ada tanggal atau jumlah baris transaksi yang
ditetapkan tetap.

Preview dapat menelusuri semua baris: pilih ukuran halaman 100/250/500/1000 atau Semua,
gunakan Sebelumnya/Berikutnya, atau cari di semua kolom seluruh file. Preview
bersifat read-only — tidak ada tombol Edit, Hapus, atau klik ganda untuk
mengubah sel; data hanya ditampilkan untuk inspeksi sebelum Generate.

Workbook Backdate ditampilkan di tab terpisah untuk sheet LB, RB, dan RECON.
Hasil dibuat secara default pada folder data Recon Builder di komputer, dengan
nama `ReconBuilder_YYYYMMDD_HHMMSS.xlsx` (ditambah pecahan detik untuk mencegah
nama ganda). Tab **File Hasil** menampilkan workbook yang sudah dibuat, membuka
file, atau menyimpan salinan ke lokasi pilihan. Tombol buka output memilih file
hasil terakhir. Aplikasi memakai akun PC/laptop yang sedang login sebagai
operator, menampilkan jam lokal dan terminal proses, serta mencatat aktivitas
ke SQLite lokal. Riwayat Generate mencatat akun OS, file dan hash sumber,
jumlah baris, hasil, lokasi, ukuran, hash SHA-256 output, serta event ID di
metadata workbook. Ekspor CSV menyertakan IP/MAC lokal penuh untuk event baru;
layar tetap menampilkan nilainya tersamarkan. Log lokal mendeteksi perubahan
pada entri yang masih ada, tetapi tidak mencegah pemilik perangkat menghapus
database atau memalsukan identitas operator; audit kuat memerlukan login
terverifikasi dan penyimpanan log terpusat.

Pada sheet LB/RB, subtotal Amount dan empat kolom fee muncul setelah kelompok
SUCCESS dan FAILED. Sheet LB mempertahankan tiga baris kosong setelah subtotal
SUCCESS; baris pemisah kosong pada sheet RB dihapus agar daftar transaksinya
rapat. Transaksi For tomorrow juga ikut dijumlahkan. Tata letak mengikuti
baris total pada file acuan: kolom F adalah Amount + Margin Biller, lalu
kolom G-K berisi Amount,
Total Fee, Customer Fee, LST Fee, dan Margin Biller.

Untuk command line:

```bash
python3 recon_exporter.py LB.xlsx RB.xlsx -o recon-output.xlsx
```

## Installer aplikasi macOS

Jalankan perintah ini di macOS setelah memasang dependencies:

```bash
./scripts/build_dmg.sh
```

Hasilnya:

- `dist/ReconBuilder-v<versi>-<YYYYMMDD_HHMMSS>.pkg`
- `dist/ReconBuilder-v<versi>-<YYYYMMDD_HHMMSS>.dmg`

DMG memuat `Recon Builder.app` langsung beserta shortcut Applications. Buka
DMG, seret app ke Applications, selesai. PKG bertimestamp juga dibuat sebagai
pilihan installer terpisah. Bundle app dan file sementara berada di `dist/.build`
selama proses packaging, lalu dibersihkan; installer final hanya berada di `dist/`.
Release publik dibuat setelah build macOS lokal dan build Windows dari workflow
manual selesai. Release Latest berisi installer Windows, DMG/PKG macOS,
flowchart PNG/JPG, dan `SHA256SUMS.txt` yang memuat hash setiap installer.

Installer menghasilkan aplikasi desktop mandiri; spreadsheet LB/RB dipilih
saat aplikasi dijalankan dan tidak dibundel bersama aplikasi. Engine Excel Polars
(`fastexcel`/Calamine) dikemas eksplisit ke dalam aplikasi, sehingga pengguna
tidak perlu memasang paket Python secara terpisah. Script build memakai
`.venv` proyek bila tersedia; interpreter lain dapat ditentukan lewat variabel
`PYTHON`.

## Installer Windows

Build Windows dijalankan manual melalui GitHub Actions workflow
`.github/workflows/release.yml`. Build macOS dilakukan lokal melalui skrip di
atas. Produk dan file sementara build ditempatkan di `dist/`; file sementara
dibersihkan setelah build.

- Workflow manual menghasilkan `dist/ReconBuilder-Setup.exe` sebagai artifact
  Actions dan mengunggahnya ke draft release untuk dirilis bersama paket macOS.
- Installer menggunakan Inno Setup dan menyertakan Polars, `fastexcel`/
  Calamine, serta file aplikasi agar pengguna tidak perlu memasang Python.

## Batas keamanan dan pemantauan

Aplikasi desktop offline tidak dapat dijamin mustahil diekstrak, disalin, atau
di-patch pada komputer yang dikuasai pemiliknya. Checksum melindungi integritas
installer saat diunduh, bukan mencegah pembongkaran program setelah terpasang.
Belum ada server lisensi, anti-tamper tepercaya, atau endpoint pemantauan
pengguna. Riwayat sekarang hanya tersimpan di komputer pengguna. Diagram
[teknologi](docs/technology-stack/technology-stack.png) menunjukkan cakupan
dan batasnya.

## Struktur source

```text
app.py                         # Launcher aplikasi desktop
recon_exporter.py              # Entry point CLI kompatibel
src/
  app.py                       # Antarmuka Tkinter
  recon_exporter.py            # Baca, transformasi Polars, ekspor XlsxWriter
  audit_log.py                 # Riwayat penggunaan SQLite dengan hash chain
  device_identity.py           # Identitas IP/MAC lokal untuk audit
  announcements.py             # Pemberitahuan suara lintas OS
  updater.py                   # Pemeriksaan, unduhan, dan pemasangan update
  version.py                   # Versi aplikasi (sumber tunggal)
  fee_tables.json              # Fee fallback untuk kode produk baru
  reference_data.json          # Nama produk dan fee dari file acuan
  ref_styles.json              # Warna, format angka, font, dan lebar kolom acuan
assets/
  recon-builder.svg            # Master ikon Recon Builder
  recon-builder.icns           # Ikon aplikasi macOS
  recon-builder.ico            # Ikon aplikasi/installer Windows
scripts/
  build_dmg.sh                 # Build .app dan .dmg di macOS (--clean untuk reset)
  macos/
    postinstall                # Skrip pasca-instal PKG
  windows/
    ReconBuilder.iss           # Konfigurasi installer Inno Setup
.github/workflows/
  release.yml                  # Build installer Windows via manual dispatch
docs/
  FLOW_LOGIC.md                # Dokumentasi lengkap logika alur data
  flowchart-sistem/            # Diagram lengkap alur aplikasi saat ini
    flowchart-sistem.png       # Flowchart PNG
    flowchart-sistem.jpg       # Flowchart JPG
  technology-stack/            # Diagram stack yang benar-benar digunakan
    technology-stack.svg       # Sumber vector diagram teknologi
    technology-stack.png       # Diagram teknologi PNG
    technology-stack.jpg       # Diagram teknologi JPG
  alur-recon/                  # Diagram sederhana empat upload ke output
    alur-recon.svg             # Sumber vector
    alur-recon.png             # Gambar PNG
    alur-recon.jpg             # Gambar JPG
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
