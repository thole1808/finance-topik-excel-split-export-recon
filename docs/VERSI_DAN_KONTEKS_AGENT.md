# Versi aplikasi dan konteks untuk agent

**Versi saat ini: 1.2.47**

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
    [`.github/workflows/release.yml`](../.github/workflows/release.yml)
   dan [`scripts/windows/ReconBuilder.iss`](../scripts/windows/ReconBuilder.iss).

## Implementasi 1.2.0 — UI, output, dan riwayat operator

### 1.2.47 — Audit hasil, ekspor identitas penuh, dan peta teknologi

- Mengikat event Generate ke sumber, hash sumber/output, akun OS, metadata
  workbook, dan ID event untuk penelusuran file.
- Ekspor CSV menyertakan IP/MAC lokal tanpa masking untuk event baru; layar
  aplikasi tetap menampilkan versi tersamarkan.
- Mendokumentasikan stack yang benar-benar digunakan, serta batas keamanan dan
  ketiadaan backend pemantauan.

### 1.2.46 — Diagram dan kesiapan update

- Menambahkan flowchart lengkap terpisah dari diagram ringkas ALUR RECON.
- Memindahkan seluruh hasil build macOS/Windows ke `dist/`; manifest checksum
  rilis bernama `SHA256SUMS.txt`.
- Release update aplikasi harus publik, berstatus Latest, dan berisi installer
  platform serta manifest checksum yang cocok.

### 1.2.42 — Pemadanan BO Biller dengan Vendor

- Output hanya memiliki sheet `LB` dan `RB`; kolom kiri bersumber dari BO
  Biller, kolom kanan dari Vendor.
- BO Backdate dibaca dari `RECON`; Vendor Backdate dibaca dari `LB`/`RB`.
  BO Enddate dibaca dari `Records` file recon; Vendor Enddate dari upload LB/RB.
- Pemadanan satu-ke-satu mengutamakan Vendor Ref dan memvalidasi nomor
  pelanggan, nominal, tanggal, dan status. Kandidat ambigu atau berbeda tidak
  dianggap cocok; baris yang belum cocok tetap ditampilkan di sisi sumbernya.
- DMG/PKG macOS dibangun lokal; installer Windows dibangun lewat workflow.

### 1.2.43 — Dasar Amount Vendor

- Matching nominal LB dan RB membandingkan BO `Amount` dengan Vendor `Total`.
- Untuk format Rajabiller standar yang tidak memiliki kolom `Total`, gunakan
  `Nominal` sebagai fallback; bila kolom `Total` tersedia, kolom itu diutamakan.
- Nilai nominal berbeda pada pasangan yang memiliki referensi kuat ditandai
  `AMOUNT MISMATCH`; pasangan cadangan tetap mensyaratkan nominal sama.

- Preview menyediakan pilihan **Semua** selain paging; ketiga sheet Backdate
  (`LB`, `RB`, `RECON`) tampil sebagai tab tersendiri.
- Output dipilih sebagai folder; nama workbook menggunakan waktu lokal sampai
  mikrodetik. Tombol output memilih file terakhir setelah generate.
- UI dimaksimalkan saat start, memakai tema navy/teal, menampilkan jam lokal,
  label **Powered By DevOps**, dan terminal tahapan generate.
- Riwayat disimpan di SQLite lokal oleh `src/audit_log.py`. Event dicatat dalam
  rantai SHA-256, termasuk nama akun PC otomatis, user OS, host, versi, aksi,
  dan rincian. Aktivitas utama meliputi file dipilih/dimuat, sheet/tab, pencarian,
  ukuran halaman/navigasi, edit nilai, generate sukses/gagal, ekspor log, dan
  file output dibuka. UI menyediakan pemeriksaan rantai dan ekspor CSV.
- Batas keamanan: akun PC bukan verifikasi identitas orang, dan database lokal dapat
  dihapus/diganti oleh pemilik perangkat. Hash chain hanya mendeteksi perubahan
  pada catatan yang tersisa. Kebutuhan agar pengguna tidak dapat menyangkal
  aktivitas memerlukan autentikasi individu dan server audit terpusat dengan
  kontrol akses/retensi; jangan menyebut log lokal ini anti-pemalsuan.
- File database: macOS `~/Library/Application Support/Recon Builder/`, Windows
  `%APPDATA%/Recon Builder/`, Linux `$XDG_DATA_HOME/recon-builder/` atau
  `~/.local/share/recon-builder/`.
- Tidak ada penghapusan source berdasarkan tebakan. Launcher root, exporter CLI,
  workflow, dan installer masih dirujuk/digunakan. Workbook, environment, dan
  artefak lokal dipertahankan.

### Catatan build dan publish 1.2.0

Perbarui versi di `src/version.py`, build aplikasi macOS melalui
`scripts/build_dmg.sh`, dan push commit source/docs ke `main`. Windows installer
dibangun oleh GitHub Actions saat push. Tag release perlu dibuat terpisah jika
pengguna meminta penerbitan release.

## Perubahan 1.2.1 — nama artefak DMG

- DMG memakai pola `ReconBuilder-v<versi>-<YYYYMMDD_HHMMSS>.dmg`, dengan waktu
  build lokal 24 jam. Contoh: `ReconBuilder-v1.2.1-20261017_234755.dmg`.
- Script menyimpan hasil di `dist/` dan menyalin artefak bernama sama ke root
  untuk workflow release. GitHub Actions memilih DMG untuk versi source yang
  sama saat membuat release.

## Implementasi 1.2.2 — release otomatis, suara, updater, dan ikon

- Push ke `main` dengan versi baru membangun Windows installer, memvalidasi
  kecocokan tag dengan `src/version.py`, membuat tag `v<versi>` dan GitHub
  Release, lalu melampirkan installer, DMG bertimestamp, diagram PNG/JPG, dan
  `SHA256SUMS.txt`. Push tag manual yang versinya cocok memakai alur release
  yang sama. Push lain dengan versi yang release-nya sudah ada tidak membuat
  release duplikat.
- Aplikasi mengucapkan “Mohon tunggu. Proses generate file recon sedang
  berjalan.” saat Generate dimulai, dan “Berhasil. Generate file recon dari BO
  Biller selesai.” hanya setelah workbook berhasil dibuat. OS memakai suara
  bawaan; pesan diantrekan agar tidak tumpang tindih.
- Tombol **Periksa Update** memeriksa GitHub Releases saat aplikasi dibuka dan
  menampilkan label versi baru. Kliknya mengunduh installer OS, memeriksa
  SHA-256, lalu menjalankan updater. macOS mengganti `.app` setelah proses lama
  ditutup; Windows menjalankan Setup dan menutup aplikasi agar installer dapat
  mengganti file. macOS meminta izin administrator bila aplikasi dipasang di
  folder sistem. Tidak ada updater otomatis untuk Linux.
- Aset ikon master `assets/recon-builder.svg` menghasilkan PNG, ICNS, dan ICO.
  Ikon dibundel di aplikasi/installer; update menggantinya bersama versi baru.
- Batas update: harus ada release GitHub yang sukses dan artefak dengan checksum.
  Bila jaringan tidak tersedia, aplikasi tetap dapat dipakai dan tombol bisa
  dicoba kembali. SHA-256 mendeteksi kerusakan/perubahan berkas terhadap
  manifest release; ia bukan tanda tangan publisher terpisah.

## Implementasi 1.2.3 — update terlihat, UI dirapikan, identitas perangkat dimasking

- Tombol update berada di kanan atas header agar mudah ditemukan. Header
  membandingkan `APP v...` dengan versi GitHub `RELEASE v...`; saat ada versi
  baru tombol menampilkan **Update v...**. Keduanya memakai satu sumber versi:
  `src/version.py` dan tag `v<versi>`.
- Perapian desktop: header/panel lebih tegas, jarak dan hierarki kontrol
  diperjelas, status release berada dekat tombol update, dan tabel riwayat punya
  kolom IP/MAC khusus.
- Setiap event riwayat SQLite menyimpan IP lokal bertopeng (tiga oktet awal)
  dan MAC bertopeng (tiga oktet awal). Nilai mentah tidak ditulis ke database
  atau ekspor CSV. Alamat dibaca lokal; aplikasi tidak meminta layanan IP luar.
- File `docs/ROADMAP_MOBILE.md` menjelaskan rencana porting Android/iOS. Tkinter
  dan Python Polars tidak dapat langsung dibangun sebagai APK/IPA.

## Implementasi 1.2.4 — waktu Indonesia, suara Indonesia, dan layar responsif

- Jam header memakai waktu lokal/zona waktu komputer. Nama hari dan bulan
  ditampilkan dalam Bahasa Indonesia; nama file dan log tetap memakai waktu
  lokal mesin.
- macOS memilih voice `id_ID` jika tersedia. Windows memilih voice terpasang
  berbahasa Indonesia jika tersedia; jika tidak, gunakan voice sistem. Tombol
  **Tes Suara** di footer memudahkan operator memeriksa speaker sebelum Generate.
- Jendela dimaksimalkan saat start dengan minimum 960×660. Header disusun ulang
  agar tetap muat pada laptop, sementara preview mengambil ruang sisa dan tabel
  memiliki scrollbar; layar lebar menggunakan ruang kerja yang lebih besar.

## Implementasi 1.2.5 — akun PC, restart update, dan waktu riwayat lokal

- Prompt operator dihilangkan; akun OS yang sedang login otomatis menjadi nama
  operator.
- Helper update macOS/Windows menunggu proses lama berakhir, memasang release,
  lalu membuka kembali aplikasi yang baru. Setup Windows berjalan senyap setelah
  konfirmasi operator dan meminta UAC bila perlu; jika UAC dibatalkan, versi
  yang saat ini terpasang dibuka kembali.
- Kolom waktu Riwayat dan CSV diformat memakai zona waktu PC serta nama hari dan
  bulan Indonesia. Timestamp ISO yang di-hash di SQLite tidak diubah, sehingga
  pemeriksaan integritas tetap valid.
- Tombol **Tes Suara** hanya tampil sebelum penggunaan pertama akun lokal, lalu
  disembunyikan. Kalimat lebih percakapan dan voice id-ID dipilih bila tersedia.
  Voice sistem tetap sintesis suara; rekaman penutur manusia harus disediakan
  sebagai file audio.

## Implementasi 1.2.6 — watermark batik dan installer OS

- Latar aplikasi memakai ornamen batik transparan yang ditaruh di tepi kanan
  bawah sebagai watermark dekoratif; panel dan teks berada di lapisan depan.
- macOS build menghasilkan PKG yang memasang `Recon Builder.app` ke
  `/Applications`, dibungkus di DMG bertimestamp agar mudah diunduh/dipasang.
- Windows Inno Setup memasang aplikasi ke Program Files, membuat shortcut Start
  Menu, dan memilih shortcut Desktop secara default. Installer menampilkan nama
  **Recon Builder** di daftar aplikasi sistem.
- PNG watermark dibundel ke aplikasi macOS/Windows. Build release menyertakan
  installer PKG, DMG, dan installer Windows dengan manifest SHA-256.

## Implementasi 1.2.7 — motif batik Indonesia

- Watermark generik diganti dengan stilisasi motif Kawung dan aksen alur
  diagonal Parang dalam SVG/PNG transparan. Motif tetap berada di latar
  kanan bawah dan dibuat samar agar tidak menutupi teks maupun tabel.

## Implementasi 1.2.8 — motif batik floral Indonesia

- Watermark menggunakan aset batik floral Indonesia yang tersedia di
  `assets/batik.png`, diolah menjadi `assets/batik-watermark.png` transparan
  berwarna teal. Pola ditempatkan kecil di kanan bawah agar tidak menghalangi
  teks atau tabel, dan dibundel pada installer Mac/Windows.

## Implementasi 1.2.9 — Launchpad dan file hasil tersimpan

- macOS PKG mengatur bundle ID `com.devops.reconbuilder` dan postinstall
  mendaftarkan app lewat Launch Services serta menyegarkan Dock/Launchpad.
- Output Generate baru disimpan di folder **Generated Files** pada data lokal
  aplikasi. Tab **File Hasil** menampilkan hasil yang masih ada, membuka workbook,
  atau menyalin workbook ke lokasi pilihan. Log SQLite tetap menjadi indeks
  daftar file; workbook tidak dimasukkan ke database.

## Implementasi 1.2.10 — pemasangan Mac eksplisit dan shortcut Desktop

- macOS PKG memakai payload root eksplisit agar app terpasang ke
  `/Applications/Recon Builder.app`, tanpa bergantung pada component bundle
  placement. Postinstall membuat shortcut symlink di Desktop pengguna aktif,
  mendaftarkan Launch Services, dan menyegarkan Dock.
- Updater macOS menerima DMG yang berisi PKG: setelah verifikasi checksum dan
  persetujuan admin, memasang PKG lalu membuka lagi app terpasang.

## Implementasi 1.2.11 — nama bundle Mac dan visibilitas pasca-instal

- Bundle macOS kini dibangun langsung sebagai `Recon Builder.app` dengan nama
  executable dan nama yang tampil konsisten.
- Postinstall mencatat hasil pemasangan, mendaftarkan app, membuat shortcut
  Desktop untuk pengguna aktif, lalu membuka Finder dengan app di-highlight
  pada `/Applications` agar lokasi pemasangan terlihat setelah installer selesai.

## Implementasi 1.2.12 — DMG drag-to-Applications

- DMG memuat bundle `Recon Builder.app` secara langsung, alias `Applications`,
  serta instruksi instalasi; operator dapat menyeret app tanpa menjalankan PKG.
- PKG tetap dibangun sebagai pilihan terpisah dan diunggah bersama DMG.

## Implementasi 1.2.13 — instalasi PKG selalu ke Applications

- Metadata component PKG sebelumnya menandai bundle sebagai relocatable, sehingga
  Installer dapat mengikuti salinan bundle lama yang ditemukan pada lokasi lain.
  Build baru mengatur `BundleIsRelocatable=false` agar payload hanya dipasang ke
  `/Applications/Recon Builder.app`.
- Postinstall sekarang mengembalikan status gagal jika app atau executable tidak
  ada di lokasi tersebut, sehingga Installer tidak menyatakan sukses palsu.
- Instruksi DMG dipersingkat menjadi: buka DMG, seret app ke Applications.

## Implementasi 1.2.14 — permission instruksi DMG

- File `CARA INSTALL.txt` diberi permission baca umum (`0644`) dan image dibuat
  tanpa mempertahankan ownership dari mesin build, agar pengguna dapat membuka
  petunjuk di dalam DMG.

## Implementasi 1.2.15 — terminal proses lebih besar

- Panel Terminal proses Generate ditambah dari 4 menjadi 9 baris dan font
  diperbesar dari 9 menjadi 10 pt. Panel tetap mengisi lebar footer aplikasi.

## Implementasi 1.2.16 — watermark batik terlihat di header

- Watermark sebelumnya ditempatkan di root window lalu diturunkan ke belakang
  panel konten yang menutupi gambar. Kini watermark berada di area kosong header,
  di bawah teks dan kontrol, dengan ukuran yang pas untuk tetap tampak samar.

## Serah-terima sesi — 16 Oktober 2026 (sesi 2)

### Perbaikan workflow: macOS DMG/PKG dibangun di CI dan release.yml

**Masalah:** `updater.py` sudah lengkap (checksum SHA256, download, install DMG/PKG), tapi
workflow lama hanya punya Windows runner dan nama file workflow belum mencerminkan
cakupan penuh. DMG dibangun lokal di macOS developer dan tidak pernah ter-upload ke
GitHub Release. Akibatnya `latest_update()` melempar
`UpdateError: "Rilis vX.X.X belum menyediakan updater untuk Darwin"`.

**Solusi yang diterapkan** di `.github/workflows/release.yml`:
- Job `build-macos` yang jalan di `macos-latest` runner, menjalankan `scripts/build_dmg.sh`
  (dengan `--clean` opsional), lalu upload DMG dan PKG sebagai artifact.
- Job `build-windows` yang jalan di `windows-latest`, membangun EXE dengan PyInstaller
  lalu membuat installer dengan Inno Setup, upload sebagai artifact.
- `concurrency` group mencegah run paralel dari workflow yang sama.
- `publish-release` job dengan `needs: [build-macos, build-windows]` menunggu kedua
  platform, download semua artifact, generate `SHA256SUMS.txt` dari **semua platform
  sekaligus** (EXE + DMG + PKG + PNG + JPG), lalu upload ke GitHub Release.
- Pada push ke `main`, workflow membuat tag Git `v<versi>` secara eksplisit sebelum
  membuat release. Pada tag push, validasi versi tag sesuai `src/version.py`.
- Urutan: macOS build → Windows build → download artifact → SHA256SUMS.txt final →
  tag (jika belum ada) → release (jika belum ada) → upload semua aset.

**Aturan ke depan:**
- Cukup push tag `vX.X.X` ke `main` — workflow otomatis build kedua platform dan publish
  release lengkap dengan `SHA256SUMS.txt`.
- Atau push ke `main` dengan versi baru di `src/version.py` — workflow otomatis membuat
  tag dan release.
- Tidak perlu lagi build DMG lokal untuk release; build lokal (`scripts/build_dmg.sh`)
  tetap bisa dipakai untuk testing sebelum push, termasuk dengan flag `--clean`.
- `SHA256SUMS.txt` selalu mencakup EXE + DMG + PKG sehingga `updater.py` bisa
  verifikasi checksum di semua platform.
- PNG/JPG flowchart berada di `docs/` dan disertakan sebagai aset release; SVG sumber
  juga tersedia. Lihat `docs/FLOW_LOGIC.md` untuk logika alur lengkap.

---

## Serah-terima sesi — 16 Oktober 2026

Gunakan catatan ini saat melanjutkan pekerjaan pada sesi atau agent baru.

### Backdate — keputusan tetap dan status implementasi

- Jangan ubah alur upload End Date/reguler yang sudah berfungsi. Tambahkan
  tepat **satu** field upload **File Backdate** pada form yang sama.
- File Backdate adalah satu workbook Excel `.xlsx`, bukan tiga upload terpisah.
  File yang ditetapkan pengguna untuk dipakai adalah
  `Backdate BILLER-new-14.xlsx` (tersedia lokal di folder Downloads per
  16 Oktober 2026). Workbook telah diperiksa: `LB` 158 baris, `RB` kosong
  (header tetap ada), dan `RECON` 158 baris.
- Pengguna mengonfirmasi: unggah semua file End Date/reguler yang diwajibkan
  serta satu file Backdate, lalu tekan **Generate**. Hanya saat itu kedua
  periode diproses dan digabung dalam satu workbook.
- Keputusan output final yang dikonfirmasi pengguna: tanpa Backdate, output
  reguler tetap persis seperti sebelumnya. Jika Backdate disediakan, gabungkan
  data ke sheet yang sama (`Records`, `LB`, `RB`), bukan sheet Backdate baru.
  Baris Backdate pada `Records` dibuat dari sumber LB/RB dan ditempatkan paling
  awal berdasarkan Datetime, lalu diikuti baris End Date. Ref dan Vendor Ref
  tidak boleh diganti nilai dari sheet RECON Backdate. Di sheet `LB`/`RB`,
  tampilkan periode Backdate sebelum End Date di setiap grup status dan hitung
  satu subtotal gabungan untuk setiap status.
- Sheet `RECON` Backdate tetap divalidasi, tetapi data transaksi dan referensi
  pada output berasal dari workbook LB/RB.
- Bila File Backdate tidak diunggah, aplikasi harus tetap menghasilkan output
  reguler seperti sekarang.
- **Versi 1.1.14:** field upload Backdate ditambahkan, belum diproses.
- **Versi 1.1.15:** pemrosesan opsional ditambahkan. Validasi `LB`, `RB`,
  `RECON` dan header wajib saat Generate; gabungkan periode. Output tanpa
  Backdate tetap pada empat sheet seperti sebelumnya.
- **Versi 1.1.16:** perbaiki penempatan transaksi Backdate pukul 23:00 atau
  lebih agar tampil di blok utama pada sheet LB/RB dan tidak terlihat kosong
  di sisi Current Day. Aturan For tomorrow tetap berlaku pada transaksi End
  Date reguler. Workbook Backdate yang diuji memiliki 158 baris LB, sheet RB
  kosong, dan 158 baris RECON (semuanya Vendor Loket Bayar); jadi tidak ada
  baris sumber RB Backdate untuk ditambahkan.

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

### Implementasi terbaru — versi 1.1.17

- Subtotal SUCCESS dan FAILED di sheet LB/RB mencakup Backdate dan End Date
  sekaligus; tidak ada subtotal terpisah per periode.
- Sheet Records mengambil Ref dan Vendor Ref Backdate langsung dari sumber
  LB/RB. Sheet RECON Backdate tetap divalidasi, tetapi nilainya tidak menimpa
  referensi sumber.
- End Date pukul 23:00 tetap ditempatkan di blok For tomorrow dan Records.
  Pada output yang diperiksa, 15 Oktober 2026 memiliki 187 transaksi pukul
  23:xx di Records (186 LB, 1 RB); LB memuat 186 pada blok For tomorrow.

### Koreksi tabel For tomorrow LB — versi 1.1.20

- Header template recon For tomorrow LB berada di kolom AN:BA, data sumber
  berada di BD:BT. Writer harus mengisi 14 kolom recon lengkap di AN:BA.
- Uji export memverifikasi timestamp End Date pukul 23, seluruh kolom tabel
  recon, status, vendor, dan Recon Result.

### Posisi baris For tomorrow LB — versi 1.1.21

- Baris End Date pukul 23:00+ ditulis mulai baris 4 di blok For tomorrow,
  terlepas dari banyaknya baris Current Day/Backdate.
- Baris Current Day dan For tomorrow dijadwalkan bersama per status agar
  workbook streaming tidak kehilangan penulisan ke baris yang sudah lewat.
- Subtotal status memakai SUMIF pada status For tomorrow dan tetap mencakup
  Backdate serta End Date.

### Perataan baris For tomorrow LB — versi 1.1.22

- Semua baris End Date pukul 23:00+ pada blok For tomorrow LB dimulai dari
  baris 4 dan berlanjut rapat, dikelompokkan SUCCESS lalu FAILED/OTHER.
- Writer streaming menulis tabel For tomorrow bersama baris Current Day dalam
  urutan baris Excel, sehingga tidak mencoba kembali menulis baris yang telah
  di-flush.

### Koreksi aturan jam 23 — versi 1.1.19

- Pengguna mengonfirmasi transaksi End Date pukul 23:00 harus masuk ke blok
  **For tomorrow because different time**. Backdate pukul 23:00 atau lebih
  tetap berada di blok utama.
- Pemisahan timestamp >=23:00 kembali aktif untuk baris End Date pada LB/RB;
  baris Backdate dikecualikan dari pemisahan.
- Uji regresi harus mencakup output dengan dan tanpa Backdate, serta memastikan
  timestamp 23:xx ditulis ke kolom Current Day baik di data sumber maupun
  tabel recon.

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
- Commit `986b79c` (`Keep Backdate values visible in LB output`) sudah dipush
  ke `main` dan diberi tag `v1.1.16`, berurutan setelah `v1.1.15`.
- GitHub Actions sedang membangun installer Windows dan menerbitkan release
  `v1.1.16`: [workflow run](https://github.com/thole1808/finance-topik-excel-split-export-recon/actions/runs/38895445409).
- DMG 1.1.16 sudah dibangun di `dist/ReconBuilder.dmg` dan disalin ke root
  `ReconBuilder.dmg`. Versi bundle, signature, checksum, dan kesamaan kedua
  salinan sudah diverifikasi.
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

## Perbaikan versi 1.1.16

- Transaksi Backdate pada pukul 23:00 atau lebih kini tetap berada di blok
  utama pada sheet LB/RB, supaya dapat dilihat bersama urutan Backdate sebelum
  End Date. Pemisahan For tomorrow tetap berjalan untuk End Date reguler.
- Memastikan hasil RB tidak mengarang baris Backdate: file sumber
  `Backdate BILLER-new-14.xlsx` memiliki sheet RB kosong. Transaksi tanggal
  15 dari file End Date tetap berasal dari input End Date.

## Perubahan versi 1.1.15

- Membaca dan memvalidasi sheet `LB`, `RB`, dan `RECON` dari satu workbook
  Backdate opsional ketika pengguna menekan Generate.
- Menggabungkan `RECON` Backdate di awal `Records` dalam urutan Datetime,
  lalu mengikuti data End Date. Menggabungkan data LB/RB sumber di sheet
  masing-masing dengan periode Backdate lebih dahulu dan subtotal sesuai
  periode serta status.
- Menghitung Summary dari transaksi Backdate dan End Date bersama-sama.
- Mempertahankan output empat sheet reguler jika Backdate tidak dipilih.
- Validasi runtime pada workbook `Backdate BILLER-new-14.xlsx`: LB=158,
  RB=0, RECON=158; uji workbook sementara dengan dua baris End Date
  memastikan semua 158 baris Backdate berada lebih dulu berdasarkan Datetime,
  Summary mencakup 160 transaksi, serta output reguler tanpa Backdate tetap
  berisi empat sheet. Perbaikan penempatan data Backdate di blok LB/RB dicatat
  sebagai 1.1.16.

## Perubahan versi 1.1.14

- Menambahkan satu field upload opsional **File Backdate** untuk workbook
  `.xlsx` pada form desktop yang sama dengan alur End Date.
- Pemrosesan workbook Backdate belum diaktifkan.
- Memperbarui panduan Finance dan mencatat ulang keputusan Backdate yang sudah
  disepakati agar dapat dilanjutkan pada sesi berikutnya.

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

## Diagram alur

Diagram yang digunakan saat ini tersedia terpisah: ringkasan di
`docs/alur-recon/`, flowchart cara kerja sistem di `docs/flowchart-sistem/`,
serta stack teknologi aktual di `docs/technology-stack/` (SVG, PNG, dan JPG).

## Riwayat versi ringkas

| 1.2.52 | Aturan tanggal terakhir aktif hanya untuk periode End Date dengan 3 hari atau lebih; periode 1–2 hari tetap memindahkan seluruh transaksi pukul 23:00+. |
| 1.2.51 | Untuk input multi-hari, transaksi jam 23:00+ hanya masuk For tomorrow pada tanggal terakhir End Date; tanggal sebelumnya tetap Current Day. |
| 1.2.50 | Tambah tab Download Aplikasi dengan tautan installer macOS DMG dan Windows untuk laptop/PC baru. |
| 1.2.49 | Rapatkan baris RB yang benar-benar kosong sambil menjaga data Vendor RB; ganti nama aplikasi menjadi Auto Recon Biller dan nama output `Recon - <tanggal>`. |
| 1.2.48 | Hapus baris pemisah kosong di sheet RB saja; log anomali menyebut nama upload dan periode, dengan pemberitahuan suara untuk operator. |
| 1.2.47 | Audit Generate terikat hash input/output dan metadata workbook; CSV memakai IP/MAC lokal penuh; diagram stack teknologi aktual. |
| 1.2.43 | Bandingkan BO Amount dengan Vendor Total untuk LB/RB; RB format lama fallback ke Nominal; rebuild lokal dan rilis publik dengan checksum. |
| 1.2.42 | Pemadanan BO Biller/Vendor satu-ke-satu; output hanya LB/RB; flow diagram baru; build macOS lokal, workflow Windows manual, release publik dengan checksum. |
| 1.2.38 | Workflow release diperbaiki (release.yml, concurrency, tag Git eksplisit); build_dmg.sh mendukung --clean dan verifikasi dependensi; bug page_size.bind() diperbaiki; tombol Hapus di preview dihilangkan; updater timeout 30s/5x retry dengan pesan error yang ramah; flow diagram diperbarui ke 1.2.38. |
| 1.2.37 | Preview data read-only: tampilan, pencarian, dan navigasi halaman hanya untuk melihat; tombol Edit, klik ganda, dan Enter tidak lagi menerapkan perubahan; sel tidak bisa dihapus/diedit. |
| 1.2.36 | Sinkronkan juga Customer Fee, LST Fee, dan Margin Biller (nilai uang) dari Sheet Records final ke sheet LB/RB output; `_lb_values`/`_rb_values` prioritas nilai sinkron, fallback hitung ulang. Ref/Vendor Ref tetap 6 level pencocokan. |
| 1.2.35 | Perkuat normalisasi billing_number (hapus non-digit) dan tambah level 6 fuzzy BN (beda 1-2 digit) agar sisa baris LB/RB tersinkron dengan Ref dari Records final. |
| 1.2.34 | Tambah level 5 pencocokan same-day fallback (pilih amount terdekat di tanggal yang sama) agar sisa baris LB/RB yang BN-nya tidak ada di Records tetap tersinkron dengan Ref dari Records. |
| 1.2.33 | Tambah level 4 pencocokan datetime-only (jika unik) untuk sinkronisasi Ref/Vendor Ref di sheet LB/RB; 4 level total (BN → datetime+amount → amount-only → datetime-only). |
| 1.2.32 | Tambah fallback pencocokan (datetime, amount) dan amount-only untuk sinkronisasi Ref/Vendor Ref di sheet LB/RB; handle float/int billing_number. Memperbaiki 252 baris LB yang BN-nya tidak match di Records. |
| 1.2.31 | Kunci pencocokan sinkronisasi Ref/Vendor Ref di sheet LB/RB diganti ke billing_number saja (dengan normalisasi leading zero); amount hanya tiebreaker jika ada duplikat. Memperbaiki amount mismatch (tagihan LB ≠ Amount Records) yang menyebabkan 799 baris LB tidak tersinkron. |
| 1.2.30 | Kunci pencocokan sinkronisasi Ref/Vendor Ref di sheet LB/RB diganti dari (datetime, amount) ke (billing_number, amount) agar lebih unik dan stabil; semua baris LB/RB dijamin Ref-nya dari Sheet Records/RECON final. |
| 1.2.29 | Ref dan Vendor Ref di sheet LB/RB output juga disinkronkan dari Sheet RECON (backdate) dan Sheet Records (upload recon); nilai mentah dari sumber LB/RB diganti berdasarkan pencocokan (datetime, amount). |
| 1.2.28 | Ref dan Vendor Ref di sheet Records output diambil HANYA dari Sheet RECON (backdate) dan Sheet Records (upload recon); fallback ke LB/RB hanya jika kedua sumber tersebut TIDAK ADA sama sekali. |
| 1.2.27 | Ref dan Vendor Ref di sheet Records output diambil dari Sheet RECON (backdate) dan Sheet Records (upload recon) sebagai prioritas utama; fallback ke LB/RB hanya jika keduanya tidak tersedia. |
| 1.2.26 | Ref dan Vendor Ref di output diambil dari Sheet RECON (backdate) dan Sheet Records (upload recon), bukan dari LB/RB. Data RECON tidak diubah oleh aplikasi. |
| 1.2.25 | Perlindungan Ref dan Vendor Ref dari BO BILLER: nilai di sheet RECON (backdate dan file uploads End Date) tidak diubah oleh aplikasi; dipertahankan asli dari sumber kode BO BILLER. |
| 1.2.24 | Perbaiki pemeriksaan update: sertakan token GitHub agar tidak terkena HTTP 403 rate-limit, ulangi otomatis, dan beri pesan error yang jelas beserta waktu coba lagi. |
| 1.2.16 | Watermark batik dipindah dari belakang panel konten ke ruang kosong header agar tampil samar tanpa menutupi teks atau kontrol. |
| 1.2.15 | Panel Terminal proses Generate diperbesar menjadi 9 baris dan font 10 pt, tetap responsif selebar jendela. |
| 1.2.14 | Permission `CARA INSTALL.txt` dibuat terbaca untuk semua user dan ownership mesin build tidak dibawa ke DMG. |
| 1.2.13 | PKG tidak relocatable dan gagal secara nyata bila bundle tidak ada di `/Applications`; instruksi DMG dibuat singkat. |
| 1.2.12 | DMG memuat app langsung dan alias Applications agar pemasangan Mac dapat dilakukan dengan drag-and-drop; PKG tetap tersedia terpisah. |
| 1.2.11 | Nama bundle dan executable Mac konsisten; postinstall mencatat hasil dan menampilkan app di `/Applications` setelah pemasangan. |
| 1.2.10 | PKG menyalin app eksplisit ke Applications, membuat shortcut Desktop dan mendaftarkan Launchpad; updater memasang PKG. |
| 1.2.9 | Mendaftarkan app Mac ke Launchpad; tab File Hasil, penyimpanan default lokal aplikasi, dan simpan salinan. |
| 1.2.8 | Watermark menggunakan pola batik floral Indonesia transparan. |
| 1.2.7 | Watermark motif Kawung dan Parang khas batik Indonesia, transparan. |
| 1.2.6 | Installer macOS ke Applications dan shortcut Start Menu/Desktop Windows. |
| 1.2.5 | Akun PC otomatis, updater menutup/memasang/membuka kembali, waktu riwayat lokal Indonesia, tes suara sekali. |
| 1.2.4 | Waktu Bahasa Indonesia sesuai zona lokal, voice Indonesia, tombol tes suara, layout untuk laptop/layar lebar. |
| 1.2.3 | Tombol update di header, label versi APP/RELEASE, UI dirapikan, IP/MAC masked di riwayat, rencana mobile. |
| 1.2.2 | Tag/release otomatis, suara status, updater checksum, ikon baru, flowchart PNG/JPG di release. |
| 1.2.1 | Nama DMG mencantumkan versi aplikasi dan waktu build. |
| 1.2.0 | Tab tiap sheet Backdate, preview Semua, output bertimestamp, UI maximize/clock/terminal, dan audit SQLite lokal. |
| 1.1.22 | Tulis semua baris For tomorrow LB berurutan mulai dari baris 4. |
| 1.1.21 | Mulai tabel For tomorrow LB tepat di bawah header meski ada banyak baris Current Day/Backdate. |
| 1.1.20 | Lengkapi tabel recon For tomorrow LB di AN:BA untuk transaksi End Date jam 23. |
| 1.1.19 | Kembalikan transaksi End Date jam 23 ke For tomorrow; Backdate tetap di blok utama. |
| 1.1.18 | Percobaan menampilkan transaksi jam 23 End Date di blok Current Day. |
| 1.1.17 | Gabungkan subtotal Backdate+End Date per status dan pertahankan Ref/Vendor Ref Backdate dari sumber LB/RB. |
| 1.1.16 | Tampilkan transaksi Backdate pukul 23:00+ di blok utama LB/RB; transaksi End Date tetap mengikuti aturan For tomorrow. |
| 1.1.15 | Menggabungkan Backdate opsional ke Summary/Records/LB/RB, dengan transaksi Backdate lebih dahulu; tanpa Backdate output reguler tetap sama. |
| 1.1.14 | Menambahkan field upload Backdate opsional pada form; pemrosesan Backdate belum aktif. |
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
- Tag Git harus sama dengan versi dan diawali `v` (mis. `v1.2.42`). Buat
  release secara manual setelah DMG/PKG lokal dan installer Windows dari
  workflow tersedia; upload checksum yang sesuai dan pastikan release publik.
- DMG dibangun di macOS melalui `scripts/build_dmg.sh`; installer Windows
  dibuat di runner Windows melalui GitHub Actions `workflow_dispatch` saja.
- Simpan manifest `SHA256SUMS.txt` lokal di `dist/` dan sertakan manifest yang
  sama sebagai asset release bersama installer.
- Bundle desktop lokal dapat diekstrak atau dipatch pemilik/admin komputer;
  checksum installer tidak mencegahnya. Jangan klaim proteksi anti-crack tanpa
  lisensi/backend dan validasi sisi server.
- Pemantauan lintas komputer belum tersedia sebelum backend autentikasi dan
  telemetry yang transparan disediakan.
- Jangan menambahkan workbook transaksi, data pelanggan, output, `.venv`,
  `build/`, atau `dist/` ke Git.
- Matching silang settlement LB/RB belum didefinisikan oleh aturan bisnis;
  jangan mengubah label matching sebelum kunci/toleransi disepakati Finance.
