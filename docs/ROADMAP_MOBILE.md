# Rencana Recon Builder untuk Android dan iOS

## Bisa langsung dibuild menjadi APK atau iOS?

Belum. Aplikasi saat ini memakai Tkinter untuk desktop dan Python Polars untuk
memproses workbook. Tkinter tidak berjalan sebagai UI Android/iOS, dan build
PyInstaller macOS/Windows tidak menghasilkan APK atau IPA. Mobile memerlukan
antarmuka dan integrasi file yang baru, serta pengujian ulang hasil Excel.

## Rekomendasi implementasi

Gunakan Flutter untuk satu UI Android/iOS, dengan pemrosesan workbook tetap di
perangkat. Jangan unggah file transaksi Finance ke server. Aturan rekonsiliasi
harus dipindahkan bertahap ke core library yang dapat dipakai mobile, lalu
dibandingkan dengan hasil exporter desktop menggunakan workbook uji Finance.

## Tahapan

1. **Kunci spesifikasi dan contoh hasil.** Kumpulkan variasi workbook LB, RB,
   Backdate, workbook besar, serta hasil yang telah disetujui Finance. Tentukan
   ukuran file maksimum, versi minimum Android/iOS, dan apakah output perlu
   dibagikan melalui email atau aplikasi internal.
2. **Pisahkan aturan bisnis dari UI desktop.** Pisahkan normalisasi tanggal,
   aturan `For tomorrow`, fee, Backdate, dan pembentukan Records ke core
   exporter. Buat uji kesetaraan agar workbook dengan input yang sama
   menghasilkan nilai dan urutan transaksi yang sama.
3. **Buat shell Flutter lokal/offline.** Tambahkan pemilih file, validasi
   workbook, tab preview dan paging, pencarian, koreksi sel, progress generate,
   penyimpanan lokal, dan ekspor ke folder/Files. File tetap berada di perangkat.
4. **Port dan validasi engine Excel.** Pilih parser/penulis XLSX yang mampu
   mempertahankan format, formula, filter, dan warna workbook ini. Jika library
   Dart tidak memenuhi kesetaraan, evaluasi core Rust lintas-platform melalui
   FFI. Jangan menyalin rumus exporter ke beberapa platform tanpa uji parity.
5. **Bawa kontrol profesional.** Tambahkan riwayat lokal, ekspor log, update
   aplikasi melalui saluran resmi platform, serta penanganan kegagalan dan
   pemulihan file. Android/iOS umumnya tidak memberi aplikasi akses MAC
   address perangkat; log mobile harus memakai ID instalasi acak yang dapat
   direset dan IP lokal tersamarkan bila tersedia.
6. **Rilis dan rollout.** Android: distribusikan AAB lewat Play internal
   testing/managed Google Play, atau APK bertanda tangan untuk MDM internal.
   iOS: gunakan TestFlight untuk pilot lalu App Store atau Apple Business
   Manager/MDM. Jalankan uji Finance pada perangkat nyata sebelum menggantikan
   aplikasi desktop.

## Catatan update mobile

Di Android, update paling stabil datang dari Play Store atau MDM; APK sideload
memerlukan tanda tangan konsisten dan konfirmasi instalasi oleh OS. Di iOS,
aplikasi tidak dapat mengganti dirinya sendiri dengan mengunduh IPA seperti
updater desktop; update dipasang melalui TestFlight, App Store, atau MDM. Karena
itu, tombol desktop **Update** tidak otomatis berlaku untuk versi mobile.

## Kriteria selesai sebelum pilot

- Total transaksi dan hasil per sheet cocok dengan output desktop untuk sampel
  Finance yang disetujui.
- File tetap lokal, ukuran workbook maksimum jelas, dan proses dapat dibatalkan
  tanpa meninggalkan file output setengah jadi.
- Accessibility: teks terbaca, kontrol bisa dijalankan dengan screen reader,
  dan UI bekerja pada layar kecil serta orientasi yang didukung.
- Penandatanganan, backup kunci release, versi, riwayat rilis, dan pemulihan
  versi terdahulu sudah didokumentasikan.
