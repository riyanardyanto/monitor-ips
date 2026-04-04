# Monitor IPS

Desktop GUI untuk memeriksa kelengkapan form Excel IPS berdasarkan cell wajib di `field_cell.txt`.

Project ini juga menyediakan IPS Generator berbasis HTML untuk membuat dokumen IPS baru melalui tombol `Buat IPS` di aplikasi utama atau langsung dari `assets/ips.html`.

## Menjalankan Aplikasi

```powershell
.venv\Scripts\python.exe main.py
```

## Build EXE

Gunakan file spec yang sudah disiapkan untuk membuat executable final:

```powershell
.venv\Scripts\python.exe -m PyInstaller --clean "IPS Checker.spec"
```

Hasil build akan tersedia di folder:

```powershell
dist\IPS Checker.exe
```

Saat EXE dijalankan, aplikasi akan otomatis membuat struktur runtime berikut di samping file executable:

```text
data/
	config/
		ips-checker-config.json
	database/
		ips-follow-up-database.csv
	docs/
		ips-checker-user-guide.pdf
		ips-generator-user-guide.pdf
```

Catatan:

- Gunakan file `IPS Checker.spec` agar resource Flet, `field_cell.txt`, `assets/ips.html`, dan PDF user guide ikut terbundle dengan benar.
- Jangan menjalankan command PyInstaller minimal tanpa file spec, karena aplikasi membutuhkan data tambahan saat dijalankan sebagai `.exe`.
- File config dan database lama di root project akan dimigrasikan otomatis ke folder `data` jika file tujuan belum ada.

## Fitur

- Pilih folder atau satu file Excel
- Pilih file mapping field
- Tombol `Buat IPS` untuk membuka generator form IPS baru
- Menyimpan config lokal untuk path terakhir, lokasi CSV terakhir, mode sheet terakhir, dan ukuran window
- Menampilkan hasil per file beserta jumlah field kosong per section
- Mendukung export sesuai view aktif di panel kanan: hasil pemeriksaan atau follow up countermeasure
- Menyediakan user guide checker dan generator secara otomatis di folder runtime `data/docs`

## Fitur GUI

- `result panel` memiliki 2 tampilan yang dapat diganti dari menu bar: `Tabel Hasil Pemeriksaan` dan `Tabel Follow Up Countermeasure`.
- Tombol `Buat IPS` di panel kiri membuka generator IPS dari `assets/ips.html`.
- Kolom `Action` pada tabel hasil menyediakan tombol untuk membuka dialog `Detail`, `Follow Up`, dan `Open file` per file.
- Dialog `Detail` menampilkan informasi file, participant, status, kelengkapan total, dan tabel persentase kelengkapan per section.
- Dialog `Follow Up` menampilkan data section `1.5` yang terisi, yaitu `countermeasure`, `responsible`, dan `due date`.
- Tombol `Open file` membuka file IPS Excel asli dari baris hasil pemeriksaan yang dipilih.
- Tampilan `Tabel Follow Up Countermeasure` memakai database lokal otomatis di file `data\database\ips-follow-up-database.csv`.
- Database follow up otomatis dibuat jika belum ada, lalu diisi dari data section `1.5` setiap selesai pemeriksaan.
- Kolom database follow up otomatis: `nama file`, `countermeasure`, `responsible`, `due date`, dan `status`.
- Setiap baris pada `Tabel Follow Up Countermeasure` memiliki tombol edit untuk mengubah `status` langsung dari GUI dan menyimpannya ke database lokal otomatis.
- Tombol `Share` pada dialog `Detail` menampilkan QR Code berisi ringkasan kelengkapan file yang sedang dipilih.
- Tombol `Share` pada dialog `Follow Up` menampilkan QR Code berisi summary data follow up section `1.5`.
- Tombol `Export` di AppBar menampilkan pilihan export `PDF`, `JPG`, `QR Code`, atau `Excel` sesuai view aktif.
- Dialog `Share` juga menyediakan tombol `Copy Summary` untuk menyalin isi summary ke clipboard.
- Aplikasi menyimpan file config lokal `data\config\ips-checker-config.json` untuk mengingat folder/file terakhir, lokasi export terakhir, mode sheet, nama sheet custom, dan ukuran window terakhir.

## Dokumentasi

- Panduan aplikasi utama tersedia di `docs/ips-checker-user-guide.pdf` pada source project.
- Panduan IPS Generator tersedia di `docs/ips-generator-user-guide.pdf` pada source project.
- Saat aplikasi berjalan, kedua PDF tersebut akan disalin ke folder runtime `data/docs`.

## Kriteria Complete

Aturan umum:

- Field dianggap complete jika semua cell yang didefinisikan pada field tersebut terisi.
- Nilai kelengkapan per section dihitung dari jumlah field atau rule section yang valid dibanding total rule pada section tersebut.

Aturan khusus per section:

1. Section 1.1

- `trigger (C6, C8, C10, C12)` dianggap complete jika minimal ada 1 cell yang terisi.
- Field lain di section 1.1 tetap harus terisi pada semua cell yang didefinisikan.

2. Section 1.2

- Semua field dihitung per field tunggal.
- Field dianggap complete jika cell-nya terisi.

3. Section 1.3

- Untuk setiap item seperti `cil (C66, E66, G66)`, tepat 1 cell harus berisi `OK`, `NOK`, atau `NA`.
- Cell lainnya harus kosong.
- Jika tidak memenuhi pola tersebut, item dianggap tidak complete.

4. Section 1.4

- Section 1.4 dianggap complete jika minimal ada 1 row yang terisi penuh pada kolom berikut:
	`standard (M)`, `Tindakan (T)`, `responsible (Y)`, `due date (AA)`.
- Row yang dicek hanya: `66`, `68`, `70`, `72`, `74`, `76`, `78`.
- Contoh row valid: `M66`, `T66`, `Y66`, `AA66` semuanya terisi.
- Jika ada row lain yang terisi sebagian, section 1.4 dianggap error.
- Row lain hanya boleh full semua atau kosong semua.

5. Section 1.5

- Section 1.5 dianggap complete jika minimal ada 1 row yang terisi penuh pada kolom berikut:
	`countermeasure (X)`, `responsible (Y)`, `due date (Z)`.
- Row yang dicek hanya: `88`, `97`, `106`, `115`, `124`, `133`.
- Contoh row valid: `X88`, `Y88`, `Z88` semuanya terisi.
- Jika ada row lain yang terisi sebagian, section 1.5 dianggap error.
- Row lain hanya boleh full semua atau kosong semua.

## Rumus Completeness

- Kelengkapan per section:

	`((total rule section - rule invalid section) / total rule section) * 100`

- Kelengkapan total file:

	rata-rata dari completeness semua section.
