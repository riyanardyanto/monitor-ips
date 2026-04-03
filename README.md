# Monitor IPS

Desktop GUI untuk memeriksa completeness form Excel IPS berdasarkan cell wajib di field_cell.txt.

## Jalankan GUI

```powershell
.venv\Scripts\python.exe main.py
```

## Build EXE

Build final executable menggunakan file spec yang sudah disiapkan:

```powershell
.venv\Scripts\python.exe -m PyInstaller --clean "IPS Checker.spec"
```

Hasil build akan tersedia di folder:

```powershell
dist\IPS Checker.exe
```

Catatan:

- Gunakan file `IPS Checker.spec` agar resource Flet dan `field_cell.txt` ikut terbundle dengan benar.
- Jangan build dengan command PyInstaller minimal tanpa file spec, karena aplikasi membutuhkan data tambahan saat dijalankan sebagai `.exe`.

## Fitur

- Pilih folder atau satu file Excel
- Pilih file mapping field
- Simpan hasil ringkas ke CSV
- Simpan config lokal untuk path terakhir, lokasi CSV terakhir, mode sheet terakhir, dan ukuran window
- Tampilan hasil per file dengan jumlah field kosong per section

## Fitur GUI

- Kolom `Action` pada tabel hasil menyediakan tombol untuk membuka dialog `Detail` dan `Follow Up`.
- Dialog `Detail` menampilkan informasi file, participant, status, completeness total, dan tabel persentase completeness per section.
- Dialog `Follow Up` menampilkan data section `1.5` yang terisi, yaitu `countermeasure`, `responsible`, dan `due date`.
- Tombol `Share` pada dialog `Detail` menampilkan QR Code berisi summary completeness file yang sedang dipilih.
- Tombol `Share` pada dialog `Follow Up` menampilkan QR Code berisi summary data follow up section `1.5`.
- Tombol `Export` di side panel menampilkan pilihan export `PDF`, `JPG`, atau `JSON` untuk summary laporan data completeness form IPS yang sedang ada di tabel.
- Dialog `Share` juga menyediakan tombol `Copy Summary` untuk menyalin isi summary ke clipboard.
- Aplikasi menyimpan file config lokal `ips-checker-config.json` untuk mengingat folder/file terakhir, lokasi CSV terakhir, mode sheet, nama sheet custom, dan ukuran window terakhir.

## Kriteria Complete

Aturan umum:

- Field dianggap complete jika semua cell yang didefinisikan di field tersebut terisi.
- Completeness per section dihitung dari jumlah field atau rule section yang valid dibanding total rule pada section tersebut.

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

- Completeness section:

	`((total rule section - rule invalid section) / total rule section) * 100`

- Completeness total file:

	rata-rata dari completeness semua section.
