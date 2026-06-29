# 🧹 Cleaner & Calculate AVG Piutang

> **Pembersih batch otomatis laporan AR dari Accurate — sekaligus menghitung rata-rata umur piutang dan nilai faktur per pelanggan**

Skrip Python satu file yang memproses **semua file `.xls`** di folder yang sama secara sekaligus (batch). Membaca ekspor laporan piutang dari Accurate, membersihkan format angka Indonesia (titik sebagai pemisah ribuan), memilih 9 kolom relevan, lalu menghasilkan file `_hasil.xlsx` per sumber — berisi dua sheet: **Data Bersih** (baris per faktur) dan **Ringkasan** (rata-rata per pelanggan).

---

## 📋 Daftar Isi

- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Detail](#-alur-kerja-detail)
- [Logika Pembersihan Data](#-logika-pembersihan-data)
- [Output & Struktur File Hasil](#-output--struktur-file-hasil)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## ✨ Fitur Utama

- **Batch processing** — Memproses semua file `.xls` di folder sekaligus dalam satu eksekusi, cocok untuk rekap bulanan yang terdiri dari banyak file.
- **Auto-detect posisi header** — Mencari baris yang mengandung `"No. Faktur"` secara dinamis, sehingga tidak terganggu meskipun ekspor Accurate menambahkan baris judul di atasnya.
- **Pembersihan angka format Indonesia** — Menangani titik sebagai pemisah ribuan (format lokal Accurate: `1.234.567`) dan koma sebagai karakter tak terduga hasil export, mengonversinya ke integer murni.
- **Strip header berulang** — Menghapus baris yang merupakan pengulangan header (baris `"No. Faktur"` yang muncul di tengah data akibat page break ekspor).
- **Ringkasan rata-rata per pelanggan** — Menghitung `AVG UMUR PIUTANG`, `AVG NILAI FAKTUR`, dan `JUMLAH INVOICE` per kombinasi kode & nama pelanggan, diurutkan berdasarkan nomor pelanggan.
- **Formatting & auto-fit otomatis** — Kolom angka diformat sebagai `#,##0`, lebar semua kolom disesuaikan otomatis dengan konten terpanjang di setiap kolom.
- **Toleransi error per file** — Jika satu file gagal diproses (korup, format berbeda), skrip melanjutkan ke file berikutnya dan melaporkan error tanpa menghentikan seluruh batch.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlrd
```

| Library | Kegunaan |
|---|---|
| `pandas` | Baca `.xls`, bersihkan, groupby, simpan ke `.xlsx` |
| `openpyxl` | Engine penulisan `.xlsx` dengan formatting sel |
| `xlrd` | Engine pembacaan file legacy `.xls` dari Accurate |
| `os`, `glob` | Deteksi file dan manipulasi path (standard library) |

> **Catatan:** `xlrd` versi 2.x ke atas **hanya** mendukung format `.xls` lama (BIFF8). Jika muncul error saat membaca, install dengan:
> ```bash
> pip install "xlrd>=1.0.0,<2.0.0"
> ```

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

Letakkan satu atau lebih file ekspor AR dari Accurate (format `.xls`) ke dalam **folder yang sama** dengan `ClenanerAndCalculateAVG.py`.

```
📁 folder-kerja/
├── ClenanerAndCalculateAVG.py
├── AR_Jan_2025.xls        ← file input 1
├── AR_Feb_2025.xls        ← file input 2
└── AR_Mar_2025.xls        ← file input 3
```

File `.xls` harus mengandung kolom-kolom berikut (nama persis, posisi bebas):
`No. Faktur`, `Tgl Faktur`, `Umur`, `Kode`, `Nama Pelanggan`, `Nilai Faktur`, `Terutang`, `Nama Gudang`, `Nama Penjual`

### Langkah 2 — Jalankan

```bash
python ClenanerAndCalculateAVG.py
```

### Langkah 3 — Ambil hasil

Setiap file `.xls` yang berhasil diproses akan menghasilkan satu file `.xlsx` baru di folder yang sama, dengan nama `[nama_asli]_hasil.xlsx`:

```
📁 folder-kerja/
├── ClenanerAndCalculateAVG.py
├── AR_Jan_2025.xls
├── AR_Jan_2025_hasil.xlsx   ← hasil untuk file Jan
├── AR_Feb_2025.xls
├── AR_Feb_2025_hasil.xlsx   ← hasil untuk file Feb
└── ...
```

Output di terminal:
```
--> Berhasil memproses: AR_Jan_2025.xls menjadi AR_Jan_2025_hasil.xlsx
--> Berhasil memproses: AR_Feb_2025.xls menjadi AR_Feb_2025_hasil.xlsx
--> Gagal memproses AR_Mar_2025.xls: [pesan error jika ada masalah]
```

---

## 🔄 Alur Kerja Detail

```
[Mulai]
   │
   ├─── Scan folder saat ini
   │       Cari semua file *.xls dengan glob
   │       Jika tidak ada → keluar (loop kosong)
   │
   └─── Untuk setiap file .xls yang ditemukan:
           │
           ├─── [1] Baca file tanpa header (header=None)
           │
           ├─── [2] Deteksi baris header
           │         Scan setiap baris, cari yang mengandung teks "No. Faktur"
           │         Jika tidak ditemukan → lewati file ini, lanjut ke berikutnya
           │
           ├─── [3] Tetapkan nama kolom
           │         Ambil baris header sebagai nama kolom
           │         Strip \xa0 (non-breaking space) dan whitespace dari nama kolom
           │
           ├─── [4] Potong data mulai baris setelah header
           │
           ├─── [5] Filter & bersihkan kolom
           │         Hapus baris yang kolom-kolom wajib (9 kolom) ada NaN-nya
           │         Hapus baris di mana No. Faktur berisi teks "No. Faktur"
           │         (menghilangkan header berulang akibat page break ekspor)
           │
           ├─── [6] Normalisasi angka (format Indonesia → integer)
           │         No. Faktur : split koma → hapus titik → to_numeric → int
           │         Umur       : split koma → hapus titik → to_numeric → int
           │         Nilai Faktur: split koma → hapus titik → to_numeric → int
           │         Terutang   : split koma → hapus titik → to_numeric → int
           │
           ├─── [7] Pilih 9 kolom relevan → simpan sebagai df_clean
           │
           ├─── [8] Hitung ringkasan per pelanggan
           │         GROUP BY: Kode + Nama Pelanggan
           │         AGG:
           │           Umur         → mean → round → int  (AVG UMUR PIUTANG)
           │           Nilai Faktur → mean → round → int  (AVG NILAI FAKTUR)
           │           No. Faktur   → count              (JUMLAH INVOICE)
           │         SORT: by No. Pelanggan (Kode)
           │
           └─── [9] Simpan sebagai [nama_asli]_hasil.xlsx
                     Sheet "Data Bersih"  → df_clean (9 kolom, format #,##0 pada angka)
                     Sheet "Ringkasan"    → grouped (5 kolom, format #,##0 pada AVG NILAI)
                     Auto-fit lebar kolom → kedua sheet
```

---

## 🔍 Logika Pembersihan Data

### Pembersihan format angka Indonesia

Accurate mengekspor angka dengan format lokal Indonesia yang menggunakan **titik** sebagai pemisah ribuan:

| Nilai asli di `.xls` | Setelah pembersihan |
|---|---|
| `1.234.567` | `1234567` |
| `75` | `75` |
| `1.234.567,89` | `1234567` *(bagian desimal dibuang)* |
| `0` | `0` |

Langkah pembersihan per kolom angka:
```
nilai asli
  └─ .split(",")[0]    → ambil bagian sebelum koma pertama (buang desimal)
  └─ .replace(".", "") → hapus semua titik (pemisah ribuan)
  └─ .strip()          → hapus spasi di awal/akhir
  └─ pd.to_numeric()   → konversi ke angka (NaN jika gagal)
  └─ .fillna(0)        → ganti NaN dengan 0
  └─ .astype(int)      → bulatkan ke integer
```

### Penanganan header berulang

Ekspor Accurate sering menyertakan ulang baris header di setiap halaman baru (page break). Baris semacam ini tampak seperti baris data namun `No. Faktur`-nya berisi teks literal `"No. Faktur"`. Skrip menghapus baris seperti ini dengan:

```python
df = df[df["No. Faktur"].astype(str).str.strip() != "No. Faktur"]
```

### Penanganan `\xa0` (non-breaking space)

Ekspor Excel terkadang menyisipkan karakter `\xa0` (non-breaking space, ASCII 160) di nama kolom. Skrip membersihkannya saat membangun nama kolom:

```python
columns = [str(c).replace('\xa0', ' ').strip() for c in df_raw.iloc[header_idx]]
```

### Penentuan `JUMLAH INVOICE`

Kolom ini menggunakan **count** (jumlah baris) dari `No. Faktur` dalam group, bukan sum. Ini memberikan jumlah faktur yang tercatat atas nama pelanggan tersebut dalam file.

---

## 📤 Output & Struktur File Hasil

### Penamaan file

```
[nama_file_asli]_hasil.xlsx
```

Contoh: `AR_Jan_2025.xls` → `AR_Jan_2025_hasil.xlsx`

### Sheet 1 — `Data Bersih`

Data per faktur, sudah dibersihkan dan difilter. Berisi 9 kolom:

| Kolom | Tipe | Keterangan |
|---|---|---|
| `No. Faktur` | Integer | Nomor faktur, titik dihapus |
| `Tgl Faktur` | Teks/Tanggal | Tanggal faktur (format asli dari Accurate) |
| `Umur` | Integer | Umur piutang dalam hari |
| `Kode` | Teks | Kode pelanggan |
| `Nama Pelanggan` | Teks | Nama pelanggan |
| `Nilai Faktur` | Integer | Nilai faktur asli, format `#,##0` |
| `Terutang` | Integer | Sisa piutang yang belum dibayar, format `#,##0` |
| `Nama Gudang` | Teks | Gudang asal transaksi |
| `Nama Penjual` | Teks | Nama sales/penjual |

### Sheet 2 — `Ringkasan`

Satu baris per pelanggan (distinct `Kode` + `Nama Pelanggan`), diurutkan berdasarkan `No. Pelanggan`. Berisi 5 kolom:

| Kolom | Tipe | Keterangan |
|---|---|---|
| `No. Pelanggan` | Teks | Kode pelanggan (dari kolom `Kode`) |
| `Nama Pelanggan` | Teks | Nama pelanggan |
| `AVG UMUR PIUTANG` | Integer | Rata-rata umur piutang semua faktur pelanggan ini (hari, dibulatkan) |
| `AVG NILAI FAKTUR` | Integer | Rata-rata nilai faktur semua faktur pelanggan ini (Rp, dibulatkan), format `#,##0` |
| `JUMLAH INVOICE` | Integer | Jumlah baris faktur tercatat atas pelanggan ini di file |

**Contoh ilustrasi:**

Jika pelanggan `MGL-1001 / Toko Makmur` memiliki 3 faktur:

| No. Faktur | Umur | Nilai Faktur |
|---|---|---|
| 10001 | 30 | 5.000.000 |
| 10002 | 45 | 7.500.000 |
| 10003 | 60 | 2.500.000 |

Maka baris ringkasannya:

| No. Pelanggan | Nama Pelanggan | AVG UMUR PIUTANG | AVG NILAI FAKTUR | JUMLAH INVOICE |
|---|---|---|---|---|
| MGL-1001 | Toko Makmur | 45 | 5.000.000 | 3 |

---

## 🛠️ Troubleshooting

### ❌ Tidak ada file yang diproses (tidak ada output sama sekali)
Pastikan file `.xls` berada di **folder yang sama** dengan `ClenanerAndCalculateAVG.py`, bukan di subfolder. Skrip menggunakan `glob("*.xls")` yang hanya mencari di direktori kerja saat ini. Coba jalankan dari terminal di dalam folder yang benar:
```bash
cd path/ke/folder/kerja
python ClenanerAndCalculateAVG.py
```

### ❌ `Header 'No. Faktur' tidak ditemukan di [file]`
Skrip tidak menemukan baris yang mengandung teks `"No. Faktur"` di seluruh isi file. Kemungkinan penyebab: (1) nama kolom di ekspor Accurate berbeda (misalnya `"No Faktur"` tanpa titik), (2) file yang diproses bukan laporan AR (mungkin laporan lain yang ikut terscan), atau (3) file `.xls` menggunakan encoding tidak standar.

### ❌ Error `xlrd.biff_xls: Excel xlsx file; not supported`
`xlrd` versi baru tidak membaca `.xlsx`. Pastikan file input benar-benar berformat `.xls` (bukan `.xlsx` yang diubah ekstensinya). Jika file aslinya `.xlsx`, ubah terlebih dahulu ke `.xls` via Accurate atau Save As di Excel.

### ❌ Kolom `Terutang` atau `Nilai Faktur` hasilnya `0` semua
Kemungkinan format angka di file sumber menggunakan pemisah desimal yang tidak terduga. Cek isi mentah sel tersebut — jika ada karakter selain digit, titik, dan koma, fungsi `parse` akan mengembalikan `0` via `.fillna(0)`. Buka file `.xls` di Excel dan periksa format sel kolom angkanya.

### ❌ Ringkasan berisi nama pelanggan yang sama dua kali
Terjadi karena kombinasi `Kode + Nama Pelanggan` tidak unik — misalnya ada satu pelanggan dengan nama yang berbeda sedikit (spasi ekstra, huruf kapital) di dua baris. Cek dan seragamkan nama di data sumber Accurate.

### ❌ `Gagal memproses [file]: [pesan error]`
Skrip menangkap semua exception per file dan melanjutkan ke file berikutnya. Salin pesan error yang muncul dan cek kemungkinan penyebab: file terkunci oleh Excel yang masih terbuka, file korup, atau dependensi library tidak terinstall.

---

## 📌 Catatan Penting

- **Hanya memproses `.xls`** — File `.xlsx` tidak akan terbaca karena `glob("*.xls")` bersifat eksak. Jika file ekspor Accurate sudah dalam format `.xlsx`, ubah ekstensinya terlebih dahulu atau ekspor ulang ke `.xls`.
- **Semua `.xls` di folder ikut diproses** — Pastikan tidak ada file `.xls` lain (template, catatan, dll.) di folder yang sama, karena semuanya akan ikut diproses.
- **File hasil tidak menimpa file sumber** — Output selalu berekstensi `.xlsx` dengan suffix `_hasil`, sehingga file `.xls` asli aman.
- **Urutan baris di Ringkasan** — Pengurutan dilakukan berdasarkan kolom `No. Pelanggan` (kode pelanggan), bukan nama. Jika kode mengandung prefix alfanumerik (mis. `MGL-`, `YY-`), urutan akan bersifat leksikografis, bukan numerik murni.
- **AVG adalah rata-rata sederhana** — `AVG UMUR PIUTANG` dan `AVG NILAI FAKTUR` dihitung sebagai rata-rata aritmatika biasa dari seluruh faktur pelanggan tersebut di file, bukan rata-rata tertimbang.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
