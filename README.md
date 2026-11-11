# Sistem Manajemen Inventaris Spare Part Alat Berat (PT. STAR WAGEN INDONESIA - VIB)

Aplikasi web dashboard modern untuk manajemen stok inventaris alat berat dengan **Master Data Material**, **Daily Stock Updates (IN/OUT/UTD/SISA)**, **Bulk Data Import dari Excel (Copy-Paste)**, **Buku Besar Transaksi & Audit Log**, serta sinkronisasi multi-user real-time dua arah (*two-way continuous sync*) dengan **Google Sheets** menggunakan **Google Apps Script**.

---

## 🚀 Fitur Utama & Modul Navigasi

Sistem website kini memiliki 6 modul utama dengan navigasi tab modern yang responsif:

### 1. 📊 Dashboard KPI & Analitik Interaktif
- **KPI Real-Time**:
  - Total Master Material terdaftar (Kode Material unik).
  - Total unit fisik stok yang tersedia.
  - Peringatan stok **Understock** (di bawah batas minimum).
  - Peringatan stok **Habis (Out of Stock / OOD)** yang butuh order mendesak.
  - Peringatan stok **Overstock** (melebihi kapasitas batas maksimum).
  - **Total Nilai Inventaris (IDR)**: Valuasi fisik stok dalam Rupiah (*Price x Stok Saat Ini*).
- **Grafik Interaktif (Chart.js 4.4)** yang mengambil data aktual dari database:
  - *Distribusi Status Stok* (Doughnut Chart): Persentase stok optimal vs understock vs habis.
  - *Komposisi per Produsen / Lokasi* (Bar Chart).
  - *Tipe Servis Part* (Doughnut Chart): Periodic Service (PS) vs Non-PS.
  - *10 Part Paling Kritis* (Horizontal Bar Chart): Perbandingan stok aktual dengan safety stock.

### 2. 📦 Master Data Material
- Setiap material memiliki identitas tunggal berdasarkan **KODE MATERIAL** yang bersifat **UNIQUE**.
- Jika Anda memasukkan kode material yang sudah ada, sistem tidak membuat data baru, melainkan menghubungkannya dengan data yang sudah ada (UPSERT).
- Kolom Master Data:
  - **Nomor**
  - **Kode Material** (Primary Key / Unique Badge)
  - **Lokasi** (contoh: *PALET E1, RAK A-3*)
  - **Nama Material / Deskripsi**
  - **Price** (Harga Satuan dalam Rupiah)
  - **Stok Awal**
  - **Stok Saat Ini** (Dihitung dinamis dan konsisten dari riwayat transaksi harian)
  - **Status Kesehatan Stok**
- Fitur pencarian instan, filter lokasi, filter produsen, filter status stok, dan ekspor ke CSV & JSON.

### 3. 📅 Daily Stock Update
- Pencatatan mutasi stok harian tanpa perulangan data material.
- Kolom pencatatan harian:
  - **Tanggal**
  - **Jam**
  - **Kode Material**
  - **Nama Material**
  - **Lokasi**
  - **IN** (Barang Masuk)
  - **OUT** (Barang Keluar)
  - **UTD** (Kolom format Excel kantor)
  - **SISA** (Saldo Stok Terhitung)
- **Aturan Perhitungan Saldo (Cascade Recalculation)**:
  - Transaksi hari pertama: $\text{SISA}_1 = \text{Stok Awal} + \text{IN}_1 - \text{OUT}_1$
  - Transaksi hari berikutnya: $\text{SISA}_k = \text{SISA}_{k-1} + \text{IN}_k - \text{OUT}_k$
  - Stok akhir hari sebelumnya secara otomatis menjadi dasar saldo awal hari berikutnya.
  - Jika Anda mengedit quantity IN atau OUT pada tanggal manapun di masa lalu, sistem langsung menghitung ulang saldo hari tersebut dan seluruh hari setelahnya secara otomatis tanpa menghitung stok dua kali!
- Filter pencarian berdasarkan: Jam, Tanggal, Bulan, Tahun, Lokasi, dan Kata Kunci.

### 4. 📋 Bulk Data Import dari Excel (Fitur Paling Utama)
- Buka file Excel kantor, blok dan salin (**Ctrl+C**) baris-baris data dari tim kantor, lalu tempel (**Ctrl+V**) ke halaman **Bulk Data Import**.
- **Contoh format yang didukung:**
  ```
  KODE MATERIAL | LOKASI   | NAMA MATERIAL | PRICE  | IN | OUT | UTD
  0092S37070    | PALET E1 | BOSCH N170    | 100000 | 5  | 2   | -
  0092S47011    | PALET E1 | BOSCH N70-M   | 150000 | 3  | 1   | -
  0092T47014    | PALET E1 | BOSCH N100-M  | 200000 | 0  | 2   | -
  ```
- **Kemampuan Otomatis Sistem**:
  1. Mengenali delimiter otomatis (Tab TSV dari Excel, Pipe `|`, CSV koma/titik-koma).
  2. Mengenali kolom berdasarkan header secara cerdas.
  3. Memeriksa apakah material sudah terdaftar di Master Data.
  4. Jika sudah ada, transaksi stok hari itu dicatat/diperbarui tanpa membuat duplikat material.
  5. Jika belum ada, sistem otomatis mendaftarkan material baru tersebut ke Master Data.
  6. Menampilkan tabel **Live Preview** lengkap dengan badge status (`BARU`, `TERDAFTAR`, `DUPLIKAT`, `GAGAL`).
  7. Metrik ringkasan sebelum disimpan: Total Baris, Material Baru, Material Lama, Baris Valid, Baris Gagal.
  8. Tombol **Batal** untuk membatalkan proses sebelum data disimpan ke database.
  9. Pembuatan **Snapshot Backup Otomatis** sebelum proses import berjalan.

### 5. 🕒 Stock History & Audit Log
- **Buku Besar Transaksi (Ledger)**: Menampilkan kronologi seluruh pergerakan barang per jam, tanggal, bulan, tahun, dan lokasi.
- **Audit Log Keamanan**: Mencatat siapa yang mengubah data (Admin/User), jam berapa, jenis aksi (*CREATE_MATERIAL, UPDATE_MATERIAL, ADD_DAILY_STOCK, EDIT_DAILY_STOCK, DELETE_DAILY_STOCK, BULK_IMPORT*), serta detail perubahannya.

### 6. 📈 Reports & Valuasi Stok
- **Posisi Stok pada Tanggal Tertentu (*Stock as of Date*)**:
  - Anda dapat memilih tanggal di masa lalu, dan sistem akan menghitung posisi saldo stok aktual setiap spare part pada tanggal tersebut.
- **Valuasi Nilai Stok per Lokasi**:
  - Rekapitulasi jumlah part, total fisik unit, dan total nilai aset (Rupiah) berdasarkan masing-masing palet/rak/gudang.
- **Cetak Laporan**:
  - Ekspor ke **CSV** atau gunakan tombol **Cetak / Simpan PDF** (dengan format stylesheet khusus cetak).

---

## 🔒 Keamanan Data & Database

1. **Client Storage**: **IndexedDB (`HeavyEquipmentInventoryDB`)**
   - Mendukung penyimpanan ribuan data transaksi tanpa batasan kuota 5MB localStorage.
   - Primary key dengan **UNIQUE constraint** pada `kodeMaterial`.
   - Transaksional dan offline-ready.
2. **Cloud Database Multi-User**: **Google Sheets via `Code.gs`**
   - Terdiri dari 3 sheet yang tersinkronisasi:
     - `MASTER_MATERIAL`
     - `DAILY_STOCK`
     - `AUDIT_LOGS`
   - Tim kantor dapat berbagi data secara langsung tanpa perlu sewa server berbayar.
3. **Pusat Backup & Pemulihan (Disaster Recovery)**:
   - Unduh salinan database utuh kapan saja dalam format `.json`.
   - Fitur restore database instan dari file cadangan.

---

## 🛠️ Panduan Menghubungkan Google Sheets Cloud

1. Buka spreadsheet baru di [Google Sheets](https://sheets.new).
2. Klik menu **Ekstensi (Extensions)** &rarr; **Apps Script**.
3. Hapus kode bawaan, lalu salin seluruh isi file **`Code.gs`** yang ada di folder proyek ini ke dalamnya.
4. Klik **Deploy (Terapkan)** &rarr; **New deployment (Penerapan baru)** &rarr; Pilih jenis **Web app**.
5. Atur konfigurasi:
   - **Execute as**: *Me (Email Anda)*
   - **Who has access**: *Anyone (Siapa saja)*
6. Klik **Deploy**, beri izin akun Google Anda, lalu salin **Web app URL** yang muncul (berakhiran `/exec`).
7. Buka website inventaris, klik tombol **Google Sheets** di navbar atas, tempelkan URL tersebut, klik **Test Koneksi**, lalu klik **Kirim Semua Data ke Sheets**.

Selesai! Sekarang data di website dan Google Sheets akan tersinkronisasi secara real-time.
