# Sistem Manajemen Inventaris Spare Part Alat Berat (Heavy Equipment Inventory)

Aplikasi web dashboard modern untuk manajemen stok, monitoring kesehatan inventaris, dan sinkronisasi real-time dua arah (*two-way continuous editing*) dengan **Google Sheets** menggunakan **Google Apps Script**.

---

## 🚀 Fitur Utama

1. **Dashboard KPI & Indikator Kesehatan Stok**:
   - Total item terdaftar (338 spare part dari data Excel).
   - Total fisik unit spare part yang tersedia.
   - Peringatan stok **Understock** (jumlah stok di bawah batas minimum).
   - Peringatan stok **Habis (Out of Stock / OOD)** yang butuh order mendesak.
   - Peringatan stok **Overstock** (melebihi kapasitas batas maksimum).
   - Ringkasan produsen alat berat (*Mercedes-Benz / MBZ, Atlas Copco / ACI, Sakai, SIP*).

2. **Analitik Visual Interaktif (Chart.js)**:
   - **Distribusi Status Stok** (Doughnut Chart): Memantau persentase stok aman vs kritis.
   - **Komposisi Part per Produsen** (Bar Chart): Jumlah suku cadang berdasarkan brand alat berat.
   - **Tipe Servis Part** (Doughnut Chart): Perbandingan suku cadang *Periodic Service (PS)* vs *Non-PS*.
   - **10 Part Paling Kritis** (Horizontal Bar Chart): Membandingkan stok fisik saat ini dengan batas minimum (*safety stock*).

3. **Pencarian & Filter Canggih**:
   - Pencarian cerdas instan berdasarkan *Part Number (Final PN)*, *Alternate Part Number*, *Deskripsi*, dan *Model Unit*.
   - Filter cepat (Tab Pills): *Semua, Kritis / Butuh Order, Habis, Understock, Optimal, Overstock, Periodic Service*.
   - Filter dropdown: Produsen (*MBZ, ACI, SAKAI, SIP*), Model Unit (*Axor 2528 E4, Arocs 4042K, Arocs 4845 E5, QAS170, TL B7+, SV526, dll.*), Status Stok, dan Tipe Servis.
   - Pengurutan dinamis (*Column Sorting*) pada tabel.
   - Pengaturan jumlah baris per halaman (10, 25, 50, 100, atau Semua).

4. **Operasi Data & CRUD Lengkap**:
   - **Quick Adjust**: Tambah (+1) atau kurangi (-1) stok langsung dari baris tabel dengan satu klik.
   - **Tambah Part Baru**: Form modal lengkap dengan validasi.
   - **Edit Detail Part**: Ubah nomor part, nomor alternatif, batas min/max, DOI (*Days of Inventory*), dan catatan (*remarks*).
   - **Hapus Part**: Disertai konfirmasi keamanan.

5. **Continuous Editing dengan Google Sheets**:
   - **Edit di Spreadsheet**: Anda atau tim dapat mengedit Google Sheet kapan saja (misal menambah baris baru, mengubah stok di kolom Qty), lalu klik **Sync Now** di web (atau biarkan fitur **Auto-Sync** berjalan di latar belakang setiap 1-5 menit).
   - **Edit di Web**: Setiap perubahan stok atau penambahan part di web langsung tersimpan ke Google Sheet secara otomatis!
   - **Bulk Push 1-Klik**: Kirim seluruh 338 data awal ke Google Sheet baru Anda hanya dengan satu kali klik.
   - **Live Embed Preview**: Lihat langsung spreadsheet di dalam modal dashboard.

6. **Tampilan & Ekspor Data**:
   - Desain responsif (*mobile & desktop*) dengan gaya visual modern (*glassmorphism*, dark & light theme toggle).
   - Ekspor data hasil filter ke format **CSV** (kompatibel Excel) dan **JSON**.
   - Fitur cetak laporan (*Print / Save to PDF*).

---

## 🛠️ Teknologi yang Digunakan

- **HTML5**: Struktur semantik modern.
- **CSS3**: Desain kustom responsif, variabel tema, *glassmorphism*, dark/light mode.
- **JavaScript (ES6+)**: Logika inventaris, manajemen *state*, integrasi REST API Google Apps Script, penyimpanan lokal (*LocalStorage*).
- **Chart.js 4.4**: Visualisasi data grafik interaktif dan responsif.
- **Google Sheets & Google Apps Script (`Code.gs`)**: Database *cloud-based* yang mendukung pengeditan tanpa henti (*continuous editing*).
- **Font Awesome 6**: Ikon antarmuka profesional.

---

## 📂 Struktur File

```
d:\abdul muiz\
├── index.html       # Halaman utama dashboard inventaris
├── styles.css       # File stylesheet styling dashboard (Dark & Light theme)
├── app.js           # Logika aplikasi, Chart.js, filter, pagination & Google Sheets sync
├── data.js          # Dataset 338 spare part dari file Excel
├── Code.gs          # Skrip Google Apps Script untuk backend Google Sheets
└── README.md        # Dokumentasi dan panduan penggunaan
```

---

## 📖 Panduan Menjalankan Website

1. Anda cukup membuka file **`index.html`** langsung dengan mengklik dua kali di File Explorer, atau klik kanan lalu pilih **Open with Google Chrome / Edge**.
2. Website langsung siap digunakan seketika dengan seluruh **338 data spare part** yang sudah terisi (*offline-ready*).
3. Anda dapat langsung mencoba pencarian, filter, grafik interaktif, dan tombol ubah stok (+ / -).

---

## 🔗 Panduan Menghubungkan Google Sheets (Hanya 3 Menit)

Agar spreadsheet Google Sheets Anda dapat terus-menerus diedit bersama tim dan tersinkronisasi ke web, ikuti langkah berikut:

### Langkah 1: Buat Spreadsheet Baru
1. Buka browser dan kunjungi: [https://sheets.new](https://sheets.new)
2. Beri nama spreadsheet Anda, misalnya: `Inventaris Spare Part Alat Berat`.

### Langkah 2: Pasang Skrip Google Apps Script
1. Di Google Sheets, klik menu **Ekstensi (Extensions)** &rarr; **Apps Script**.
2. Hapus seluruh isi kode bawaan di editor `Code.gs`.
3. Buka file **`Code.gs`** yang ada di folder proyek ini (atau salin dari tombol *Salin Kode* di menu Google Sheets pada website), lalu paste ke editor Apps Script.
4. Klik ikon **Simpan (Save)** (ikon disket).

### Langkah 3: Deploy sebagai Web App
1. Di pojok kanan atas Apps Script, klik tombol biru **Deploy (Terapkan)** &rarr; pilih **New deployment (Penerapan baru)**.
2. Klik ikon gerigi di sebelah *Select type*, pilih **Web app (Aplikasi Web)**.
3. Konfigurasikan:
   - **Description**: `Inventaris API`
   - **Execute as**: `Me (Email Anda)`
   - **Who has access**: `Anyone (Siapa saja)` *(Sangat penting agar website dapat membaca dan menyimpan data)*
4. Klik **Deploy**, lalu klik **Authorize access** (Izinkan akses akun Google Anda).
5. Salin teks **Web app URL** yang muncul (berakhiran `/exec`).

### Langkah 4: Masukkan URL ke Website
1. Kembali ke website dashboard, klik tombol hijau **Google Sheets** di navbar atas.
2. Tempelkan (*paste*) URL Web App tadi ke kolom **Google Apps Script Web App URL**.
3. Klik tombol **Test Koneksi**. Jika muncul notifikasi hijau *Koneksi Berhasil*, klik tombol **Kirim Semua Data ke Sheets** untuk mengisi 338 part ke spreadsheet Anda!
4. Centang pilihan **Aktifkan Auto-Sync Otomatis**, lalu klik **Simpan Pengaturan**.

Selesai! Sekarang Anda dapat mengedit stok langsung dari Google Sheets maupun dari website, dan keduanya akan selalu terhubung secara real-time!
