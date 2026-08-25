# Excel Smart Assistant Add-In 📊✨

Asisten otomatisasi dan produktivitas cerdas untuk Microsoft Excel berbasis **Office.js** modern (Web Add-in, bukan VBA/Macro) dengan antarmuka Fluent UI dan integrasi Google Gemini AI.

---

## 🌟 Fitur-Fitur Unggulan

### 1. ✨ AI Formula Assistant & Explainer
* **Generate Formula:** Menerjemahkan bahasa alami Indonesia menjadi formula Excel siap pakai ke sel aktif.
* **💡 Jelaskan Rumus (Formula Explainer):** Membaca formula pada sel aktif dan menjelaskannya langkah demi langkah dengan bantuan AI Gemini / parser lokal.
* **🛠️ Auto-Fix Error:** Otomatis membungkus formula yang rentan error dengan proteksi `IFERROR(..., 0)` atau `IFERROR(..., "Tidak Ditemukan")`.

### 2. ⚡ Quick Statistics & Auto-Total
* **Live Selection Stats:** Menghitung total (*Sum*), rata-rata (*Average*), nilai min/max, dan jumlah sel unik secara instan dari rentang sel yang diblok.
* **➕ Baris Total Akuntansi:** Otomatis menyisipkan baris total di bawah tabel dengan rumus `=SUM(...)`, teks tebal, dan standar border ganda akuntansi.

### 3. 🧹 1-Click Data Cleaner
* **✂️ Trim Spasi:** Menghapus spasi liar dan spasi ganda.
* **🔤 Proper Case:** Mengubah huruf awal setiap kata menjadi kapital.
* **🔠 UPPERCASE / 🔡 lowercase:** Transformasi huruf instan.
* **💰 Format Rp & 📅 Format Tanggal:** Format angka keuangan standar Rupiah (`Rp #,##0`) dan tanggal Indonesia (`DD/MM/YYYY`).

### 4. 🔍 Duplikat & Text Tools
* **🎨 Sorot Duplikat:** Menandai sel bernilai ganda dengan highlight kuning/amber lembut.
* **🗑️ Hapus Duplikat:** Membersihkan baris data duplikat secara instan.
* **📋 Ekstrak Unik:** Menyalin daftar unik ke kolom baru di samping tabel.
* **🔀 Pecah Kolom (Split Text):** Memisahkan teks berdasarkan spasi, koma, strip, atau tanda lainnya menjadi banyak kolom.
* **🔗 Gabung Kolom (Merge Text):** Menggabungkan beberapa kolom yang diblok menjadi satu kolom.

### 5. 📈 1-Click Auto Chart Generator
* **📊 Diagram Batang / Kolom:** Perbandingan nilai antar kategori.
* **📈 Grafik Garis:** Menampilkan tren data dari waktu ke waktu.
* **🥧 Diagram Lingkaran:** Visualisasi proporsi persentase.

### 6. 🛡️ Data Quality Auditor
* **🔍 Sorot Sel Kosong:** Mendeteksi sel bolong di tengah tabel data.
* **🔢 Teks ➔ Angka:** Mengonversi angka yang tersimpan sebagai string kembali menjadi numerik murni agar rumus kalkulasi tidak menghasilkan 0 / error.
* **🧹 Bersihkan Format:** Mereset kembali warna latar highlight pemeriksaan data.

### 7. 📊 Template Generator Otomatis
* Evaluasi Kinerja Karyawan, Laporan Arus Kas, Jadwal Proyek / Project Tracker, Daftar Tagihan (Invoice), Media Partner Tracker, dan Tinjauan Pustaka dengan styling hijau Fluent `#107c41`, autofit, data validation dropdown, dan freeze panes.

### 8. 📑 Mass Action: Split Sheets
* Memisahkan data tabel besar ke banyak worksheet baru berdasarkan nilai kategori unik secara otomatis.

---

## 🛠️ Tech Stack & Arsitektur

* **UI & Styling:** HTML5, Vanilla CSS3 (Fluent UI Design System)
* **Logika & API:** Vanilla JavaScript (`taskpane.js`) + `Office.js` API (`Excel.run`)
* **AI Engine:** Google Gemini 1.5 Flash API + Local Smart Parser Fallback
* **Konfigurasi:** XML Manifest (`manifest.xml`)
* **Live Deployment:** [GitHub Pages](https://chigoov.github.io/ADD-IN-EXCEL-/index.html)

---

## 🚀 Cara Menjalankan

### Cara 1: Excel for Web (Browser)
1. Buka [https://excel.new](https://excel.new).
2. Masuk ke tab **Sisipkan / Insert** ➔ **Add-ins** ➔ **Add-in Saya**.
3. Klik **Unggah Add-in (Upload My Add-in)** ➔ Pilih `manifest.xml`.

### Cara 2: Excel Desktop (Aplikasi Laptop)
1. Buka aplikasi **Excel Desktop**.
2. Pada tab **Home / Beranda** di pojok kanan, klik tombol **"Buka Smart Assistant"**.
