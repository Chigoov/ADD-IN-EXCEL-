# Excel Smart Assistant Add-In 📊✨

Asisten otomatisasi dan produktivitas cerdas untuk Microsoft Excel berbasis **Office.js** modern (Web Add-in, bukan VBA/Macro).

---

## 🚀 Fitur Utama (MVP)

1. **✨ AI Formula Assistant & Explainer**
   - Mengubah prompt bahasa alami menjadi formula Excel langsung di sel aktif (misal: *“jumlahkan A1 sampai A10”* -> `=SUM(A1:A10)`).
2. **🧹 1-Click Smart Data Cleaner**
   - **Trim Spasi**: Membersihkan spasi di awal, akhir, dan spasi ganda di tengah teks.
   - **Proper Case**: Mengubah huruf awal setiap kata menjadi kapital secara instan pada rentang sel yang dipilih.
3. **📊 Template Generator (Matriks Otomatis)**
   - Otomatisasi pembuatan tabel matriks terstruktur (Evaluasi Kinerja Karyawan, Media Partner Tracker, Tinjauan Pustaka) lengkap dengan header hijau Fluent `#107c41`, autofit, data validation dropdown, dan freeze panes.
4. **📑 Mass Action: Split Sheets**
   - Memisahkan data tabel ke banyak worksheet baru berdasarkan nilai kategori unik secara otomatis.

---

## 🛠️ Tech Stack & Arsitektur

- **Frontend:** HTML5, CSS3 (Desain Fluent UI)
- **Logika:** Vanilla JavaScript (`taskpane.js`)
- **API Excel:** Office.js (`Excel.run`)
- **Manifest:** XML (`manifest.xml`)
- **Runtime Server:** Node.js HTTPS

---

## 💻 Cara Menjalankan Lokal

1. **Jalankan Server Lokal:**
   ```bash
   node server.js
   ```
2. **Sideloading ke Excel:**
   - Tambahkan `manifest.xml` melalui Developer Catalog atau Excel for Web (Upload My Add-in).
   - Buka Excel -> Tab **Home** -> Klik **"Buka Smart Assistant"**.
