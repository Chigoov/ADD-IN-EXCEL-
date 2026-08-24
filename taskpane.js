/**
 * EXCEL SMART ASSISTANT ADD-IN
 * File: taskpane.js
 * Deskripsi: Skrip logika utama untuk antarmuka Task Pane Excel menggunakan Office.js API.
 */

/* ==========================================================================
   1. INISIALISASI OFFICE.JS & EVENT LISTENERS
   ========================================================================== */

/**
 * Office.onReady memastikan seluruh pustaka Office.js telah dimuat
 * dan siap berinteraksi dengan buku kerja Excel sebelum event listener didaftarkan.
 */
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Office.js siap untuk Excel.");

        // Daftarkan Event Listener untuk masing-masing tombol aksi
        document.getElementById("btn-generate-formula")?.addEventListener("click", handleGenerateFormula);
        document.getElementById("btn-clean-trim")?.addEventListener("click", () => handleCleanData("trim"));
        document.getElementById("btn-clean-proper")?.addEventListener("click", () => handleCleanData("proper"));
        document.getElementById("btn-generate-template")?.addEventListener("click", handleGenerateTemplate);
        document.getElementById("btn-split-sheets")?.addEventListener("click", handleSplitDataToSheets);
    }
});

/* ==========================================================================
   2. FITUR 1: AI FORMULA ASSISTANT & EXPLAINER
   ========================================================================== */

/**
 * Mengubah instruksi bahasa alami dari textarea menjadi formula Excel,
 * kemudian memasukkannya langsung ke sel yang sedang aktif.
 */
async function handleGenerateFormula() {
    const inputElement = document.getElementById("formula-input");
    const userPrompt = inputElement ? inputElement.value.trim() : "";

    if (!userPrompt) {
        showStatus("Silakan masukkan instruksi formula terlebih dahulu!", "warning");
        return;
    }

    try {
        await Excel.run(async (context) => {
            // Dapatkan sel yang saat ini sedang aktif dipilih oleh pengguna
            const activeCell = context.workbook.getActiveCell();

            // Konversi teks bahasa alami menjadi formula Excel (fungsi parser/mock AI)
            const generatedFormula = parseNaturalLanguageToFormula(userPrompt);

            // Masukkan formula ke dalam sel aktif
            activeCell.formulas = [[generatedFormula]];
            
            // Pilih kembali sel aktif untuk memicu pembaruan tampilan
            activeCell.select();

            await context.sync();
            showStatus(`Formula berhasil dibuat: ${generatedFormula}`, "success");
        });
    } catch (error) {
        console.error("Gagal membuat formula:", error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Helper: Mengubah instruksi teks alami menjadi formula Excel.
 * Mendukung logika umum atau memberikan formula berbasis pola kata kunci.
 * (Dapat diintegrasikan dengan API AI seperti OpenAI / Gemini / endpoint kustom).
 * 
 * @param {string} prompt - Instruksi dari pengguna
 * @returns {string} Formula Excel lengkap (diawali tanda '=')
 */
function parseNaturalLanguageToFormula(prompt) {
    const text = prompt.toLowerCase();

    if (text.includes("jumlah") || text.includes("total") || text.includes("sum")) {
        // Deteksi rentang sel jika disebutkan (misal: "jumlahkan A1 sampai A10")
        const match = text.match(/([a-z]+[0-9]+)\s*(?:sampai|hingga|to|-|:)\s*([a-z]+[0-9]+)/i);
        if (match) {
            return `=SUM(${match[1].toUpperCase()}:${match[2].toUpperCase()})`;
        }
        return "=SUM(A1:A10)";
    }

    if (text.includes("rata-rata") || text.includes("average") || text.includes("mean")) {
        const match = text.match(/([a-z]+[0-9]+)\s*(?:sampai|hingga|to|-|:)\s*([a-z]+[0-9]+)/i);
        if (match) {
            return `=AVERAGE(${match[1].toUpperCase()}:${match[2].toUpperCase()})`;
        }
        return "=AVERAGE(A1:A10)";
    }

    if (text.includes("hitung jumlah data") || text.includes("banyak data") || text.includes("count")) {
        return "=COUNT(A1:A10)";
    }

    if (text.includes("jika") || text.includes("if")) {
        if (text.includes("lulus")) {
            return '=IF(A1>=75, "Lulus", "Tidak Lulus")';
        }
        return '=IF(A1>0, "Positif", "Negatif")';
    }

    if (text.includes("vlookup") || text.includes("cari data")) {
        return '=VLOOKUP(A2, Sheet2!A:B, 2, FALSE)';
    }

    if (text.includes("xlookup")) {
        return '=XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B, "Tidak Ditemukan")';
    }

    // Default jika pengguna langsung mengetikkan formula atau teks lainnya
    if (prompt.startsWith("=")) {
        return prompt;
    }

    return `=CONCATENATE("${prompt} - ", A1)`;
}

/* ==========================================================================
   3. FITUR 2: SMART DATA CLEANER (1-CLICK CLEANER)
   ========================================================================== */

/**
 * Membersihkan data pada rentang sel yang sedang diblok oleh pengguna.
 * Pilihan mode: 'trim' (hapus spasi berlebih) atau 'proper' (kapitalisasi tiap kata).
 * 
 * @param {'trim' | 'proper'} mode - Jenis pembersihan data
 */
async function handleCleanData(mode) {
    try {
        await Excel.run(async (context) => {
            // Dapatkan rentang sel yang sedang diblok pengguna
            const selectedRange = context.workbook.getSelectedRange();

            // Muat nilai matriks 2D dari rentang sel tersebut
            selectedRange.load("values");
            await context.sync();

            const originalValues = selectedRange.values;
            let modifiedCount = 0;

            // Proses iterasi array 2D untuk membersihkan setiap sel
            const cleanedValues = originalValues.map((row) =>
                row.map((cellValue) => {
                    if (typeof cellValue === "string") {
                        if (mode === "trim") {
                            // Hapus spasi di awal/akhir dan rapikan spasi ganda di dalam kalimat
                            const trimmed = cellValue.trim().replace(/\s+/g, " ");
                            if (trimmed !== cellValue) modifiedCount++;
                            return trimmed;
                        } else if (mode === "proper") {
                            // Ubah huruf pertama setiap kata menjadi huruf kapital
                            const proper = toProperCase(cellValue);
                            if (proper !== cellValue) modifiedCount++;
                            return proper;
                        }
                    }
                    return cellValue;
                })
            );

            // Tulis kembali data array yang sudah dibersihkan ke Excel
            selectedRange.values = cleanedValues;
            await context.sync();

            const actionLabel = mode === "trim" ? "Trim (Hapus Spasi Ganda)" : "Proper Case";
            showStatus(`Berhasil menjalankan ${actionLabel}! ${modifiedCount} sel disesuaikan.`, "success");
        });
    } catch (error) {
        console.error(`Gagal membersihkan data (${mode}):`, error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Helper: Mengubah string menjadi format Proper Case (Title Case).
 * Contoh: "laporan KEUANGAN divisi a" -> "Laporan Keuangan Divisi A"
 * 
 * @param {string} str - Teks input
 * @returns {string} Teks dalam format Proper Case
 */
function toProperCase(str) {
    return str
        .toLowerCase()
        .replace(/\b\w/g, (char) => char.toUpperCase());
}

/* ==========================================================================
   4. FITUR 3: TEMPLATE GENERATOR (MATRIKS OTOMATIS)
   ========================================================================== */

/**
 * Membuat struktur tabel template otomatis pada sheet aktif
 * lengkap dengan header berwarna, formatting, dan validasi data.
 */
async function handleGenerateTemplate() {
    const selectElement = document.getElementById("template-select");
    const templateType = selectElement ? selectElement.value : "kinerja";

    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();

            // Ambil konfigurasi header dan contoh baris berdasarkan template
            const templateConfig = getTemplateConfiguration(templateType);

            // Tentukan rentang sel untuk header dan data awal
            const totalCols = templateConfig.headers.length;
            const totalRows = templateConfig.sampleRows.length + 1; // Termasuk header
            
            // Rentang mulai dari sel A1
            const tableRange = worksheet.getRangeByIndexes(0, 0, totalRows, totalCols);
            
            // Masukkan data (Header di baris 0, diikuti baris contoh)
            const allData = [templateConfig.headers, ...templateConfig.sampleRows];
            tableRange.values = allData;

            // Format Header (Baris pertama)
            const headerRange = worksheet.getRangeByIndexes(0, 0, 1, totalCols);
            headerRange.format.fill.color = "#107c41"; // Warna Hijau Excel Khas Microsoft Fluent
            headerRange.format.font.color = "#FFFFFF"; // Font Putih
            headerRange.format.font.bold = true;
            headerRange.format.horizontalAlignment = "Center";

            // Format seluruh tabel (Border garis tipis dan perataan vertikal)
            tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
            tableRange.format.borders.getItem("InsideHorizontal").color = "#D3D3D3";
            tableRange.format.borders.getItem("InsideVertical").style = "Continuous";
            tableRange.format.borders.getItem("InsideVertical").color = "#D3D3D3";
            tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
            tableRange.format.borders.getItem("EdgeBottom").color = "#107c41";

            // Auto-fit lebar kolom agar teks terlihat rapi
            tableRange.format.autofitColumns();

            // Bekukan (Freeze Panes) baris header agar tetap terlihat saat scroll
            worksheet.freezePanes.freezeRows(1);

            // Tambahkan Data Validation dropdown jika didefinisikan pada template
            if (templateConfig.dropdownValidation) {
                const { columnIndex, listFormula } = templateConfig.dropdownValidation;
                // Terapkan ke 100 baris data pertama
                const validationRange = worksheet.getRangeByIndexes(1, columnIndex, 100, 1);
                validationRange.dataValidation.rule = {
                    list: {
                        inCellDropDown: true,
                        source: listFormula
                    }
                };
            }

            await context.sync();
            showStatus(`Template "${templateConfig.title}" berhasil dibuat!`, "success");
        });
    } catch (error) {
        console.error("Gagal membuat template:", error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Helper: Menyediakan skema konfigurasi header dan contoh data per template.
 * 
 * @param {string} type - Identifier jenis template
 * @returns {object} Konfigurasi template
 */
function getTemplateConfiguration(type) {
    switch (type) {
        case "media":
        case "media_partner":
            return {
                title: "Media Partner Tracker",
                headers: ["No", "Nama Media", "Kategori Media", "Kontak Person", "Email / No. Telp", "Status MoU", "Biaya / Paket", "Keterangan"],
                sampleRows: [
                    [1, "Tech Daily ID", "Online News", "Budi Santoso", "budi@techdaily.id", "Disetujui", 2500000, "Tayang H-3 Event"],
                    [2, "Kampus Update", "Instagram", "Siti Rahma", "081234567890", "Draft", 1000000, "Menunggu revisi proposal"]
                ],
                dropdownValidation: {
                    columnIndex: 5, // Kolom Status MoU (Index ke-5 / Kolom F)
                    listFormula: "Draft, Diskusi, Disetujui, Selesai, Dibatalkan"
                }
            };

        case "literatur":
        case "tinjauan_pustaka":
            return {
                title: "Tinjauan Pustaka (Literature Review)",
                headers: ["No", "Judul Artikel / Buku", "Penulis", "Tahun", "Jurnal / Penerbit", "Metodologi", "Temuan Utama", "Relevansi Riset"],
                sampleRows: [
                    [1, "Deep Learning in Spreadsheet Analytics", "Smith et al.", 2024, "IEEE Access", "Kuantitatif - Eksperimen", "Akurasi model mencapai 98%", "Sangat Relevan"],
                    [2, "Productivity in Office Workflow", "Johnson & Lee", 2023, "Harvard Business Review", "Kualitatif - Survei", "Add-in menghemat 30% waktu kerja", "Relevan"]
                ],
                dropdownValidation: null
            };

        case "kinerja":
        case "evaluasi_kinerja":
        default:
            return {
                title: "Evaluasi Kinerja Karyawan",
                headers: ["No", "Nama Karyawan", "Divisi", "Target KPI", "Capaian KPI", "Persentase", "Status Evaluasi", "Catatan"],
                sampleRows: [
                    [1, "Andi Pratama", "Teknologi", 100, 95, "=E2/D2", "Baik", "Performa stabil dan konsisten"],
                    [2, "Dewi Lestari", "Pemasaran", 50, 55, "=E3/D3", "Sangat Baik", "Melampaui target bulanan"]
                ],
                dropdownValidation: {
                    columnIndex: 6, // Kolom Status Evaluasi (Index ke-6 / Kolom G)
                    listFormula: "Sangat Baik, Baik, Cukup, Kurang, Perlu Peningkatan"
                }
            };
    }
}

/* ==========================================================================
   5. FITUR 4: MASS ACTION (SPLIT DATA KE BANYAK SHEET)
   ========================================================================== */

/**
 * Membaca data pada tabel atau kolom yang dipilih pengguna,
 * memisahkan baris data berdasarkan nilai kategori unik,
 * dan membuatkan Worksheet baru untuk setiap kategori secara otomatis.
 */
async function handleSplitDataToSheets() {
    try {
        await Excel.run(async (context) => {
            const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
            const selectedRange = context.workbook.getSelectedRange();
            
            // Muat data dan koordinat dari rentang yang dipilih
            selectedRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
            const workbookWorksheets = context.workbook.worksheets;
            workbookWorksheets.load("items/name");

            await context.sync();

            const values = selectedRange.values;

            if (!values || values.length < 2) {
                showStatus("Pilih minimal 2 baris data (termasuk baris header) untuk di-split!", "warning");
                return;
            }

            // Baris pertama dianggap sebagai Header
            const headers = values[0];
            const dataRows = values.slice(1);

            // Ambil nama sheet yang sudah ada untuk menghindari duplikasi error
            const existingSheetNames = new Set(workbookWorksheets.items.map((ws) => ws.name.toLowerCase()));

            // Cari kategori unik dari kolom pertama rentang yang dipilih
            const categoryMap = new Map();

            dataRows.forEach((row) => {
                const categoryRaw = row[0];
                if (categoryRaw !== undefined && categoryRaw !== null && String(categoryRaw).trim() !== "") {
                    const categoryKey = String(categoryRaw).trim();
                    if (!categoryMap.has(categoryKey)) {
                        categoryMap.set(categoryKey, []);
                    }
                    categoryMap.get(categoryKey).push(row);
                }
            });

            if (categoryMap.size === 0) {
                showStatus("Tidak ditemukan data kategori unik pada kolom yang diblok.", "warning");
                return;
            }

            let createdSheetCount = 0;

            // Iterasi untuk setiap kategori unik dan buat sheet baru
            for (const [categoryName, rows] of categoryMap.entries()) {
                // Bersihkan karakter terlarang untuk nama worksheet Excel (: \ / ? * [ ]) dan batasi 31 karakter
                const sanitizedSheetName = sanitizeSheetName(categoryName);
                
                let targetSheet;
                if (!existingSheetNames.has(sanitizedSheetName.toLowerCase())) {
                    targetSheet = workbookWorksheets.add(sanitizedSheetName);
                    existingSheetNames.add(sanitizedSheetName.toLowerCase());
                    createdSheetCount++;
                } else {
                    targetSheet = workbookWorksheets.getItem(sanitizedSheetName);
                }

                // Tulis Header dan Data ke sheet baru
                const outputData = [headers, ...rows];
                const targetRange = targetSheet.getRangeByIndexes(0, 0, outputData.length, headers.length);
                targetRange.values = outputData;

                // Format Header Sheet Baru
                const headerRange = targetSheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.format.fill.color = "#107c41";
                headerRange.format.font.color = "#FFFFFF";
                headerRange.format.font.bold = true;
                
                // Auto-fit kolom di sheet baru
                targetRange.format.autofitColumns();
            }

            await context.sync();
            showStatus(`Selesai! Berhasil memisahkan data ke dalam ${categoryMap.size} kategori (${createdSheetCount} sheet baru dibuat).`, "success");
        });
    } catch (error) {
        console.error("Gagal melakukan Split Data ke Sheets:", error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Helper: Membersihkan nama sheet agar memenuhi syarat penamaan Worksheet Excel.
 * Karakter yang dilarang: \ / ? * : [ ] dan maksimal panjang 31 karakter.
 * 
 * @param {string} name - Nama kategori mentah
 * @returns {string} Nama sheet yang valid
 */
function sanitizeSheetName(name) {
    const cleaned = name.replace(/[\\/?*:[\]]/g, "_").trim();
    return cleaned.substring(0, 31) || "Sheet_Kategori";
}

/* ==========================================================================
   6. HELPER FEEDBACK STATUS UI
   ========================================================================== */

/**
 * Menampilkan pesan notifikasi status aksi pada UI Task Pane jika elemen status tersedia,
 * atau menampilkan pesan log dan alert sederhana.
 * 
 * @param {string} message - Pesan yang ditampilkan
 * @param {'success' | 'warning' | 'danger' | 'info'} type - Tipe notifikasi
 */
function showStatus(message, type = "info") {
    const statusBox = document.getElementById("status-message");
    if (statusBox) {
        statusBox.textContent = message;
        statusBox.className = `status-alert status-${type}`;
        statusBox.style.display = "block";

        // Sembunyikan otomatis setelah 5 detik
        setTimeout(() => {
            if (statusBox) {
                statusBox.style.display = "none";
            }
        }, 5000);
    } else {
        console.log(`[Status - ${type.toUpperCase()}]: ${message}`);
    }
}
