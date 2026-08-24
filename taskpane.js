/**
 * EXCEL SMART ASSISTANT ADD-IN
 * File: taskpane.js
 * Deskripsi: Skrip logika utama untuk antarmuka Task Pane Excel menggunakan Office.js API.
 */

/* ==========================================================================
   1. INISIALISASI OFFICE.JS & EVENT LISTENERS
   ========================================================================== */

/**
 * Office.onReady memastikan pustaka Office.js telah dimuat
 * dan siap berinteraksi dengan buku kerja Excel sebelum event listener didaftarkan.
 */
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Office.js siap untuk Excel.");

        // Muat API Key dari penyimpanan lokal jika pernah disimpan
        initSettings();

        // 1. AI Formula Assistant
        document.getElementById("btn-generate-formula")?.addEventListener("click", handleGenerateFormula);

        // 2. Smart Data Cleaner
        document.getElementById("btn-clean-trim")?.addEventListener("click", () => handleCleanData("trim"));
        document.getElementById("btn-clean-proper")?.addEventListener("click", () => handleCleanData("proper"));
        document.getElementById("btn-clean-upper")?.addEventListener("click", () => handleCleanData("upper"));
        document.getElementById("btn-clean-lower")?.addEventListener("click", () => handleCleanData("lower"));
        document.getElementById("btn-format-rupiah")?.addEventListener("click", handleFormatRupiah);
        document.getElementById("btn-format-date")?.addEventListener("click", handleFormatDate);

        // 3. Template Generator
        document.getElementById("btn-generate-template")?.addEventListener("click", handleGenerateTemplate);

        // 4. Mass Action: Split Sheets
        document.getElementById("btn-split-sheets")?.addEventListener("click", handleSplitDataToSheets);

        // Pengaturan API Key (Toggle & Simpan)
        document.getElementById("btn-toggle-settings")?.addEventListener("click", toggleSettingsCard);
        document.getElementById("btn-save-key")?.addEventListener("click", saveApiKey);
    }
});

/* ==========================================================================
   2. PENGATURAN API KEY (GEMINI AI CONFIG)
   ========================================================================== */

/**
 * Memuat API Key yang tersimpan di localStorage browser.
 */
function initSettings() {
    const savedKey = localStorage.getItem("GEMINI_API_KEY");
    const keyInput = document.getElementById("gemini-api-key");
    if (savedKey && keyInput) {
        keyInput.value = savedKey;
    }
}

/**
 * Menampilkan atau menyembunyikan kartu pengaturan API Key.
 */
function toggleSettingsCard() {
    const card = document.getElementById("settings-card");
    if (card) {
        const isVisible = card.style.display === "block";
        card.style.display = isVisible ? "none" : "block";
    }
}

/**
 * Menyimpan API Key ke localStorage agar tidak hilang saat add-in dimuat ulang.
 */
function saveApiKey() {
    const keyInput = document.getElementById("gemini-api-key");
    const apiKey = keyInput ? keyInput.value.trim() : "";
    
    if (apiKey) {
        localStorage.setItem("GEMINI_API_KEY", apiKey);
        showStatus("API Key berhasil disimpan!", "success");
    } else {
        localStorage.removeItem("GEMINI_API_KEY");
        showStatus("API Key dihapus. Sistem menggunakan parser bawaan.", "info");
    }
    toggleSettingsCard();
}

/* ==========================================================================
   3. FITUR 1: AI FORMULA ASSISTANT & EXPLAINER (GEMINI + FALLBACK)
   ========================================================================== */

/**
 * Mengubah prompt bahasa alami menjadi formula Excel menggunakan Gemini API
 * atau parser lokal bawaan, lalu memasukkan formula ke sel yang sedang aktif.
 */
async function handleGenerateFormula() {
    const inputElement = document.getElementById("formula-input");
    const userPrompt = inputElement ? inputElement.value.trim() : "";

    if (!userPrompt) {
        showStatus("Silakan masukkan instruksi formula terlebih dahulu!", "warning");
        return;
    }

    const apiKey = localStorage.getItem("GEMINI_API_KEY");
    showStatus("Sedang menyusun formula...", "info");

    let formulaResult = "";

    try {
        if (apiKey) {
            // Gunakan integrasi Google Gemini AI jika API Key tersedia
            formulaResult = await fetchFormulaFromGemini(userPrompt, apiKey);
        } else {
            // Gunakan smart parser lokal jika tanpa API Key
            formulaResult = parseNaturalLanguageToFormula(userPrompt);
        }

        if (!formulaResult) {
            formulaResult = parseNaturalLanguageToFormula(userPrompt);
        }

        // Tuliskan formula ke sel aktif di Excel
        await Excel.run(async (context) => {
            const activeCell = context.workbook.getActiveCell();
            activeCell.formulas = [[formulaResult]];
            activeCell.select();
            await context.sync();

            showStatus(`Formula berhasil dibuat: ${formulaResult}`, "success");
        });
    } catch (error) {
        console.error("Gagal membuat formula:", error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Helper: Memanggil Google Gemini API untuk menghasilkan formula Excel murni.
 * 
 * @param {string} prompt - Instruksi dari pengguna
 * @param {string} apiKey - Google Gemini API Key
 * @returns {Promise<string>} Formula Excel murni
 */
async function fetchFormulaFromGemini(prompt, apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    
    const requestBody = {
        contents: [
            {
                parts: [
                    {
                        text: `Kamu adalah asisten formula Microsoft Excel. Ubah instruksi bahasa alami berikut menjadi SATU formula Excel yang valid dan siap pakai.
Aturan:
1. Hanya balas dengan satu string formula Excel lengkap yang diawali tanda '=' (contoh: =SUM(A1:A10) atau =IF(A1>10, "Besar", "Kecil")).
2. Jangan berikan teks pembuka, penutup, atau tanda markdown backtick.

Instruksi pengguna: "${prompt}"`
                    }
                ]
            }
        ]
    };

    const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
        throw new Error(`Gemini API Error: ${response.statusText}`);
    }

    const data = await response.json();
    let text = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "";
    
    // Bersihkan karakter markdown jika model menyertakan backticks
    text = text.replace(/```[a-z]*\n?/gi, "").replace(/```/g, "").trim();
    if (!text.startsWith("=")) {
        text = "=" + text;
    }
    return text;
}

/**
 * Helper: Parser bawaan lokal untuk menerjemahkan kata kunci bahasa Indonesia menjadi rumus Excel.
 * 
 * @param {string} prompt - Instruksi dari pengguna
 * @returns {string} Formula Excel lengkap (diawali tanda '=')
 */
function parseNaturalLanguageToFormula(prompt) {
    const text = prompt.toLowerCase();

    // 1. SUM / Penjumlahan
    if (text.includes("jumlah") || text.includes("total") || text.includes("sum")) {
        const match = text.match(/([a-z]+[0-9]+)\s*(?:sampai|hingga|to|-|:)\s*([a-z]+[0-9]+)/i);
        if (match) {
            return `=SUM(${match[1].toUpperCase()}:${match[2].toUpperCase()})`;
        }
        return "=SUM(A1:A10)";
    }

    // 2. AVERAGE / Rata-rata
    if (text.includes("rata-rata") || text.includes("average") || text.includes("mean")) {
        const match = text.match(/([a-z]+[0-9]+)\s*(?:sampai|hingga|to|-|:)\s*([a-z]+[0-9]+)/i);
        if (match) {
            return `=AVERAGE(${match[1].toUpperCase()}:${match[2].toUpperCase()})`;
        }
        return "=AVERAGE(A1:A10)";
    }

    // 3. COUNT / Menghitung Data
    if (text.includes("hitung jumlah data") || text.includes("banyak data") || text.includes("count")) {
        return "=COUNT(A1:A10)";
    }

    // 4. IF / Logika Kondisional
    if (text.includes("jika") || text.includes("if")) {
        if (text.includes("lulus")) {
            return '=IF(A1>=75, "Lulus", "Tidak Lulus")';
        }
        return '=IF(A1>0, "Positif", "Negatif")';
    }

    // 5. VLOOKUP & XLOOKUP
    if (text.includes("vlookup") || text.includes("cari data")) {
        return '=VLOOKUP(A2, Sheet2!A:B, 2, FALSE)';
    }

    if (text.includes("xlookup")) {
        return '=XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B, "Tidak Ditemukan")';
    }

    // Default jika pengguna langsung mengetikkan formula
    if (prompt.startsWith("=")) {
        return prompt;
    }

    return `=CONCATENATE("${prompt} - ", A1)`;
}

/* ==========================================================================
   4. FITUR 2: SMART DATA CLEANER (1-CLICK CLEANER)
   ========================================================================== */

/**
 * Membersihkan data teks pada rentang sel yang dipilih.
 * Mendukung mode: 'trim', 'proper', 'upper', 'lower'.
 * 
 * @param {'trim' | 'proper' | 'upper' | 'lower'} mode - Jenis pembersihan data
 */
async function handleCleanData(mode) {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load("values");
            await context.sync();

            const originalValues = selectedRange.values;
            let modifiedCount = 0;

            const cleanedValues = originalValues.map((row) =>
                row.map((cellValue) => {
                    if (typeof cellValue === "string") {
                        if (mode === "trim") {
                            const trimmed = cellValue.trim().replace(/\s+/g, " ");
                            if (trimmed !== cellValue) modifiedCount++;
                            return trimmed;
                        } else if (mode === "proper") {
                            const proper = toProperCase(cellValue);
                            if (proper !== cellValue) modifiedCount++;
                            return proper;
                        } else if (mode === "upper") {
                            const upper = cellValue.toUpperCase();
                            if (upper !== cellValue) modifiedCount++;
                            return upper;
                        } else if (mode === "lower") {
                            const lower = cellValue.toLowerCase();
                            if (lower !== cellValue) modifiedCount++;
                            return lower;
                        }
                    }
                    return cellValue;
                })
            );

            selectedRange.values = cleanedValues;
            await context.sync();

            const labelMap = {
                trim: "Trim Spasi Ganda",
                proper: "Proper Case",
                upper: "UPPERCASE",
                lower: "lowercase"
            };

            showStatus(`Berhasil menjalankan ${labelMap[mode]}! ${modifiedCount} sel disesuaikan.`, "success");
        });
    } catch (error) {
        console.error(`Gagal membersihkan data (${mode}):`, error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Format sel angka menjadi format mata uang Rupiah Indonesia (Rp #,##0).
 */
async function handleFormatRupiah() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            // Terapkan format angka akuntansi Rupiah standar Excel
            selectedRange.numberFormat = [["_(\"Rp\"* #,##0_);_(\"Rp\"* (#,##0);_(\"Rp\"* \"-\"_);_(@_)"]];
            await context.sync();
            showStatus("Format Mata Uang Rupiah berhasil diterapkan!", "success");
        });
    } catch (error) {
        console.error("Gagal menerapkan format Rupiah:", error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Format sel tanggal menjadi format standar DD/MM/YYYY.
 */
async function handleFormatDate() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.numberFormat = [["DD/MM/YYYY"]];
            await context.sync();
            showStatus("Format Tanggal (DD/MM/YYYY) berhasil diterapkan!", "success");
        });
    } catch (error) {
        console.error("Gagal menerapkan format tanggal:", error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Helper: Mengubah string menjadi format Proper Case (Title Case).
 * 
 * @param {string} str - Teks input
 * @returns {string} Teks format Proper Case
 */
function toProperCase(str) {
    return str
        .toLowerCase()
        .replace(/\b\w/g, (char) => char.toUpperCase());
}

/* ==========================================================================
   5. FITUR 3: TEMPLATE GENERATOR (EXPANDED TEMPLATES)
   ========================================================================== */

/**
 * Membuat tabel template otomatis pada sheet aktif
 * lengkap dengan styling hijau Fluent, border, formula, dan validasi data.
 */
async function handleGenerateTemplate() {
    const selectElement = document.getElementById("template-select");
    const templateType = selectElement ? selectElement.value : "kinerja";

    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const templateConfig = getTemplateConfiguration(templateType);

            const totalCols = templateConfig.headers.length;
            const totalRows = templateConfig.sampleRows.length + 1;
            
            const tableRange = worksheet.getRangeByIndexes(0, 0, totalRows, totalCols);
            const allData = [templateConfig.headers, ...templateConfig.sampleRows];
            tableRange.values = allData;

            // 1. Format Header Baris Pertama
            const headerRange = worksheet.getRangeByIndexes(0, 0, 1, totalCols);
            headerRange.format.fill.color = "#107c41"; // Hijau Excel Khas Microsoft Fluent
            headerRange.format.font.color = "#FFFFFF"; // Font Putih
            headerRange.format.font.bold = true;
            headerRange.format.horizontalAlignment = "Center";

            // 2. Format Garis Border
            tableRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
            tableRange.format.borders.getItem("InsideHorizontal").color = "#D3D3D3";
            tableRange.format.borders.getItem("InsideVertical").style = "Continuous";
            tableRange.format.borders.getItem("InsideVertical").color = "#D3D3D3";
            tableRange.format.borders.getItem("EdgeBottom").style = "Continuous";
            tableRange.format.borders.getItem("EdgeBottom").color = "#107c41";

            // 3. Auto-fit lebar kolom
            tableRange.format.autofitColumns();

            // 4. Freeze Panes baris atas
            worksheet.freezePanes.freezeRows(1);

            // 5. Data Validation Dropdown jika ada
            if (templateConfig.dropdownValidation) {
                const { columnIndex, listFormula } = templateConfig.dropdownValidation;
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
 * Helper: Skema konfigurasi header dan contoh baris data untuk semua template.
 * 
 * @param {string} type - Identifier jenis template
 * @returns {object} Konfigurasi template
 */
function getTemplateConfiguration(type) {
    switch (type) {
        case "arus_kas":
            return {
                title: "Laporan Arus Kas (Cash Flow)",
                headers: ["Tanggal", "Deskripsi Transaksi", "Kategori", "Pemasukan (Debit)", "Pengeluaran (Kredit)", "Saldo", "Keterangan"],
                sampleRows: [
                    ["01/08/2026", "Modal Awal Operasional", "Modal", 50000000, 0, "=D2-E2", "Kas Utama"],
                    ["05/08/2026", "Penjualan Produk Paket A", "Penjualan", 7500000, 0, "=F2+D3-E3", "Transfer Bank"],
                    ["10/08/2026", "Pembayaran Sewa Server", "Operasional", 0, 1200000, "=F3+D4-E4", "Tagihan Bulanan"]
                ],
                dropdownValidation: {
                    columnIndex: 2, // Kolom Kategori (Index 2 / Kolom C)
                    listFormula: "Modal, Penjualan, Operasional, Pemasaran, Gaji, Lain-lain"
                }
            };

        case "jadwal_proyek":
            return {
                title: "Jadwal Proyek / Project Tracker",
                headers: ["No", "Nama Tugas / Task", "Penanggung Jawab", "Tanggal Mulai", "Target Selesai", "Status", "Prioritas", "Catatan"],
                sampleRows: [
                    [1, "Riset Kebutuhan Pengguna", "Budi Santoso", "01/08/2026", "07/08/2026", "Selesai", "Tinggi", "Hasil riset valid"],
                    [2, "Desain UI & Arsitektur Add-in", "Dewi Lestari", "08/08/2026", "15/08/2026", "Sedang Berjalan", "Tinggi", "Desain Fluent UI"],
                    [3, "Pengujian Fitur & Sideloading", "Ahmad Fauzi", "16/08/2026", "22/08/2026", "Belum Dimulai", "Sedang", "Menunggu tahap coding"]
                ],
                dropdownValidation: {
                    columnIndex: 5, // Kolom Status (Index 5 / Kolom F)
                    listFormula: "Belum Dimulai, Sedang Berjalan, Dalam Review, Selesai, Tertunda"
                }
            };

        case "invoice":
            return {
                title: "Daftar Tagihan / Invoice",
                headers: ["No", "Deskripsi Barang / Layanan", "Qty", "Harga Satuan", "Total Harga", "Status Pembayaran"],
                sampleRows: [
                    [1, "Jasa Pengembangan Add-In Excel", 1, 15000000, "=C2*D2", "Lunas"],
                    [2, "Maintenance & Cloud Hosting Server", 12, 500000, "=C3*D3", "Belum Bayar"]
                ],
                dropdownValidation: {
                    columnIndex: 5, // Kolom Status Pembayaran (Index 5 / Kolom F)
                    listFormula: "Lunas, Belum Bayar, Uang Muka (DP), Dibatalkan"
                }
            };

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
                    columnIndex: 5,
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
                    columnIndex: 6,
                    listFormula: "Sangat Baik, Baik, Cukup, Kurang, Perlu Peningkatan"
                }
            };
    }
}

/* ==========================================================================
   6. FITUR 4: MASS ACTION (SPLIT DATA KE BANYAK SHEET)
   ========================================================================== */

/**
 * Membaca data pada tabel yang diblok pengguna,
 * memisahkan data berdasarkan kategori unik pada kolom pertama,
 * dan membuat sheet baru secara otomatis untuk setiap kategori.
 */
async function handleSplitDataToSheets() {
    try {
        await Excel.run(async (context) => {
            const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
            const selectedRange = context.workbook.getSelectedRange();
            
            selectedRange.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
            const workbookWorksheets = context.workbook.worksheets;
            workbookWorksheets.load("items/name");

            await context.sync();

            const values = selectedRange.values;

            if (!values || values.length < 2) {
                showStatus("Pilih minimal 2 baris data (termasuk baris header) untuk di-split!", "warning");
                return;
            }

            const headers = values[0];
            const dataRows = values.slice(1);
            const existingSheetNames = new Set(workbookWorksheets.items.map((ws) => ws.name.toLowerCase()));

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

            for (const [categoryName, rows] of categoryMap.entries()) {
                const sanitizedSheetName = sanitizeSheetName(categoryName);
                
                let targetSheet;
                if (!existingSheetNames.has(sanitizedSheetName.toLowerCase())) {
                    targetSheet = workbookWorksheets.add(sanitizedSheetName);
                    existingSheetNames.add(sanitizedSheetName.toLowerCase());
                    createdSheetCount++;
                } else {
                    targetSheet = workbookWorksheets.getItem(sanitizedSheetName);
                }

                const outputData = [headers, ...rows];
                const targetRange = targetSheet.getRangeByIndexes(0, 0, outputData.length, headers.length);
                targetRange.values = outputData;

                const headerRange = targetSheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.format.fill.color = "#107c41";
                headerRange.format.font.color = "#FFFFFF";
                headerRange.format.font.bold = true;
                
                targetRange.format.autofitColumns();
            }

            await context.sync();
            showStatus(`Selesai! Berhasil memisahkan data ke ${categoryMap.size} kategori (${createdSheetCount} sheet baru dibuat).`, "success");
        });
    } catch (error) {
        console.error("Gagal melakukan Split Data ke Sheets:", error);
        showStatus(`Terjadi kesalahan: ${error.message}`, "danger");
    }
}

/**
 * Helper: Membersihkan nama sheet agar valid di Excel (maksimal 31 karakter).
 * 
 * @param {string} name - Nama kategori mentah
 * @returns {string} Nama sheet yang valid
 */
function sanitizeSheetName(name) {
    const cleaned = name.replace(/[\\/?*:[\]]/g, "_").trim();
    return cleaned.substring(0, 31) || "Sheet_Kategori";
}

/* ==========================================================================
   7. HELPER FEEDBACK STATUS UI
   ========================================================================== */

/**
 * Menampilkan pesan status aksi pada kotak notifikasi UI Task Pane.
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

        setTimeout(() => {
            if (statusBox) {
                statusBox.style.display = "none";
            }
        }, 5000);
    } else {
        console.log(`[Status - ${type.toUpperCase()}]: ${message}`);
    }
}
