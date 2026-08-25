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

        // 1. AI Formula Assistant & Explainer
        document.getElementById("btn-generate-formula")?.addEventListener("click", handleGenerateFormula);
        document.getElementById("btn-explain-formula")?.addEventListener("click", handleExplainFormula);
        document.getElementById("btn-fix-formula-error")?.addEventListener("click", handleFixFormulaError);
        document.getElementById("btn-close-explain")?.addEventListener("click", closeFormulaExplainBox);

        // 2. Quick Statistics & Auto-Total
        document.getElementById("btn-calc-stats")?.addEventListener("click", handleCalculateQuickStats);
        document.getElementById("btn-insert-total-row")?.addEventListener("click", handleInsertTotalRow);

        // 3. Smart Data Cleaner
        document.getElementById("btn-clean-trim")?.addEventListener("click", () => handleCleanData("trim"));
        document.getElementById("btn-clean-proper")?.addEventListener("click", () => handleCleanData("proper"));
        document.getElementById("btn-clean-upper")?.addEventListener("click", () => handleCleanData("upper"));
        document.getElementById("btn-clean-lower")?.addEventListener("click", () => handleCleanData("lower"));
        document.getElementById("btn-format-rupiah")?.addEventListener("click", handleFormatRupiah);
        document.getElementById("btn-format-date")?.addEventListener("click", handleFormatDate);

        // 4. Smart Deduplication & Text Tools
        document.getElementById("btn-highlight-duplicates")?.addEventListener("click", handleHighlightDuplicates);
        document.getElementById("btn-remove-duplicates")?.addEventListener("click", handleRemoveDuplicates);
        document.getElementById("btn-extract-unique")?.addEventListener("click", handleExtractUnique);
        document.getElementById("btn-split-text")?.addEventListener("click", handleSplitTextColumn);
        document.getElementById("btn-merge-text")?.addEventListener("click", handleMergeTextColumns);

        // 5. 1-Click Auto Chart Generator
        document.getElementById("btn-chart-column")?.addEventListener("click", () => handleCreateChart(Excel.ChartType.columnClustered));
        document.getElementById("btn-chart-line")?.addEventListener("click", () => handleCreateChart(Excel.ChartType.line));
        document.getElementById("btn-chart-pie")?.addEventListener("click", () => handleCreateChart(Excel.ChartType.pie));

        // 6. Data Quality Auditor
        document.getElementById("btn-highlight-blanks")?.addEventListener("click", handleHighlightBlankCells);
        document.getElementById("btn-convert-text-numbers")?.addEventListener("click", handleConvertTextToNumbers);
        document.getElementById("btn-clear-format")?.addEventListener("click", handleClearFormatting);

        // 7. Template Generator
        document.getElementById("btn-generate-template")?.addEventListener("click", handleGenerateTemplate);

        // 8. Mass Action: Split Sheets
        document.getElementById("btn-split-sheets")?.addEventListener("click", handleSplitDataToSheets);

        // Pengaturan API Key (Toggle & Simpan)
        document.getElementById("btn-toggle-settings")?.addEventListener("click", toggleSettingsCard);
        document.getElementById("btn-save-key")?.addEventListener("click", saveApiKey);

        // Listener Otomatis saat Pengguna Memilih Rentang Sel (Auto Stats)
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            () => handleCalculateQuickStats(true),
            () => {}
        );
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
   3. FITUR 1: AI FORMULA ASSISTANT & EXPLAINER & ERROR FIXER
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
            formulaResult = await fetchFormulaFromGemini(userPrompt, apiKey);
        } else {
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
 * Membaca formula pada sel aktif dan menjelaskannya langkah demi langkah
 * menggunakan Google Gemini AI atau parser analisis lokal.
 */
async function handleExplainFormula() {
    try {
        await Excel.run(async (context) => {
            const activeCell = context.workbook.getActiveCell();
            activeCell.load(["formulas", "values", "address"]);
            await context.sync();

            const cellFormula = activeCell.formulas[0][0];
            const cellAddress = activeCell.address;

            if (!cellFormula || !String(cellFormula).startsWith("=")) {
                showStatus(`Sel ${cellAddress} tidak berisi rumus/formula!`, "warning");
                return;
            }

            showStatus("Menganalisis rumus dengan AI...", "info");

            const apiKey = localStorage.getItem("GEMINI_API_KEY");
            let explanation = "";

            if (apiKey) {
                explanation = await fetchFormulaExplanationFromGemini(cellFormula, apiKey);
            } else {
                explanation = explainFormulaLocally(cellFormula);
            }

            // Tampilkan hasil di box penjelasan
            const explainBox = document.getElementById("formula-explain-box");
            const formulaText = document.getElementById("explain-formula-text");
            const contentText = document.getElementById("explain-content");

            if (explainBox && formulaText && contentText) {
                formulaText.textContent = cellFormula;
                contentText.innerHTML = explanation.replace(/\n/g, "<br>");
                explainBox.style.display = "block";
            }

            showStatus("Penjelasan rumus berhasil dimuat!", "success");
        });
    } catch (error) {
        console.error("Gagal menjelaskan formula:", error);
        showStatus(`Gagal menjelaskan rumus: ${error.message}`, "danger");
    }
}

/**
 * Menutup kotak penjelasan formula.
 */
function closeFormulaExplainBox() {
    const explainBox = document.getElementById("formula-explain-box");
    if (explainBox) {
        explainBox.style.display = "none";
    }
}

/**
 * Otomatis memperbaiki rumus pada sel aktif dengan membungkusnya dalam IFERROR
 * untuk mencegah tampilan error seperti #N/A, #VALUE!, #DIV/0!, #REF!.
 */
async function handleFixFormulaError() {
    try {
        await Excel.run(async (context) => {
            const activeCell = context.workbook.getActiveCell();
            activeCell.load(["formulas", "values", "address"]);
            await context.sync();

            const currentFormula = activeCell.formulas[0][0];
            if (!currentFormula || !String(currentFormula).startsWith("=")) {
                showStatus("Sel aktif tidak memiliki formula untuk diperbaiki.", "warning");
                return;
            }

            const cleanFormula = currentFormula.replace(/^=/, "").trim();

            // Jika sudah ada IFERROR, jangan bungkus ganda
            if (cleanFormula.toUpperCase().startsWith("IFERROR(")) {
                showStatus("Formula sudah diproteksi dengan IFERROR.", "info");
                return;
            }

            const fixedFormula = `=IFERROR(${cleanFormula}, 0)`;
            activeCell.formulas = [[fixedFormula]];
            await context.sync();

            showStatus(`Rumus berhasil diproteksi: ${fixedFormula}`, "success");
        });
    } catch (error) {
        console.error("Gagal memperbaiki formula:", error);
        showStatus(`Gagal memperbaiki rumus: ${error.message}`, "danger");
    }
}

/**
 * Helper: Memanggil Google Gemini API untuk menghasilkan formula Excel murni.
 */
async function fetchFormulaFromGemini(prompt, apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    
    const requestBody = {
        contents: [
            {
                parts: [
                    {
                        text: `Kamu adalah asisten formula Microsoft Excel profesional. Ubah instruksi berikut menjadi SATU formula Excel valid dan siap pakai.
Aturan:
1. Hanya balas dengan satu string formula Excel lengkap yang diawali tanda '=' (contoh: =SUM(A1:A10) atau =XLOOKUP(A2,Sheet2!A:A,Sheet2!B:B,"N/A")).
2. Jangan berikan teks pembuka, penutup, atau tanda markdown backtick.

Instruksi: "${prompt}"`
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
        throw new Error(`Gemini API: ${response.statusText}`);
    }

    const data = await response.json();
    let text = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "";
    text = text.replace(/```[a-z]*\n?/gi, "").replace(/```/g, "").trim();
    if (!text.startsWith("=")) {
        text = "=" + text;
    }
    return text;
}

/**
 * Helper: Memanggil Google Gemini API untuk menjelaskan alur rumus Excel.
 */
async function fetchFormulaExplanationFromGemini(formula, apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    
    const requestBody = {
        contents: [
            {
                parts: [
                    {
                        text: `Jelaskan fungsi dan cara kerja rumus Microsoft Excel berikut dalam Bahasa Indonesia yang ringkas, mudah dipahami, dan jelas poin-poinnya (maksimal 3-4 kalimat ringkas):
Formula: "${formula}"`
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
        return explainFormulaLocally(formula);
    }

    const data = await response.json();
    return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || explainFormulaLocally(formula);
}

/**
 * Helper: Penjelasan rumus lokal sederhana jika tanpa koneksi Gemini API.
 */
function explainFormulaLocally(formula) {
    const upper = formula.toUpperCase();
    let steps = [];

    if (upper.includes("SUM(")) steps.push("• <b>SUM</b>: Menjumlahkan seluruh angka dalam rentang sel yang dipilih.");
    if (upper.includes("AVERAGE(")) steps.push("• <b>AVERAGE</b>: Menghitung nilai rata-rata dari rentang angka.");
    if (upper.includes("IF(")) steps.push("• <b>IF</b>: Memeriksa suatu kondisi logika dan memberikan hasil berbeda jika Benar atau Salah.");
    if (upper.includes("VLOOKUP(") || upper.includes("XLOOKUP(")) steps.push("• <b>LOOKUP</b>: Mencari nilai kunci pada kolom referensi dan mengambil data kolom yang sesuai.");
    if (upper.includes("COUNT(") || upper.includes("COUNTA(")) steps.push("• <b>COUNT</b>: Menghitung banyaknya sel yang terisi data.");
    if (upper.includes("IFERROR(")) steps.push("• <b>IFERROR</b>: Menangkap potensi error dan menggantinya dengan nilai default aman.");

    if (steps.length === 0) {
        steps.push(`• Rumus ini mengevaluasi ekspresi matematika atau fungsi kustom pada lembar kerja Excel.`);
    }

    return steps.join("<br>");
}

/**
 * Helper: Parser bawaan lokal untuk menerjemahkan kata kunci bahasa Indonesia menjadi rumus Excel.
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
   4. FITUR 2: QUICK STATISTICS & AUTO-TOTAL
   ========================================================================== */

/**
 * Menghitung statistik instan dari rentang sel yang diblok pengguna.
 */
async function handleCalculateQuickStats(silent = false) {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const values = range.values;
            let sum = 0;
            let countNumbers = 0;
            let countTotal = 0;
            let min = Infinity;
            let max = -Infinity;
            const uniqueSet = new Set();

            for (let r = 0; r < values.length; r++) {
                for (let c = 0; c < values[r].length; c++) {
                    const val = values[r][c];
                    if (val !== "" && val !== null && val !== undefined) {
                        countTotal++;
                        uniqueSet.add(String(val).trim());

                        const num = Number(val);
                        if (!isNaN(num) && typeof val !== "boolean") {
                            sum += num;
                            countNumbers++;
                            if (num < min) min = num;
                            if (num > max) max = num;
                        }
                    }
                }
            }

            const statSumEl = document.getElementById("stat-sum");
            const statAvgEl = document.getElementById("stat-avg");
            const statMinMaxEl = document.getElementById("stat-minmax");
            const statCountEl = document.getElementById("stat-count");

            if (countNumbers > 0) {
                const avg = sum / countNumbers;
                if (statSumEl) statSumEl.textContent = formatNumberDisplay(sum);
                if (statAvgEl) statAvgEl.textContent = formatNumberDisplay(avg);
                if (statMinMaxEl) statMinMaxEl.textContent = `${formatNumberDisplay(min)} / ${formatNumberDisplay(max)}`;
            } else {
                if (statSumEl) statSumEl.textContent = "-";
                if (statAvgEl) statAvgEl.textContent = "-";
                if (statMinMaxEl) statMinMaxEl.textContent = "-";
            }

            if (statCountEl) {
                statCountEl.textContent = `${countTotal} sel (${uniqueSet.size} unik)`;
            }

            if (!silent) {
                showStatus(`Statistik dihitung untuk ${countTotal} sel terpilih.`, "info");
            }
        });
    } catch (error) {
        if (!silent) {
            console.error("Gagal menghitung statistik:", error);
        }
    }
}

/**
 * Menyisipkan baris TOTAL akuntansi otomatis di bawah tabel data yang diblok.
 */
async function handleInsertTotalRow() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["rowCount", "columnCount", "address", "values"]);
            await context.sync();

            if (range.rowCount < 1 || range.columnCount < 1) {
                showStatus("Pilih rentang tabel terlebih dahulu!", "warning");
                return;
            }

            // Dapatkan baris tepat di bawah rentang yang dipilih
            const totalRow = range.getOffsetRange(range.rowCount, 0).getResizedRange(-(range.rowCount - 1), 0);
            totalRow.load(["address"]);
            await context.sync();

            const rowCount = range.rowCount;
            const colCount = range.columnCount;
            const formulas = [[]];

            // Tulis 'TOTAL' di kolom pertama dan SUM untuk kolom angka lainnya
            for (let c = 0; c < colCount; c++) {
                if (c === 0) {
                    formulas[0].push("TOTAL");
                } else {
                    // Dapatkan referensi kolom sel atas sampai bawah
                    const colRange = range.getColumn(c);
                    colRange.load(["address"]);
                    await context.sync();
                    formulas[0].push(`=SUM(${colRange.address})`);
                }
            }

            totalRow.formulas = formulas;

            // Format Baris Total Gaya Akuntansi Standar (Double Bottom Border, Bold)
            totalRow.format.font.bold = true;
            totalRow.format.fill.color = "#f4fbf7";
            totalRow.format.borders.getItem("EdgeTop").style = "Continuous";
            totalRow.format.borders.getItem("EdgeTop").weight = "Thin";
            totalRow.format.borders.getItem("EdgeTop").color = "#242424";
            totalRow.format.borders.getItem("EdgeBottom").style = "Double";
            totalRow.format.borders.getItem("EdgeBottom").weight = "Thick";
            totalRow.format.borders.getItem("EdgeBottom").color = "#107c41";

            await context.sync();
            showStatus("Baris TOTAL akuntansi berhasil disisipkan!", "success");
        });
    } catch (error) {
        console.error("Gagal menyisipkan baris total:", error);
        showStatus(`Gagal menyisipkan total: ${error.message}`, "danger");
    }
}

/**
 * Format ringkas angka statistik (memotong desimal panjang).
 */
function formatNumberDisplay(num) {
    if (Math.abs(num) >= 1000) {
        return num.toLocaleString("id-ID", { maximumFractionDigits: 2 });
    }
    return Number(num.toFixed(2)).toString();
}

/* ==========================================================================
   5. FITUR 3: SMART DATA CLEANER (1-CLICK CLEANER)
   ========================================================================== */

/**
 * Membersihkan data teks pada rentang sel yang dipilih.
 * Mendukung mode: 'trim', 'proper', 'upper', 'lower'.
 */
async function handleCleanData(mode) {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const values = range.values;
            let modifiedCount = 0;

            const updatedValues = values.map(row => {
                return row.map(cellValue => {
                    if (typeof cellValue !== "string" || !cellValue) {
                        return cellValue;
                    }

                    modifiedCount++;
                    switch (mode) {
                        case "trim":
                            return cellValue.replace(/\s+/g, " ").trim();
                        case "proper":
                            return toProperCase(cellValue);
                        case "upper":
                            return cellValue.toUpperCase();
                        case "lower":
                            return cellValue.toLowerCase();
                        default:
                            return cellValue;
                    }
                });
            });

            range.values = updatedValues;
            await context.sync();

            const labelMap = {
                trim: "Trim Spasi",
                proper: "Proper Case",
                upper: "UPPERCASE",
                lower: "lowercase"
            };

            showStatus(`Pembersihan [${labelMap[mode] || mode}] berhasil diterapkan pada ${modifiedCount} sel.`, "success");
        });
    } catch (error) {
        console.error("Gagal membersihkan data:", error);
        showStatus(`Gagal membersihkan data: ${error.message}`, "danger");
    }
}

/**
 * Mengubah format rentang sel aktif menjadi format Rupiah Indonesia (Rp #,##0).
 */
async function handleFormatRupiah() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.numberFormat = [["_(\"Rp\"* #,##0_);_(\"Rp\"* (#,##0);_(\"Rp\"* \"-\"_);_(@_)"]];
            await context.sync();
            showStatus("Format Mata Uang Rupiah berhasil diterapkan!", "success");
        });
    } catch (error) {
        console.error("Gagal format Rupiah:", error);
        showStatus(`Gagal format Rupiah: ${error.message}`, "danger");
    }
}

/**
 * Mengubah format rentang sel aktif menjadi format Tanggal Standar Indonesia (DD/MM/YYYY).
 */
async function handleFormatDate() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.numberFormat = [["DD/MM/YYYY"]];
            await context.sync();
            showStatus("Format Tanggal (DD/MM/YYYY) berhasil diterapkan!", "success");
        });
    } catch (error) {
        console.error("Gagal format Tanggal:", error);
        showStatus(`Gagal format Tanggal: ${error.message}`, "danger");
    }
}

/**
 * Helper: Mengubah string menjadi Proper Case (Kapital di setiap awal kata).
 */
function toProperCase(str) {
    return str
        .toLowerCase()
        .replace(/\s+/g, " ")
        .trim()
        .replace(/(?:^|\s|\/|-)\S/g, (char) => char.toUpperCase());
}

/* ==========================================================================
   6. FITUR 4: SMART DEDUPLICATION & TEXT TOOLS
   ========================================================================== */

/**
 * Menyorot data ganda (duplikat) di dalam rentang sel yang diblok dengan warna amber lembut.
 */
async function handleHighlightDuplicates() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const values = range.values;
            const frequencyMap = new Map();

            // Hitung frekuensi setiap nilai
            for (let r = 0; r < values.length; r++) {
                for (let c = 0; c < values[r].length; c++) {
                    const val = String(values[r][c] ?? "").trim();
                    if (val !== "") {
                        frequencyMap.set(val, (frequencyMap.get(val) || 0) + 1);
                    }
                }
            }

            let dupCount = 0;
            // Sorot sel yang muncul lebih dari 1 kali
            for (let r = 0; r < values.length; r++) {
                for (let c = 0; c < values[r].length; c++) {
                    const val = String(values[r][c] ?? "").trim();
                    if (val !== "" && frequencyMap.get(val) > 1) {
                        const cell = range.getCell(r, c);
                        cell.format.fill.color = "#fff3cd"; // Kuning/Amber lembut
                        cell.format.font.color = "#856404";
                        dupCount++;
                    }
                }
            }

            await context.sync();
            if (dupCount > 0) {
                showStatus(`Ditemukan & disorot ${dupCount} sel bernilai duplikat.`, "warning");
            } else {
                showStatus("Tidak ada data duplikat yang ditemukan.", "success");
            }
        });
    } catch (error) {
        console.error("Gagal menyorot duplikat:", error);
        showStatus(`Gagal menyorot duplikat: ${error.message}`, "danger");
    }
}

/**
 * Menghapus baris duplikat dari rentang data yang diblok dan mempertahankan baris unik pertama.
 */
async function handleRemoveDuplicates() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const values = range.values;
            if (values.length <= 1) {
                showStatus("Pilih minimal 2 baris data untuk menghapus duplikat.", "warning");
                return;
            }

            const seen = new Set();
            const uniqueRows = [];
            let removedCount = 0;

            for (let r = 0; r < values.length; r++) {
                const rowKey = JSON.stringify(values[r]);
                if (!seen.has(rowKey)) {
                    seen.add(rowKey);
                    uniqueRows.push(values[r]);
                } else {
                    removedCount++;
                }
            }

            if (removedCount === 0) {
                showStatus("Seluruh baris data sudah unik (tidak ada duplikat).", "info");
                return;
            }

            // Bersihkan rentang lama lalu tulis kembali baris unik
            range.clear();
            const writeRange = range.getResizedRange(-(range.rowCount - uniqueRows.length), 0);
            writeRange.values = uniqueRows;

            await context.sync();
            showStatus(`Berhasil menghapus ${removedCount} baris duplikat!`, "success");
        });
    } catch (error) {
        console.error("Gagal menghapus duplikat:", error);
        showStatus(`Gagal menghapus duplikat: ${error.message}`, "danger");
    }
}

/**
 * Mengekstrak daftar nilai unik dari kolom yang diblok ke kolom baru di sampingnya.
 */
async function handleExtractUnique() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const values = range.values;
            const uniqueSet = new Set();

            for (let r = 0; r < values.length; r++) {
                const val = String(values[r][0] ?? "").trim();
                if (val !== "") {
                    uniqueSet.add(val);
                }
            }

            const uniqueArray = Array.from(uniqueSet);
            if (uniqueArray.length === 0) {
                showStatus("Tidak ada data untuk diekstrak.", "warning");
                return;
            }

            const outValues = [["NILAI UNIK"], ...uniqueArray.map(v => [v])];
            const targetRange = range.getOffsetRange(0, range.columnCount).getResizedRange(outValues.length - 1, -(range.columnCount - 1));

            targetRange.values = outValues;
            targetRange.getRow(0).format.font.bold = true;
            targetRange.getRow(0).format.fill.color = "#e8f5e9";
            targetRange.format.autofitColumns();

            await context.sync();
            showStatus(`Berhasil mengekstrak ${uniqueArray.length} nilai unik ke kolom samping!`, "success");
        });
    } catch (error) {
        console.error("Gagal ekstrak unik:", error);
        showStatus(`Gagal ekstrak unik: ${error.message}`, "danger");
    }
}

/**
 * Memecah isi satu kolom teks menjadi beberapa kolom berdasarkan delimiter.
 */
async function handleSplitTextColumn() {
    const delimiter = document.getElementById("text-delimiter")?.value || " ";

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const values = range.values;
            let maxCols = 1;

            const splitData = values.map(row => {
                const text = String(row[0] ?? "");
                const parts = text.split(delimiter).map(p => p.trim());
                if (parts.length > maxCols) maxCols = parts.length;
                return parts;
            });

            if (maxCols <= 1) {
                showStatus("Tidak ada teks yang dapat dipecah dengan pemisah ini.", "warning");
                return;
            }

            // Normalisasi panjang kolom per baris
            const finalData = splitData.map(row => {
                while (row.length < maxCols) row.push("");
                return row;
            });

            const targetRange = range.getOffsetRange(0, 1).getResizedRange(0, maxCols - 2);
            targetRange.values = finalData.map(row => row.slice(1));
            range.values = finalData.map(row => [row[0]]);

            await context.sync();
            showStatus(`Teks berhasil dipecah menjadi ${maxCols} kolom!`, "success");
        });
    } catch (error) {
        console.error("Gagal memecah teks:", error);
        showStatus(`Gagal memecah teks: ${error.message}`, "danger");
    }
}

/**
 * Menggabungkan beberapa kolom teks yang diblok menjadi 1 kolom dengan delimiter.
 */
async function handleMergeTextColumns() {
    const delimiter = document.getElementById("text-delimiter")?.value || " ";

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            if (range.columnCount <= 1) {
                showStatus("Blok minimal 2 kolom untuk digabungkan!", "warning");
                return;
            }

            const values = range.values;
            const mergedValues = [["GABUNGAN"], ...values.slice(1).map(row => {
                return [row.map(cell => String(cell ?? "").trim()).filter(Boolean).join(delimiter)];
            })];

            // Tulis hasil gabungan di kolom setelah rentang yang dipilih
            const targetRange = range.getOffsetRange(0, range.columnCount).getResizedRange(mergedValues.length - 1, -(range.columnCount - 1));
            targetRange.values = mergedValues;
            targetRange.getRow(0).format.font.bold = true;
            targetRange.getRow(0).format.fill.color = "#e8f5e9";
            targetRange.format.autofitColumns();

            await context.sync();
            showStatus("Kolom berhasil digabungkan!", "success");
        });
    } catch (error) {
        console.error("Gagal menggabung kolom:", error);
        showStatus(`Gagal menggabung kolom: ${error.message}`, "danger");
    }
}

/* ==========================================================================
   7. FITUR 5: 1-CLICK AUTO CHART GENERATOR
   ========================================================================== */

/**
 * Membuat grafik (Chart) otomatis dari tabel/rentang yang diblok.
 * 
 * @param {Excel.ChartType} chartType - Tipe diagram Excel
 */
async function handleCreateChart(chartType) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = context.workbook.getSelectedRange();
            range.load(["rowCount", "columnCount", "address"]);
            await context.sync();

            if (range.rowCount < 2 || range.columnCount < 1) {
                showStatus("Blok tabel data beserta label headernya untuk membuat chart!", "warning");
                return;
            }

            // Tambahkan diagram baru
            const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
            chart.title.text = "Grafik Visualisasi Data";

            // Posisikan diagram di sebelah kanan tabel
            const offsetCell = range.getOffsetRange(0, range.columnCount + 1).getCell(0, 0);
            chart.setPosition(offsetCell, null);
            chart.width = 420;
            chart.height = 260;

            await context.sync();
            showStatus("Grafik otomatis berhasil dibuat!", "success");
        });
    } catch (error) {
        console.error("Gagal membuat chart:", error);
        showStatus(`Gagal membuat chart: ${error.message}`, "danger");
    }
}

/* ==========================================================================
   8. FITUR 6: DATA QUALITY AUDITOR
   ========================================================================== */

/**
 * Menyorot sel kosong (bolong) pada tabel yang diblok dengan warna salmon/merah muda.
 */
async function handleHighlightBlankCells() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const values = range.values;
            let blankCount = 0;

            for (let r = 0; r < values.length; r++) {
                for (let c = 0; c < values[r].length; c++) {
                    const val = values[r][c];
                    if (val === "" || val === null || val === undefined) {
                        const cell = range.getCell(r, c);
                        cell.format.fill.color = "#f8d7da"; // Merah muda lembut
                        cell.format.font.color = "#721c24";
                        blankCount++;
                    }
                }
            }

            await context.sync();
            if (blankCount > 0) {
                showStatus(`Ditemukan ${blankCount} sel kosong dan telah disorot.`, "warning");
            } else {
                showStatus("Semua sel terisi lengkap (tidak ada sel kosong).", "success");
            }
        });
    } catch (error) {
        console.error("Gagal menyorot sel kosong:", error);
        showStatus(`Gagal menyorot sel kosong: ${error.message}`, "danger");
    }
}

/**
 * Mengonversi angka yang tersimpan sebagai teks kembali menjadi format angka numerik murni.
 */
async function handleConvertTextToNumbers() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const values = range.values;
            let convertedCount = 0;

            const updatedValues = values.map(row => {
                return row.map(cell => {
                    if (typeof cell === "string") {
                        // Bersihkan simbol rupiah atau spasi sebelum parsing
                        const cleanStr = cell.replace(/^Rp\s?/i, "").replace(/\./g, "").replace(/,/g, ".").trim();
                        const num = Number(cleanStr);
                        if (!isNaN(num) && cleanStr !== "") {
                            convertedCount++;
                            return num;
                        }
                    }
                    return cell;
                });
            });

            range.values = updatedValues;
            range.numberFormat = [["#,##0"]];
            await context.sync();

            showStatus(`Berhasil mengonversi ${convertedCount} teks menjadi angka numerik!`, "success");
        });
    } catch (error) {
        console.error("Gagal konversi teks ke angka:", error);
        showStatus(`Gagal konversi teks: ${error.message}`, "danger");
    }
}

/**
 * Membersihkan semua warna latar belakang (highlight) dan mereset format sel yang diblok.
 */
async function handleClearFormatting() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.clear();
            range.format.font.color = "#242424";
            await context.sync();
            showStatus("Warna highlight & format berhasil dibersihkan!", "info");
        });
    } catch (error) {
        console.error("Gagal membersihkan format:", error);
        showStatus(`Gagal membersihkan format: ${error.message}`, "danger");
    }
}

/* ==========================================================================
   9. FITUR 7: TEMPLATE GENERATOR
   ========================================================================== */

/**
 * Membuat tabel template standar dengan header hijau Fluent, autofit, dan validasi data.
 */
async function handleGenerateTemplate() {
    const select = document.getElementById("template-select");
    const templateKey = select ? select.value : "kinerja";

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const config = getTemplateConfig(templateKey);

            const range = sheet.getRangeByIndexes(0, 0, config.rows.length, config.headers.length);
            range.values = [config.headers, ...config.rows];

            // 1. Format Header Fluent Green
            const headerRange = range.getRow(0);
            headerRange.format.fill.color = "#107c41";
            headerRange.format.font.color = "#ffffff";
            headerRange.format.font.bold = true;
            headerRange.format.font.name = "Segoe UI";

            // 2. Format Body Data
            const bodyRange = range.getRowsBelow(1);
            bodyRange.format.font.name = "Segoe UI";

            // 3. Freeze Top Row
            sheet.freezePanes.freezeRows(1);

            // 4. Data Validation untuk kolom status jika tersedia
            if (config.validation) {
                const validationRange = sheet.getRange(config.validation.rangeAddress);
                validationRange.dataValidation.rule = {
                    list: {
                        inCellDropDown: true,
                        source: config.validation.options.join(",")
                    }
                };
            }

            // 5. Auto-fit Columns
            range.format.autofitColumns();
            await context.sync();

            showStatus(`Template [${config.name}] berhasil dibuat!`, "success");
        });
    } catch (error) {
        console.error("Gagal membuat template:", error);
        showStatus(`Gagal membuat template: ${error.message}`, "danger");
    }
}

/**
 * Konfigurasi data template bawaan.
 */
function getTemplateConfig(key) {
    const templates = {
        kinerja: {
            name: "Evaluasi Kinerja Karyawan",
            headers: ["ID Karyawan", "Nama Karyawan", "Departemen", "Target (%)", "Realisasi (%)", "Skor Kinerja", "Status Review"],
            rows: [
                ["EMP-001", "Budi Santoso", "Pemasaran", 100, 105, 92, "Sangat Baik"],
                ["EMP-002", "Siti Rahma", "Keuangan", 100, 98, 88, "Baik"],
                ["EMP-003", "Andi Wijaya", "Teknologi", 100, 85, 76, "Cukup"]
            ],
            validation: {
                rangeAddress: "G2:G100",
                options: ["Sangat Baik", "Baik", "Cukup", "Perlu Ditingkatkan"]
            }
        },
        arus_kas: {
            name: "Laporan Arus Kas",
            headers: ["Tanggal", "Kategori Transaksi", "Keterangan", "Pemasukan (Rp)", "Pengeluaran (Rp)", "Saldo (Rp)"],
            rows: [
                ["01/01/2026", "Modal Awal", "Penyetoran modal", 50000000, 0, 50000000],
                ["05/01/2026", "Operasional", "Biaya hosting & domain", 0, 1500000, 48500000],
                ["10/01/2026", "Penjualan", "Pendapatan jasa klien", 12000000, 0, 60500000]
            ],
            validation: {
                rangeAddress: "B2:B100",
                options: ["Penjualan", "Operasional", "Modal Awal", "Pajak", "Gaji"]
            }
        },
        jadwal_proyek: {
            name: "Jadwal Proyek / Project Tracker",
            headers: ["No", "Nama Tugas", "PIC / Penanggung Jawab", "Tgl Mulai", "Tgl Selesai", "Status", "Progress (%)"],
            rows: [
                [1, "Perencanaan UI/UX", "Ahmad", "01/08/2026", "07/08/2026", "Selesai", 100],
                [2, "Pengembangan Backend", "Rian", "08/08/2026", "20/08/2026", "Sedang Berjalan", 65],
                [3, "Pengujian Mutu & QC", "Dewi", "21/08/2026", "25/08/2026", "Belum Dimulai", 0]
            ],
            validation: {
                rangeAddress: "F2:F100",
                options: ["Belum Dimulai", "Sedang Berjalan", "Selesai", "Tertunda"]
            }
        },
        invoice: {
            name: "Daftar Tagihan / Invoice",
            headers: ["No Invoice", "Nama Klien", "Tgl Invoice", "Jatuh Tempo", "Total Tagihan (Rp)", "Status Pembayaran"],
            rows: [
                ["INV-2026-001", "PT Maju Bersama", "10/08/2026", "24/08/2026", 15000000, "Lunas"],
                ["INV-2026-002", "CV Berkah Abadi", "15/08/2026", "29/08/2026", 8500000, "Belum Lunas"],
                ["INV-2026-003", "Klinik Sehat", "20/08/2026", "03/09/2026", 22000000, "Belum Lunas"]
            ],
            validation: {
                rangeAddress: "F2:F100",
                options: ["Lunas", "Belum Lunas", "Jatuh Tempo"]
            }
        },
        media: {
            name: "Media Partner Tracker",
            headers: ["Nama Media", "Kontak / PIC", "Email", "Paket Promosi", "Status Kerjasama"],
            rows: [
                ["Tech Daily", "Rudi", "rudi@techdaily.id", "Sponsored Post", "Deal"],
                ["Campus Info", "Dewi", "dewi@campus.id", "Instagram Story", "Follow Up"],
                ["Berita Event", "Hendra", "redaksi@event.com", "Liputan Khusus", "Terkirim"]
            ],
            validation: {
                rangeAddress: "E2:E100",
                options: ["Terkirim", "Follow Up", "Deal", "Ditolak"]
            }
        },
        literatur: {
            name: "Tinjauan Pustaka",
            headers: ["Tahun", "Penulis", "Judul Artikel / Buku", "Metode Penelitian", "Temuan Utama", "Relevansi"],
            rows: [
                [2024, "Smith, J. et al.", "Machine Learning in Financial Forecasting", "Kuantitatif - LSTM", "Akurasi prediksi naik 14%", "Tinggi"],
                [2025, "Pratama, A.", "Penerapan Office.js pada Otomatisasi Administrasi", "Eksperimen", "Efisiensi waktu kerja meningkat 60%", "Sangat Tinggi"]
            ],
            validation: {
                rangeAddress: "F2:F100",
                options: ["Sangat Tinggi", "Tinggi", "Sedang", "Rendah"]
            }
        }
    };

    return templates[key] || templates.kinerja;
}

/* ==========================================================================
   10. FITUR 8: MASS ACTION (SPLIT DATA TO SHEETS)
   ========================================================================== */

/**
 * Memecah tabel data besar menjadi Worksheet terpisah berdasarkan kategori unik pada kolom yang diblok.
 */
async function handleSplitDataToSheets() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount"]);

            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            activeSheet.load(["name"]);
            await context.sync();

            const values = range.values;
            if (values.length <= 1) {
                showStatus("Pilih minimal 1 baris header dan data!", "warning");
                return;
            }

            const headerRow = values[0];
            const dataRows = values.slice(1);

            const groupedData = new Map();
            dataRows.forEach(row => {
                const category = String(row[0] ?? "Tanpa_Kategori").trim() || "Tanpa_Kategori";
                if (!groupedData.has(category)) {
                    groupedData.set(category, []);
                }
                groupedData.get(category).push(row);
            });

            if (groupedData.size === 0) {
                showStatus("Tidak ada data kategori yang valid.", "warning");
                return;
            }

            const sheets = context.workbook.worksheets;
            sheets.load(["items"]);
            await context.sync();

            const existingSheetNames = new Set(sheets.items.map(s => s.name));
            let createdCount = 0;

            for (const [categoryName, rows] of groupedData.entries()) {
                const sanitizedName = sanitizeSheetName(categoryName);
                let targetSheet;

                if (existingSheetNames.has(sanitizedName)) {
                    targetSheet = sheets.getItem(sanitizedName);
                } else {
                    targetSheet = sheets.add(sanitizedName);
                    existingSheetNames.add(sanitizedName);
                    createdCount++;
                }

                const newSheetRange = targetSheet.getRangeByIndexes(0, 0, rows.length + 1, headerRow.length);
                newSheetRange.values = [headerRow, ...rows];

                // Header styling
                const newHeader = newSheetRange.getRow(0);
                newHeader.format.fill.color = "#107c41";
                newHeader.format.font.color = "#ffffff";
                newHeader.format.font.bold = true;
                newSheetRange.format.autofitColumns();
            }

            await context.sync();
            showStatus(`Berhasil memecah data ke ${groupedData.size} sheet (${createdCount} sheet baru dibuat)!`, "success");
        });
    } catch (error) {
        console.error("Gagal memecah sheet:", error);
        showStatus(`Gagal memecah data ke sheet: ${error.message}`, "danger");
    }
}

/**
 * Helper: Membersihkan nama sheet agar valid sesuai aturan Microsoft Excel.
 */
function sanitizeSheetName(name) {
    const cleaned = name.replace(/[\\/?*:[\]]/g, "_").trim();
    return cleaned.substring(0, 31) || "Sheet_Kategori";
}

/* ==========================================================================
   11. HELPER STATUS NOTIFIKASI
   ========================================================================== */

/**
 * Menampilkan pesan status notifikasi pada Task Pane.
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
