/* ============================================================
   FACILITY + MODALITY DEFINITIONS — EXACT FORMAT
   ============================================================ */

const FACILITIES = {
    "Diagnostic Medical Group Arcadia": "DMG Arcadia",
    "Diagnostic Medical Group City of Industry": "DMG City of Industry",
    "Diagnostic Medical Group Monterey Park": "DMG Monterey Park",
    "Diagnostic Medical Group San Gabriel": "DMG San Gabriel",
    "Synergy San Gabriel": "SYN San Gabriel"
};

const FACILITY_MODALITIES = {
    "DMG Arcadia": ["CT", "Bone Density", "MG", "US", "XR"],
    "DMG City of Industry": ["CT", "Bone Density", "MG", "MR", "US", "XR"],
    "DMG Monterey Park": ["CT", "Bone Density", "MG", "MR", "US", "XR"],
    "DMG San Gabriel": ["CT", "Bone Density", "EKG", "MG", "MR", "US", "XR"],
    "SYN San Gabriel": ["CT", "EKG", "MR", "US", "XR"]
};

/* ============================================================
   STATUS
   ============================================================ */

function showStatus(id, msg) {
    ["processing", "errorMessage", "successMessage"].forEach(el => {
        document.getElementById(el).style.display = "none";
    });
    const target = document.getElementById(id);
    target.innerText = msg;
    target.style.display = "block";
}

/* ============================================================
   MAIN ENTRY
   ============================================================ */

function runDailySummary() {
    showStatus("processing", "Processing…");

    const fileInput = document.getElementById("dailyFile");
    const file = fileInput.files[0];
    if (!file) return showStatus("errorMessage", "No file selected.");

    const reader = new FileReader();
    const fileName = file.name.toLowerCase();

    reader.onload = e => {
        if (fileName.endsWith(".csv")) parseCSV(e.target.result);
        else parseExcel(e.target.result);
    };

    if (fileName.endsWith(".csv")) reader.readAsBinaryString(file);
    else reader.readAsArrayBuffer(file);
}

/* ============================================================
   DATE FIX
   ============================================================ */

function fixDate(value) {
    if (!value) return "";
    const num = Number(value);
    if (!isNaN(num) && num > 20000 && num < 60000) {
        const base = new Date(1899, 11, 30);
        base.setDate(base.getDate() + num);
        return base.toLocaleDateString("en-US");
    }
    const d = new Date(String(value).trim());
    return isNaN(d) ? "" : d.toLocaleDateString("en-US");
}

/* ============================================================
   PARSE CSV
   ============================================================ */

function parseCSV(raw) {
    const uint8 = new Uint8Array(raw.length);
    for (let i = 0; i < raw.length; i++) uint8[i] = raw.charCodeAt(i) & 0xFF;

    const workbook = XLSX.read(uint8, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    generateDailyTables(aoa);
}

/* ============================================================
   PARSE XLSX
   ============================================================ */

function parseExcel(buffer) {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    generateDailyTables(aoa);
}

/* ============================================================
   BUILD DAILY DATA
   ============================================================ */

function generateDailyTables(aoa) {
    const HEADER_ROW = 8;
    const DATA_START = HEADER_ROW + 1;

    if (!aoa[HEADER_ROW]) return showStatus("errorMessage", "Header row 8 not found.");

    const daily = {};

    for (let r = DATA_START; r < aoa.length; r++) {
        const row = aoa[r];
        if (!row || row.length < 7) continue;

        const modalityRaw = String(row[0] || "").trim();
        const locationFull = String(row[1] || "").trim();
        const dosRaw = row[5];

        const dos = fixDate(dosRaw);
        const facility = FACILITIES[locationFull];

        if (!dos || !facility || !modalityRaw) continue;

        const modalityMap = {
            "DEXA": "Bone Density",
            "EKG": "EKG"
        };
        const modality = modalityMap[modalityRaw] || modalityRaw;

        if (!daily[dos]) daily[dos] = {};
        if (!daily[dos][facility]) daily[dos][facility] = {};

        daily[dos][facility][modality] =
            (daily[dos][facility][modality] || 0) + 1;
    }

    displayDailyTables(daily);
    showStatus("successMessage", "Completed! Tables generated below.");
}

/* ============================================================
   GROUP DATES BY MONTH
   ============================================================ */

function getMonthKey(dateStr) {
    const d = new Date(dateStr);
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    return `${y}-${m.toString().padStart(2, "0")}`;
}

/* ============================================================
   DISPLAY — SEPARATE TABLES PER MONTH + MONTHLY TOTAL ROW + FINAL TOTAL
   ============================================================ */

function displayDailyTables(daily) {
    const left = document.getElementById("leftColumn");
    left.innerHTML = "";

    const allDates = Object.keys(daily).sort((a, b) => new Date(a) - new Date(b));

    const months = {};
    allDates.forEach(d => {
        const key = getMonthKey(d);
        if (!months[key]) months[key] = [];
        months[key].push(d);
    });

    const finalTotals = {};
    Object.values(FACILITIES).forEach(fac => {
        finalTotals[fac] = {};
        FACILITY_MODALITIES[fac].forEach(mod => finalTotals[fac][mod] = 0);
        finalTotals[fac].Total = 0;
    });

    let finalGrandTotal = 0;

    Object.keys(months).sort().forEach(monthKey => {
        const dates = months[monthKey];

        const table = document.createElement("table");

        let header = "<tr><th>Date</th>";
        Object.values(FACILITIES).forEach(fac => {
            const mods = FACILITY_MODALITIES[fac];
            header += `<th colspan="${mods.length + 1}">${fac}</th>`;
        });
        header += `<th>Grand Total</th></tr><tr><th></th>`;

        Object.values(FACILITIES).forEach(fac => {
            const mods = FACILITY_MODALITIES[fac];
            mods.forEach(mod => header += `<th>${mod}</th>`);
            header += `<th>Total</th>`;
        });
        header += `<th></th></tr>`;
        table.innerHTML = header;

        let monthlyTotals = {};
        Object.values(FACILITIES).forEach(fac => {
            monthlyTotals[fac] = {};
            FACILITY_MODALITIES[fac].forEach(mod => monthlyTotals[fac][mod] = 0);
            monthlyTotals[fac].Total = 0;
        });

        let monthlyGrandTotal = 0;

        dates.forEach(dos => {
            let row = `<tr><td>${dos}</td>`;
            let grandTotal = 0;

            Object.values(FACILITIES).forEach(fac => {
                const mods = FACILITY_MODALITIES[fac];
                let facTotal = 0;

                mods.forEach(mod => {
                    const val = daily[dos][fac]?.[mod] || "";
                    row += `<td>${val}</td>`;
                    facTotal += Number(val || 0);

                    monthlyTotals[fac][mod] += Number(val || 0);
                    finalTotals[fac][mod] += Number(val || 0);
                });

                row += `<td>${facTotal}</td>`;
                grandTotal += facTotal;

                monthlyTotals[fac].Total += facTotal;
                finalTotals[fac].Total += facTotal;
            });

            row += `<td>${grandTotal}</td></tr>`;
            table.innerHTML += row;

            monthlyGrandTotal += grandTotal;
            finalGrandTotal += grandTotal;
        });

        let totalRow = `<tr><td><b>Monthly Total</b></td>`;

        Object.values(FACILITIES).forEach(fac => {
            FACILITY_MODALITIES[fac].forEach(mod => {
                totalRow += `<td><b>${monthlyTotals[fac][mod]}</b></td>`;
            });
            totalRow += `<td><b>${monthlyTotals[fac].Total}</b></td>`;
        });

        totalRow += `<td><b>${monthlyGrandTotal}</b></td></tr>`;
        table.innerHTML += totalRow;

        left.appendChild(table);
    });

    const finalTable = document.createElement("table");

    let finalHeader = "<tr><th>ALL MONTHS TOTAL</th>";
    Object.values(FACILITIES).forEach(fac => {
        const mods = FACILITY_MODALITIES[fac];
        finalHeader += `<th colspan="${mods.length + 1}">${fac}</th>`;
    });
    finalHeader += `<th>Grand Total</th></tr><tr><th></th>`;

    Object.values(FACILITIES).forEach(fac => {
        FACILITY_MODALITIES[fac].forEach(mod => finalHeader += `<th>${mod}</th>`);
        finalHeader += `<th>Total</th>`;
    });
    finalHeader += `<th></th></tr>`;

    finalTable.innerHTML = finalHeader;

    let totalRow = `<tr><td><b>Final Total</b></td>`;

    Object.values(FACILITIES).forEach(fac => {
        FACILITY_MODALITIES[fac].forEach(mod => {
            totalRow += `<td><b>${finalTotals[fac][mod]}</b></td>`;
        });
        totalRow += `<td><b>${finalTotals[fac].Total}</b></td>`;
    });

    totalRow += `<td><b>${finalGrandTotal}</b></td></tr>`;
    finalTable.innerHTML += totalRow;

    left.appendChild(finalTable);
}

/* ============================================================
   DOWNLOAD
   ============================================================ */

function downloadOutput() {
    alert("Daily tables are visual only.");
}
