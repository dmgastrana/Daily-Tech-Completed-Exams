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

// EXACT modality sets per facility
const FACILITY_MODALITIES = {
    "DMG Arcadia": ["CT", "Bone Density", "MG", "US", "XR"],
    "DMG City of Industry": ["CT", "Bone Density", "MG", "MR", "US", "XR"],
    "DMG Monterey Park": ["CT", "Bone Density", "MG", "MR", "US", "XR"],
    "DMG San Gabriel": ["CT", "Bone Density", "EKG", "MG", "MR", "US", "XR"],
    "SYN San Gabriel": ["CT", "EKG", "MR", "US", "XR"]
};

/* ============================================================
   STATUS HANDLING
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
   FIX DATE (Excel serials + text)
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
   PARSE CSV USING XLSX ENGINE
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
   PARSE EXCEL
   ============================================================ */

function parseExcel(buffer) {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    generateDailyTables(aoa);
}

/* ============================================================
   MAIN DAILY TABLE GENERATION — EXACT FORMAT
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
   DISPLAY DAILY TABLES — EXACT FORMAT + MONTHLY TOTAL ROW
   ============================================================ */

function displayDailyTables(daily) {
    const left = document.getElementById("leftColumn");
    left.innerHTML = "";

    const dates = Object.keys(daily).sort((a, b) => new Date(a) - new Date(b));

    const table = document.createElement("table");

    /* GROUPED FACILITY HEADERS — EXACT FORMAT */
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

    /* DAILY ROWS */
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
            });

            row += `<td>${facTotal}</td>`;
            grandTotal += facTotal;
        });

        row += `<td>${grandTotal}</td></tr>`;
        table.innerHTML += row;
    });

    /* MONTHLY TOTAL ROW — EXACT FORMAT */
    let totalRow = `<tr><td><b>Monthly Total</b></td>`;
    let monthlyGrand = 0;

    Object.values(FACILITIES).forEach(fac => {
        const mods = FACILITY_MODALITIES[fac];
        let facMonthly = 0;

        mods.forEach(mod => {
            let sum = 0;
            dates.forEach(d => {
                sum += Number(daily[d][fac]?.[mod] || 0);
            });
            totalRow += `<td><b>${sum}</b></td>`;
            facMonthly += sum;
        });

        totalRow += `<td><b>${facMonthly}</b></td>`;
        monthlyGrand += facMonthly;
    });

    totalRow += `<td><b>${monthlyGrand}</b></td></tr>`;
    table.innerHTML += totalRow;

    left.appendChild(table);
}

/* ============================================================
   DOWNLOAD OUTPUT
   ============================================================ */

function downloadOutput() {
    alert("Daily tables are visual only.");
}


