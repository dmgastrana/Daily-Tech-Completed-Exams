/* ============================================================
   CONFIGURATION
   ============================================================ */

const LOCATION_MAP = {
    "Diagnostic Medical Group Arcadia": "AR",
    "Diagnostic Medical Group City of Industry": "CI",
    "Diagnostic Medical Group Monterey Park": "MP",
    "Diagnostic Medical Group San Gabriel": "SG",
    "Synergy San Gabriel": "SSG"
};

const TABLE_STRUCTURE = {
    AR: ["CT", "Bone Density", "MG", "US", "XR"],
    CI: ["CT", "Bone Density", "MG", "MR", "US", "XR"],
    MP: ["CT", "Bone Density", "MG", "MR", "US", "XR"],
    SG: ["CT", "EKG", "MG", "MR", "US", "XR"],
    SSG: ["CT", "EKG", "MG", "MR", "US", "XR"]
};

/* ============================================================
   STATUS MESSAGE HANDLING
   ============================================================ */

function showStatus(id, msg) {
    document.getElementById("processing").style.display = "none";
    document.getElementById("errorMessage").style.display = "none";
    document.getElementById("successMessage").style.display = "none";

    const el = document.getElementById(id);
    el.innerText = msg;
    el.style.display = "block";
}

/* ============================================================
   MAIN ENTRY
   ============================================================ */

function runDailySummary() {
    showStatus("processing", "Starting…");

    const fileInput = document.getElementById("dailyFile");
    const file = fileInput.files[0];

    if (!file) {
        showStatus("errorMessage", "No file selected.");
        return;
    }

    const fileName = file.name.toLowerCase();
    showStatus("processing", "File detected: " + fileName);

    const reader = new FileReader();

    reader.onload = function(e) {
        let text = e.target.result;

        // CSV fallback detection
        if (fileName.endsWith(".csv")) {
            if (!text || text.trim().length < 10) {
                // Retry using ArrayBuffer
                const reader2 = new FileReader();
                reader2.onload = function(ev) {
                    const buffer = ev.target.result;
                    const decoder = new TextDecoder("utf-8");
                    const decoded = decoder.decode(buffer);
                    parseCSV(decoded);
                };
                reader2.readAsArrayBuffer(file);
                return;
            }

            parseCSV(text);
            return;
        }

        // Excel
        parseExcel(e.target.result);
    };

    if (fileName.endsWith(".csv")) {
        reader.readAsText(file); // may fail silently → fallback handles it
    } else {
        reader.readAsArrayBuffer(file);
    }
}

/* ============================================================
   PARSE CSV
   ============================================================ */

function parseCSV(text) {
    showStatus("processing", "CSV loaded — splitting rows…");

    const rows = text.split(/\r?\n/).map(r => r.split(","));
    generateMonthlyTables(rows);
}

/* ============================================================
   PARSE EXCEL
   ============================================================ */

function parseExcel(buffer) {
    showStatus("processing", "Excel loaded — converting…");

    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    generateMonthlyTables(aoa);
}

/* ============================================================
   DATE FIX
   ============================================================ */

function fixDate(value) {
    if (!value) return "";

    if (typeof value === "number") {
        const d = XLSX.SSF.parse_date_code(value);
        return `${String(d.m).padStart(2, "0")}/${String(d.d).padStart(2, "0")}/${d.y}`;
    }

    const d = new Date(String(value).trim());
    return isNaN(d) ? "" : d.toLocaleDateString("en-US");
}

/* ============================================================
   MONTHLY TABLE GENERATION
   ============================================================ */

function generateMonthlyTables(aoa) {
    showStatus("processing", "Checking header row…");

    const HEADER_ROW = 8;
    const DATA_START = HEADER_ROW + 1;

    if (!aoa[HEADER_ROW]) {
        showStatus("errorMessage", "Header row 8 not found.");
        return;
    }

    let foundValidRow = false;
    const monthlyData = {};

    for (let r = DATA_START; r < aoa.length; r++) {
        const modality = aoa[r][0];
        const locationFull = aoa[r][1];
        const dosRaw = aoa[r][5];

        const dos = fixDate(dosRaw);
        const loc = LOCATION_MAP[locationFull];

        if (dos && loc) {
            foundValidRow = true;
        }
    }

    if (!foundValidRow) {
        showStatus("errorMessage", "No valid exam rows found.");
        return;
    }

    showStatus("processing", "Building tables…");

    // (Your existing table-building code goes here)
    showStatus("successMessage", "Completed! Tables generated below.");
}

