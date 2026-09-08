/* ============================================================
   LOCATION + MODALITY TABLE STRUCTURE
   ============================================================ */

const LOCATION_MAP = {
    "Diagnostic Medical Group Arcadia": "AR",
    "Diagnostic Medical Group City of Industry": "CI",
    "Diagnostic Medical Group Monterey Park": "MP",
    "Diagnostic Medical Group San Gabriel": "SG",
    "Synergy San Gabriel": "SSG"
};

const TABLE_STRUCTURE = {
    AR: ["CT", "DEXA", "MG", "US", "XR"],
    CI: ["CT", "DEXA", "MG", "MR", "US", "XR"],
    MP: ["CT", "DEXA", "MG", "MR", "US", "XR"],
    SG: ["CT", "ECG", "MG", "MR", "US", "XR"],
    SSG: ["CT", "ECG", "MG", "MR", "US", "XR"]
};

/* ============================================================
   STATUS HANDLING
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

    const reader = new FileReader();
    const fileName = file.name.toLowerCase();

    reader.onload = function(e) {
        let text = e.target.result;

        if (fileName.endsWith(".csv")) {
            parseCSV(text);
        } else {
            parseExcel(e.target.result);
        }
    };

    if (fileName.endsWith(".csv")) {
        reader.readAsText(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}

/* ============================================================
   PARSE CSV — UTF‑16 + TAB DELIMITED
   ============================================================ */

function parseCSV(text) {
    // Detect UTF‑16 LE BOM
    if (text.charCodeAt(0) === 0xFEFF || text.charCodeAt(1) === 0x00) {
        const decoder = new TextDecoder("utf-16le");
        const uint8 = new Uint8Array(text.length * 2);
        for (let i = 0; i < text.length; i++) {
            uint8[i * 2] = text.charCodeAt(i) & 0xFF;
            uint8[i * 2 + 1] = text.charCodeAt(i) >> 8;
        }
        text = decoder.decode(uint8);
    }

    // Split rows
    const rows = text.split(/\r?\n/);

    // Split columns by TAB
    const aoa = rows.map(r => r.split("\t"));

    generateMonthlyTables(aoa);
}

/* ============================================================
   PARSE EXCEL
   ============================================================ */

function parseExcel(buffer) {
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
    const d = new Date(String(value).trim());
    return isNaN(d) ? "" : d.toLocaleDateString("en-US");
}

/* ============================================================
   MAIN TABLE GENERATION
   ============================================================ */

function generateMonthlyTables(aoa) {
    const HEADER_ROW = 8;
    const DATA_START = HEADER_ROW + 1;

    if (!aoa[HEADER_ROW]) {
        showStatus("errorMessage", "Header row 8 not found.");
        return;
    }

    let foundValidRow = false;
    const monthlyData = {};

    for (let r = DATA_START; r < aoa.length; r++) {
        const row = aoa[r];
        if (!row || row.length < 7) continue;

        const modality = String(row[0] || "").trim();
        const locationFull = String(row[1] || "").trim();
        const dosRaw = row[5];
        const apptID = String(row[6] || "").trim();

        const dos = fixDate(dosRaw);
        const loc = LOCATION_MAP[locationFull];

        if (!dos || !loc || !modality) continue;

        foundValidRow = true;

        const dateObj = new Date(dos);
        const monthKey = `${dateObj.getFullYear()}-${dateObj.getMonth() + 1}`;

        if (!monthlyData[monthKey]) monthlyData[monthKey] = {};
        if (!monthlyData[monthKey][dos]) monthlyData[monthKey][dos] = {};
        if (!monthlyData[monthKey][dos][loc]) monthlyData[monthKey][dos][loc] = {};

        monthlyData[monthKey][dos][loc].apptCount =
            (monthlyData[monthKey][dos][loc].apptCount || 0) + 1;

        monthlyData[monthKey][dos][loc][modality] =
            (monthlyData[monthKey][dos][loc][modality] || 0) + 1;
    }

    if (!foundValidRow) {
        showStatus("errorMessage", "No valid exam rows found.");
        return;
    }

    displayMonthlyTables(monthlyData);
    showStatus("successMessage", "Completed! Tables generated below.");
}

/* ============================================================
   DISPLAY TABLES
   ============================================================ */

function displayMonthlyTables(monthlyData) {
    const left = document.getElementById("leftColumn");
    const right = document.getElementById("rightColumn");

    left.innerHTML = "";
    right.innerHTML = "";

    Object.keys(monthlyData)
        .sort()
        .forEach(monthKey => {
            const dates = Object.keys(monthlyData[monthKey]).sort(
                (a, b) => new Date(a) - new Date(b)
            );

            const firstDOS = dates[0];
            const lastDOS = dates[dates.length - 1];

            const title = `${firstDOS} – ${lastDOS}`;
            const table = buildMonthlyTable(monthlyData[monthKey], dates);

            const wrapper = document.createElement("div");
            wrapper.style.width = "100%";
            wrapper.style.marginBottom = "40px";

            const titleEl = document.createElement("div");
            titleEl.className = "titleRow";
            titleEl.textContent = title;

            wrapper.appendChild(titleEl);
            wrapper.appendChild(table);

            left.appendChild(wrapper);
        });
}

/* ============================================================
   BUILD TABLE
   ============================================================ */

function buildMonthlyTable(monthData, dates) {
    const table = document.createElement("table");

    let header = "<tr><th>Date</th>";

    Object.keys(TABLE_STRUCTURE).forEach(loc => {
        TABLE_STRUCTURE[loc].forEach(mod => {
            header += `<th>${loc} ${mod}</th>`;
        });
        header += `<th>${loc} Total</th>`;
    });

    header += "<th>Grand Total</th></tr>";
    table.innerHTML = header;

    dates.forEach(dos => {
        let row = `<tr><td>${dos}</td>`;
        let grandTotal = 0;

        Object.keys(TABLE_STRUCTURE).forEach(loc => {
            let locTotal = monthData[dos][loc].apptCount || 0;

            TABLE_STRUCTURE[loc].forEach(mod => {
                const val =
                    monthData[dos][loc] && monthData[dos][loc][mod]
                        ? monthData[dos][loc][mod]
                        : "";
                row += `<td>${val}</td>`;
            });

            row += `<td>${locTotal || ""}</td>`;
            grandTotal += locTotal;
        });

        row += `<td>${grandTotal || ""}</td></tr>`;
        table.innerHTML += row;
    });

    return table;
}

/* ============================================================
   DOWNLOAD OUTPUT
   ============================================================ */

function downloadOutput() {
    alert("Monthly tables are visual only.");
}

