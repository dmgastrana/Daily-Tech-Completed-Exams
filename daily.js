/* ============================================================
   CONFIGURATION — EXACT TABLE FORMAT
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
   MAIN ENTRY — DETECT CSV OR EXCEL
   ============================================================ */

function runDailySummary() {
    const fileInput = document.getElementById("dailyFile");
    const file = fileInput.files[0];

    if (!file) {
        alert("Please upload a file.");
        return;
    }

    const reader = new FileReader();
    const fileName = file.name.toLowerCase();

    if (fileName.endsWith(".csv")) {
        reader.onload = function(e) {
            parseCSV(e.target.result);
        };
        reader.readAsText(file);
    } else {
        reader.onload = function(e) {
            parseExcel(e.target.result);
        };
        reader.readAsArrayBuffer(file);
    }
}

/* ============================================================
   PARSE CSV — HEADER IS ROW 8, DATA STARTS ROW 9
   ============================================================ */

function parseCSV(text) {
    const rows = text.split(/\r?\n/).map(r => r.split(","));
    generateMonthlyTables(rows);
}

/* ============================================================
   PARSE EXCEL — HEADER IS ROW 8, DATA STARTS ROW 9
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
    const monthlyData = {};

    const HEADER_ROW = 8;
    const DATA_START = HEADER_ROW + 1;

    for (let r = DATA_START; r < aoa.length; r++) {
        const modality = String(aoa[r][0] || "").trim(); // Column A
        const locationFull = String(aoa[r][1] || "").trim(); // Column B
        const dosRaw = aoa[r][5]; // Column F
        const apptID = String(aoa[r][6] || "").trim(); // Column G

        const dos = fixDate(dosRaw);
        if (!dos) continue;

        const loc = LOCATION_MAP[locationFull];
        if (!loc) continue;

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

    displayMonthlyTables(monthlyData);
}

/* ============================================================
   DISPLAY MONTHLY TABLES
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
   BUILD EXACT TABLE FORMAT
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

    let totalRow = "<tr><td>Grand Total</td>";
    let monthGrandTotal = 0;

    Object.keys(TABLE_STRUCTURE).forEach(loc => {
        let locMonthTotal = 0;

        dates.forEach(dos => {
            locMonthTotal += monthData[dos][loc].apptCount || 0;
        });

        TABLE_STRUCTURE[loc].forEach(mod => {
            let modTotal = 0;

            dates.forEach(dos => {
                modTotal +=
                    monthData[dos][loc] && monthData[dos][loc][mod]
                        ? monthData[dos][loc][mod]
                        : 0;
            });

            totalRow += `<td>${modTotal || ""}</td>`;
        });

        totalRow += `<td>${locMonthTotal || ""}</td>`;
        monthGrandTotal += locMonthTotal;
    });

    totalRow += `<td>${monthGrandTotal || ""}</td></tr>`;
    table.innerHTML += totalRow;

    return table;
}

/* ============================================================
   DOWNLOAD OUTPUT 
   ============================================================ */

function downloadOutput() {
    alert("Monthly tables are visual only. No text output generated.");
}
