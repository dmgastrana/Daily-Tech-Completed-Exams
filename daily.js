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
   MAIN ENTRY — RUN DAILY SUMMARY
   ============================================================ */

function runDailySummary() {
    const fileInput = document.getElementById("dailyFile");
    const file = fileInput.files[0];

    // Reset messages
    document.getElementById("processing").style.display = "none";
    document.getElementById("errorMessage").style.display = "none";
    document.getElementById("successMessage").style.display = "none";

    if (!file) {
        document.getElementById("errorMessage").innerText = "No file selected.";
        document.getElementById("errorMessage").style.display = "block";
        return;
    }

    // Show processing
    document.getElementById("processing").style.display = "block";

    const reader = new FileReader();
    const fileName = file.name.toLowerCase();

    reader.onload = function(e) {
        try {
            if (fileName.endsWith(".csv")) {
                parseCSV(e.target.result);
            } else {
                parseExcel(e.target.result);
            }

            document.getElementById("processing").style.display = "none";
            document.getElementById("successMessage").style.display = "block";

        } catch (err) {
            document.getElementById("processing").style.display = "none";
            document.getElementById("errorMessage").innerText = "Error: " + err.message;
            document.getElementById("errorMessage").style.display = "block";
        }
    };

    if (fileName.endsWith(".csv")) {
        reader.readAsText(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}

/* ============================================================
   PARSE CSV — HEADER ROW 8, DATA STARTS ROW 9
   ============================================================ */

function parseCSV(text) {
    const rows = text.split(/\r?\n/).map(r => r.split(","));
    generateMonthlyTables(rows);
}

/* ============================================================
   PARSE EXCEL — HEADER ROW 8, DATA STARTS ROW 9
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

    document.getElementById("errorMessage").style.display = "none";
    document.getElementById("successMessage").style.display = "none";

    const HEADER_ROW = 8;
    const DATA_START = HEADER_ROW + 1;

    if (!aoa || aoa.length <= DATA_START) {
        document.getElementById("errorMessage").innerText = "Error: File is too short.";
        document.getElementById("errorMessage").style.display = "block";
        return;
    }

    if (!aoa[HEADER_ROW] || aoa[HEADER_ROW].length < 7) {
        document.getElementById("errorMessage").innerText = "Error: Header row not found at row 8.";
        document.getElementById("errorMessage").style.display = "block";
        return;
    }

    let foundValidRow = false;
    const monthlyData = {};

    for (let r = DATA_START; r < aoa.length; r++) {
        const modality = String(aoa[r][0] || "").trim();
        const locationFull = String(aoa[r][1] || "").trim();
        const dosRaw = aoa[r][5];
        const apptID = String(aoa[r][6] || "").trim();

        const dos = fixDate(dosRaw);
        if (!dos) continue;

        const loc = LOCATION_MAP[locationFull];
        if (!loc) continue;

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
        document.getElementById("errorMessage").innerText = "Error: No valid exam rows found.";
        document.getElementById("errorMessage").style.display = "block";
        return;
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
