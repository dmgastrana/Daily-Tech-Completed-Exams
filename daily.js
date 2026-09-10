function downloadOutput() {
    const leftTables = document.querySelectorAll("#leftColumn table");
    const rightTables = document.querySelectorAll("#rightColumn table");

    const tables = [...leftTables, ...rightTables];

    if (tables.length === 0) {
        alert("No tables to download.");
        return;
    }

    const wb = XLSX.utils.book_new();

    tables.forEach((table, index) => {

        const ws = XLSX.utils.table_to_sheet(table);

        const rows = table.querySelectorAll("tr");
        let maxCols = 0;

        rows.forEach(row => {
            const cells = row.querySelectorAll("th, td");
            if (cells.length > maxCols) maxCols = cells.length;
        });

        const range = {
            s: { r: 0, c: 0 },
            e: { r: rows.length - 1, c: maxCols - 1 }
        };

        ws['!ref'] = XLSX.utils.encode_range(range);

        // TRUE auto-fit
        ws['!cols'] = [];
        for (let C = 0; C < maxCols; C++) {
            let maxLen = 0;
            for (let R = 0; R <= range.e.r; R++) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[cellAddress];
                if (cell && cell.v != null) {
                    const text = String(cell.v);
                    if (text.length > maxLen) maxLen = text.length;
                }
            }
            ws['!cols'][C] = { wch: maxLen + 2 };
        }

        ws['!freeze'] = { rows: 2 };

        const lastRow = range.e.r;

        for (let R = 0; R <= range.e.r; R++) {
            for (let C = 0; C <= range.e.c; C++) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });

                if (!ws[cellAddress]) {
                    ws[cellAddress] = { t: "s", v: "" };
                }

                ws[cellAddress].s = {
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } },
                        left: { style: "thin", color: { rgb: "000000" } },
                        right: { style: "thin", color: { rgb: "000000" } }
                    },
                    font: {
                        bold: R === lastRow
                    }
                };
            }
        }

        // DYNAMIC sheet naming — finds first real date automatically
        let sheetName;

        if (index === tables.length - 1) {
            sheetName = "All Months Total";
        } else {
            let firstDate = null;

            for (let r = 0; r < rows.length; r++) {
                const cell = rows[r].querySelector("td:first-child");
                if (!cell) continue;

                const value = cell.textContent.trim();
                const d = new Date(value);

                if (!isNaN(d)) {
                    firstDate = d;
                    break;
                }
            }

            if (firstDate) {
                const month = firstDate.toLocaleString("en-US", { month: "long" });
                const year = firstDate.getFullYear();
                sheetName = `${month} ${year}`;
            } else {
                sheetName = `Month_${index + 1}`;
            }
        }

        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSX.writeFile(wb, "Monthly_Completed_Exams.xlsx");
}


