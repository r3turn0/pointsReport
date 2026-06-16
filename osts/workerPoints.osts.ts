function main(workbook: ExcelScript.Workbook) {
    const sourceSheetName = "Worker Productivity";
    const matrixSheetName = "Offer Letter";
    const workSheets = workbook.getWorksheets();
    let sourceSheet: ExcelScript.Worksheet = null;
    let matrixSheet: ExcelScript.Worksheet = null;
    for (var i = 0; i < workSheets.length; i++) {
        if (workSheets[i].getName().toLowerCase().includes(sourceSheetName.toLowerCase())) {
            sourceSheet = workSheets[i];
            break;
        }
    }
    for (var i = 0; i < workSheets.length; i++) {
        if (workSheets[i].getName().toLowerCase().includes(matrixSheetName.toLowerCase())) {
            matrixSheet = workSheets[i];
            break;
        }
    }
    if (!sourceSheet) {
        throw new Error(`Source sheet '${sourceSheetName}' not found.`);
    }

    if (!matrixSheet) {
        throw new Error(`Matrix sheet '${matrixSheetName}' not found.`);
    }
    const data = sourceSheet.getUsedRange().getValues();
    const headers = data[0];
    const workerIdx = headers.indexOf("Worker Name");
    const serviceIdx = headers.indexOf("Service");
    const dateIdx = headers.indexOf("Service Date");
    //const visitIdx = headers.indexOf("Visit Type");
    const pointsIdx = headers.indexOf("Points");
    if (workerIdx === -1 || serviceIdx === -1 || dateIdx === -1 || pointsIdx === -1) {
        throw new Error("Column headers not found. Check exact names.");
    }
    // Helpers
    type ExcelValue = string | number | boolean;
    function parseExcelDate(value: ExcelValue): number {
        if (typeof value === "number") {
            return new Date((value - 25569) * 86400 * 1000).getTime();
        }

        return new Date(String(value)).getTime();
    }
    function formatDate(value: ExcelValue): string {
        const d = new Date(parseExcelDate(value));
        return d.toLocaleDateString("en-US");
    }
    function safeSheetName(name: string): string {
        return name.replace(/[\\\/\*\?:\[\]]/g, "").substring(0, 31);
    }
    function normalizeWorkerName(name: string): string {
        name = name.toLowerCase().trim();
        const match = name.match(/^([^,]+),([^(]+)\(([^)]+)\)$/);
        if (match) {
            const last = match[1].trim();
            let first = match[2].trim();
            const role = match[3].trim();

            // Remove middle initials if present
            first = first.split(" ")[0];

            return `${first} ${last}, ${role}`;
        }
        // Normalize matrix names too
        const parts = name.split(",");
        if (parts.length === 2) {
            const namePart = parts[0].trim();
            const role = parts[1].trim();

            // Remove middle initials
            const namePieces = namePart.split(" ");
            const first = namePieces[0];
            const last = namePieces[namePieces.length - 1];

            return `${first} ${last}, ${role}`;
        }
        return name;
    }
    // Load each threshold
    // Employee-specific goal threshold and bonus rate
    const thresholdMap = new Map<string, number>();
    const bonusPerPointMap = new Map<string, number>();
    if (matrixSheet) {
        const matrixData = matrixSheet.getUsedRange().getValues();
        const matrixHeaders = matrixData[0];
        const nameIdx = matrixHeaders.indexOf("Employee");
        const bonusIdx = matrixHeaders.indexOf("Bonus $/pt over goal");
        const thresholdIdx = matrixHeaders.indexOf("Threshold (pt)");
        for (let i = 1; i < matrixData.length; i++) {
            const workerName = normalizeWorkerName(String(matrixData[i][nameIdx]));
            thresholdMap.set(workerName, Number(matrixData[i][thresholdIdx]));
            bonusPerPointMap.set(workerName, Number(matrixData[i][bonusIdx]));
        }
    } else {
        throw new Error(`Sheet '${matrixSheetName}' not found.`);
    }
    // Group by worker (O(N))
    type variant = (string | number | boolean)[][];
    const workerMap = new Map<string, variant>();

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const worker = normalizeWorkerName(String(row[workerIdx]));
        if (!workerMap.has(worker)) {
            workerMap.set(worker, []);
        }
        workerMap.get(worker)!.push(row);
    }
    const workers = Array.from(workerMap.keys()).sort();
    // Process each worker
    console.log('Process each worker');
    console.log("Normalized source:", Array.from(workerMap.keys()));
    console.log("Normalized matrix:", Array.from(thresholdMap.keys()));
    for (const worker of workers) {
        if (!thresholdMap.has(worker)) {
            console.log(`No threshold found for: ${worker}`);
            continue;
        }
        const safeName = safeSheetName(worker);
        let sheet = workbook.getWorksheet(safeName);
        if (!sheet) {
            sheet = workbook.addWorksheet(safeName);
        } else {
            sheet.getUsedRange()?.clear();
        }
        const rows = workerMap.get(worker)!;
        const timestamps = rows.map(r => parseExcelDate(r[dateIdx]));
        const minDate = new Date(Math.min(...timestamps));
        const maxDate = new Date(Math.max(...timestamps));
        const weekLabel = `Week ${minDate.toLocaleDateString("en-US")} - ${maxDate.toLocaleDateString("en-US")}`;
        // Sort safely
        rows.sort((a, b) =>
            parseExcelDate(a[dateIdx]) - parseExcelDate(b[dateIdx])
        );
        // Group by date
        const grouped = new Map<string, variant>();
        for (const r of rows) {
            const d = formatDate(r[dateIdx]);
            if (!grouped.has(d)) {
                grouped.set(d, []);
            }
            grouped.get(d)!.push(r);
        }
        const sortedDates = Array.from(grouped.keys()).sort(
            (a, b) => new Date(a).getTime() - new Date(b).getTime()
        );
        // Compute totals
        let totalPoints = 0;
        for (const r of rows) {
            totalPoints += Number(r[pointsIdx]) || 0;
        }
        const bonusPerPoint = bonusPerPointMap.get(worker) ?? 0;
        const threshold = thresholdMap.get(worker) ?? 0;
        console.log('bonus points per worker:', worker,bonusPerPoint);
        console.log('threshold per worker:', threshold);
        // Build EXACT report structure
        let output: (string | number)[][] = [];
        // Header (matches your Excel file)
        output.push([`${worker} Weekly Points Report`]);
        output.push([
            `Bonus per visit for anything over ${bonusPerPoint} points.`
        ]);
        output.push([weekLabel, ""]);
        output.push(["", ""]);
        // Summary
        output.push(["Bonus Points", 0]);
        output.push(["Bonus Amount", 0]);
        output.push(["", ""]);
        // Table
        output.push(["Row Labels", "Sum of Points"]);
        
        // Worker total row (critical fix vs your old script)
        output.push([worker, totalPoints]);
        // Render Data
        for (const date of sortedDates) {
            output.push([date, ""]);
            const visitRows = grouped.get(date)!;
            for (const r of visitRows) {
                output.push([
                    `  ${r[serviceIdx]}`,
                    Number(r[pointsIdx]) || 0
                ]);
            }
        }
        const colCount = Math.max(...output.map(r => r.length));
        const normalized = output.map(r => {
            const row = [...r];
            while (row.length < colCount) {
                row.push("");
            }
            return row;
        });
        // Write to sheet
        const range = sheet.getRangeByIndexes(
            0,
            0,
            normalized.length,
            colCount
        );
        range.setValues(normalized);
        // Formatting
        for (let i = 0; i < output.length; i++) {
            const label = String(output[i][0]);
            if (
                i === 0 || // title
                label === "Row Labels" ||
                label === worker ||
                !label.startsWith(" ")
            ) {
                sheet.getRange(`A${i + 1}:B${i + 1}`).getFormat().getFont().setBold(true);
            }
        }
        sheet.getUsedRange().getFormat().autofitColumns();
        // ---- CREATE TABLE FROM OUTPUT ----
        let headerRowIndex = -1;
        for (let i = 0; i < normalized.length; i++) {
            if (normalized[i][0] === "Row Labels") {
                headerRowIndex = i;
                break;
            }
        }
        if (headerRowIndex !== -1) {
            const totalRows = normalized.length - headerRowIndex;

            const tableRange = sheet.getRangeByIndexes(
                headerRowIndex,
                0,
                totalRows,
                colCount
            );

            // Remove existing tables (prevents duplicate name issues)
            sheet.getTables().forEach(t => t.delete());

            const table = sheet.addTable(tableRange, true);

            // Table name must be safe (no spaces/special chars)
            const safeTableName = safeName.replace(/[^a-zA-Z0-9]/g, "_");

            table.setName(safeTableName);
        } else {
            console.log(`No table header found for ${worker}`);
        }
    }
}
