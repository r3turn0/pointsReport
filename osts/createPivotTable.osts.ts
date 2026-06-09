function main(workbook: ExcelScript.Workbook) {

    const sourceSheet = workbook.getWorksheet("Worker Productivity Report");
    if (!sourceSheet) throw new Error("Source sheet not found");

    let pivotSheet = workbook.getWorksheet("Pivot Table");
    if (pivotSheet) {
        pivotSheet.getUsedRange()?.clear();
    } else {
        pivotSheet = workbook.addWorksheet("Pivot Table");
    }

    // Headers
    pivotSheet.getRange("A1").setValue("Row Labels");
    pivotSheet.getRange("B1").setValue("Sum of Points");

    // ---------------------------------------
    // READ DATA
    // ---------------------------------------
    const usedRange = sourceSheet.getUsedRange();
    const values = usedRange.getValues();

    // Find column indexes dynamically
    const headers = values[0];

    const workerIdx = headers.indexOf("Worker Name");
    const serviceIdx = headers.indexOf("Service");
    const dateIdx = headers.indexOf("Service Date");
    const pointsIdx = headers.indexOf("Points");

    if (workerIdx === -1 || serviceIdx === -1 || dateIdx === -1 || pointsIdx === -1) {
        throw new Error("Required columns not found");
    }

    // ---------------------------------------
    // BUILD PIVOT (Nested Map)
    // ---------------------------------------
    const pivot = new Map<
        string,
        Map<string, Map<string, number>>
    >();

    for (let i = 1; i < values.length; i++) {
        const row = values[i];

        const worker = String(row[workerIdx]).trim();
        const service = String(row[serviceIdx]).trim();
        const dateRaw = row[dateIdx];
        const points = Number(row[pointsIdx]) || 0;

        if (!worker || !service || !dateRaw) continue;

        const date = new Date(dateRaw).toLocaleDateString("en-US");

        if (!pivot.has(worker)) {
            pivot.set(worker, new Map());
        }

        const workerMap = pivot.get(worker)!;

        if (!workerMap.has(date)) {
            workerMap.set(date, new Map());
        }

        const dateMap = workerMap.get(date)!;

        if (!dateMap.has(service)) {
            dateMap.set(service, 0);
        }

        dateMap.set(service, dateMap.get(service)! + points);
    }

    // ---------------------------------------
    // SORT KEYS
    // ---------------------------------------
    const sortedWorkers = Array.from(pivot.keys()).sort();

    // ---------------------------------------
    // WRITE OUTPUT (LIKE PIVOT TABLE)
    // ---------------------------------------
    let rowIndex = 1;

    for (const worker of sortedWorkers) {
        const workerMap = pivot.get(worker)!;

        // Worker row
        pivotSheet.getCell(rowIndex, 0).setValue(worker);
        rowIndex++;

        const sortedDates = Array.from(workerMap.keys()).sort(
            (a, b) => new Date(a).getTime() - new Date(b).getTime()
        );

        for (const date of sortedDates) {
            const dateMap = workerMap.get(date)!;

            // Date row (indented)
            pivotSheet.getCell(rowIndex, 0).setValue("  " + date);
            rowIndex++;

            const sortedServices = Array.from(dateMap.keys()).sort();

            for (const service of sortedServices) {
                const sumPoints = dateMap.get(service)!;

                // Service row (double indented)
                pivotSheet.getCell(rowIndex, 0).setValue("    " + service);
                pivotSheet.getCell(rowIndex, 1).setValue(sumPoints);
                rowIndex++;
            }
        }
    }

    // ---------------------------------------
    // FORMAT (to match pivot look)
    // ---------------------------------------
    pivotSheet.getRange(`A2:B${rowIndex}`)
        .getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    pivotSheet.getRange("A:A").getFormat().autofitColumns();
    pivotSheet.getRange("B:B").getFormat().autofitColumns();
}