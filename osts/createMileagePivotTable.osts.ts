function main(workbook: ExcelScript.Workbook) {
    const outputSheetName = "Mileage Pivot Table";
    // ✅ Get all sheets
    //const sheets = workbook.getWorksheets();
    const sourceSheet = workbook.getWorksheet('Detailed Mileage Analysis R');
    // const ws = workbook.getActiveWorksheet();
    // // ✅ Find the correct sheet (returns Worksheet | undefined)
    // const sourceSheet = ws.getUsedRange().find((s: ExcelScript.Worksheet) => s.getName().toLowerCase().includes("detailed mileage"));
    // const sourceSheet = sheetRange.find(
    //   (s:ExcelScript.Worksheet) =>
    //     s.getName().toLowerCase().includes("detailed mileage")
    // );

    // ✅ Safety check
    if (!sourceSheet) {
        throw new Error("Mileage sheet not found");
    }

    // ✅ Output sheet
    let outputSheet = workbook.getWorksheet(outputSheetName);
    if (!outputSheet) {
        outputSheet = workbook.addWorksheet(outputSheetName);
    } else {
        outputSheet.getUsedRange()?.clear();
    }


    // === READ DATA ===
    const data = sourceSheet.getUsedRange().getValues();
    const headers = data[0];

    const workerIdx = headers.indexOf("Worker Name");
    const dateIdx = headers.indexOf("Date");
    const locIdx = headers.indexOf("End Location");
    const milesIdx = headers.indexOf("Final Calculated Miles");

    // === BUILD NESTED STRUCTURE ===
    const tree = new Map<
        string,
        Map<string, Map<string, number>>>();

    for (let i = 1; i < data.length; i++) {
        const worker = String(data[i][workerIdx]);
        const date = String(data[i][dateIdx]);
        const location = String(data[i][locIdx]);
        const miles = Number(data[i][milesIdx]) || 0;

        if (!tree.has(worker)) {
            tree.set(worker, new Map());
        }

        const dates = tree.get(worker)!;

        if (!dates.has(date)) {
            dates.set(date, new Map());
        }

        const locations = dates.get(date)!;

        if (!locations.has(location)) {
            locations.set(location, 0);
        }

        locations.set(location, locations.get(location)! + miles);
    }

    // === SORT KEYS ===
    const workers = Array.from(tree.keys()).sort();

    const output: (string | number)[][] = [];

    // Header row
    output.push(["Row Labels", "Sum of Final Calculated Miles"]);

    let grandTotal = 0;

    // === BUILD PIVOT LAYOUT ===
    for (const worker of workers) {
        const dates = tree.get(worker)!;
        let workerTotal = 0;

        output.push([worker, ""]);

        const sortedDates = Array.from(dates.keys()).sort();

        for (const date of sortedDates) {
            const locations = dates.get(date)!;
            let dateTotal = 0;

            output.push([`  ${date}`, ""]);

            const sortedLocations = Array.from(locations.keys()).sort();

            for (const loc of sortedLocations) {
                const val = locations.get(loc)!;

                output.push([`    ${loc}`, val]);

                dateTotal += val;
            }

            // Date subtotal
            output.push([`  ${date} Total`, dateTotal]);

            workerTotal += dateTotal;
        }

        // Worker subtotal
        output.push([`${worker} Total`, workerTotal]);

        grandTotal += workerTotal;
    }

    // Grand total
    output.push(["Grand Total", grandTotal]);

    // === WRITE TO SHEET ===
    const range = outputSheet.getRangeByIndexes(
        0,
        0,
        output.length,
        output[0].length
    );

    range.setValues(output);

    // === FORMATTING ===
    outputSheet.getRange("A1:B1").getFormat().getFont().setBold(true);

    // Bold totals & worker rows
    for (let i = 1; i < output.length; i++) {
        const label = String(output[i][0]);

        if (
            !label.startsWith(" ") || // worker row
            label.includes("Total")   // totals
        ) {
            outputSheet.getRange(`A${i + 1}:B${i + 1}`)
                .getFormat()
                .getFont()
                .setBold(true);
        }
    }

    outputSheet.getUsedRange().getFormat().autofitColumns();
}