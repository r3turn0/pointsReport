function main(workbook: ExcelScript.Workbook) {
  // Find worksheet that contains "Detailed Mileage" in its name
  const sheets = workbook.getWorksheets();  
  let sheet: ExcelScript.Worksheet = null;
  for (const s of workbook.getWorksheets()) {
    if (s.getName().toLowerCase().includes("detailed mileage")) {
      sheet = s;
      break;
    }
  }
  if (!sheet) {
    throw new Error('No worksheet found containing "Detailed Mileage".');
  }
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    throw new Error("No data found in worksheet.");
  }
  const values = usedRange.getValues();
  const headers = values[0] as string[];
  // Find column indexes by header name
  const calculatedMilesCol = headers.findIndex(
    h => String(h).trim().toLowerCase() === "calculated miles"
  );
  const adjustmentCol = headers.findIndex(
    h => String(h).trim().toLowerCase() === "adjustment"
  );
  if (calculatedMilesCol === -1) {
    throw new Error('Column "Calculated Miles" not found.');
  }
  if (adjustmentCol === -1) {
    throw new Error('Column "Adjustment" not found.');
  }
  // Add new column header AFTER Adjustment
  const finalColIndex = adjustmentCol + 1;
  headers.splice(finalColIndex, 0, "Final Calculated Miles");
  // Process rows
  for (let i = 1; i < values.length; i++) {
    const calculatedMiles = Number(values[i][calculatedMilesCol]) || 0;
    const adjustment = Number(values[i][adjustmentCol]) || 0;
    const finalMiles = Number((calculatedMiles + adjustment).toFixed(2));
    values[i].splice(finalColIndex, 0, finalMiles);
  }
  // Write data back to sheet
  const targetRange = sheet.getRangeByIndexes(
    0,
    0,
    values.length,
    values[0].length
  );
  targetRange.setValues(values);
  // Format new column to always show 2 decimals
  targetRange
    .getColumn(finalColIndex)
    .setNumberFormat("0.00");
}
function main2(workbook: ExcelScript.Workbook) {
    // Find worksheet that contains "Detailed Mileage" in its name
    const sheets = workbook.getWorksheets();
    let sheet: ExcelScript.Worksheet = null;
    for (const s of workbook.getWorksheets()) {
        if (s.getName().toLowerCase().includes("worker productivity report")) {
            sheet = s;
            break;
        }
    }
    if (!sheet) {
        console.log("No Worksheet found containing 'Worker Productivity Report' in its name.");
        throw new Error('No worksheet found containing "Worker Productivity Report".');
    }
    const usedRange = sheet.getUsedRange();
    if (!usedRange) {
        console.log("No data found in worksheet.");
        throw new Error("No data found in worksheet.");
    }
    const values = usedRange.getValues();
    const headers = values[0] as string[];
    // Find column indexes by header name
    const totalPointsCol = headers.findIndex(
        h => String(h).trim().toLowerCase() === "total points earned"
    );
    const expectedPointsCol = headers.findIndex(
        h => String(h).trim().toLowerCase() === "expected points"
    );
    let totalVarCol = headers.findIndex(
        h => String(h).trim().toLowerCase() === "total variance"
    );
    if (totalPointsCol === -1) {
        console.log('Column "Total Points Earned" not found.');
        throw new Error('Column "Total Points Earned" not found.');
    }
    if (expectedPointsCol === -1) {
        console.log('Column "Expected Points" not found.');
        throw new Error('Column "Expected Points" not found.');
    }
    if (totalVarCol !== -1) {
        // update existing column
        for (let i = 1; i < values.length; i++) {
            const totalPoints = Number(values[i][totalPointsCol]) || 0;
            const expectedPoints = Number(values[i][expectedPointsCol]) || 0;
            values[i][totalVarCol] = Number((totalPoints - expectedPoints).toFixed(2));
        }
    }
    else {
        // create new column
        const totalVarColIndex = totalPointsCol + 1;
        values[0].splice(totalVarColIndex, 0, "Total Variance");
        for (let i = 1; i < values.length; i++) {
            const totalPoints = Number(values[i][totalPointsCol]) || 0;
            const expectedPoints = Number(values[i][expectedPointsCol]) || 0;
            values[i].splice(totalVarColIndex, 0, Number((totalPoints - expectedPoints).toFixed(2)));
        }
        totalVarCol = totalVarColIndex;
    }
    // write once
    const targetRange = sheet.getRangeByIndexes(0, 0, values.length, values[0].length);
    targetRange.setValues(values);
    // format
    targetRange.getColumn(totalVarCol).setNumberFormat("0.00");
}
function main3(workbook: ExcelScript.Workbook) {
  // ---------------------------------------
  // Source Sheet
  // ---------------------------------------
  const sourceSheet = workbook.getWorksheet("Worker Productivity Report");
  if (!sourceSheet) {
    throw new Error("Source sheet not found");
  }

  // ---------------------------------------
  // Create / Reset Pivot Sheet
  // ---------------------------------------
  let pivotSheet = workbook.getWorksheet("Pivot Table");
  if (pivotSheet) {
    pivotSheet.getUsedRange()?.clear();
  } else {
    pivotSheet = workbook.addWorksheet("Pivot Table");
  }

  // ---------------------------------------
  // Headers
  // ---------------------------------------
  pivotSheet.getRange("A1").setValue("Row Labels");
  pivotSheet.getRange("B1").setValue("Sum of Points");

  // Helper headers
  pivotSheet.getRange("E1").setValue("Worker");
  pivotSheet.getRange("F1").setValue("Date");
  pivotSheet.getRange("G1").setValue("Service");
  pivotSheet.getRange("H1").setValue("SumPoints");

  // ---------------------------------------
  // ASSUMPTION:
  // Columns in source:
  // A = Worker Name
  // B = Service
  // C = Service Date
  // D = Points
  // ---------------------------------------

  // ---------------------------------------
  // UNIQUE COMBINATIONS (Pivot grouping)
  // ---------------------------------------
  pivotSheet.getRange("E2").setFormula(`
=UNIQUE(
    CHOOSE({1,2,3},
        'Worker Productivity Report'!A:A,
        'Worker Productivity Report'!C:C,
        'Worker Productivity Report'!B:B
    )
)
`);

  // ---------------------------------------
  // SUMIFS (Aggregation)
  // ---------------------------------------
  pivotSheet.getRange("H2").setFormula(`
=SUMIFS(
    'Worker Productivity Report'!D:D,
    'Worker Productivity Report'!A:A, E2,
    'Worker Productivity Report'!C:C, F2,
    'Worker Productivity Report'!B:B, G2
)
`);

  // Autofill formulas down
  pivotSheet.getRange("H2").autoFill("H2:H1000", ExcelScript.AutoFillType.fillDefault);

  // ---------------------------------------
  // ROW LABELS (Hierarchy formatting)
  // ---------------------------------------
  pivotSheet.getRange("A2").setFormula(`
=IF(E2<>E1,
    E2,
    IF(F2<>F1,
        "  " & TEXT(F2,"mm/dd/yyyy"),
        "    " & G2
    )
)
`);

  pivotSheet.getRange("A2").autoFill("A2:A1000", ExcelScript.AutoFillType.fillDefault);

  // ---------------------------------------
  // VALUES COLUMN
  // ---------------------------------------
  pivotSheet.getRange("B2").setFormula(`=H2`);
  pivotSheet.getRange("B2").autoFill("B2:B1000", ExcelScript.AutoFillType.fillDefault);

  // ---------------------------------------
  // Optional Sorting (mimics Pivot behavior)
  // ---------------------------------------
  const sortRange = pivotSheet.getRange("E2:H1000");
  sortRange.getSort().apply([
    { key: 0, ascending: true }, // Worker
    { key: 1, ascending: true }, // Date
    { key: 2, ascending: true }  // Service
  ]);

  // ---------------------------------------
  // Auto-fit columns
  // ---------------------------------------
  pivotSheet.getUsedRange()?.getFormat().autofitColumns();
}
function main4(workbook: ExcelScript.Workbook) {
  const sourceSheetName = "Detail Mileage Analysis";
  const outputSheetName = "Mileage Pivot (Exact)";

  const sourceSheet = workbook.getWorksheet(sourceSheetName);

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
function main5(workbook: ExcelScript.Workbook) {
  const sourceSheetName = "Worker_Productivity_Data";
  const matrixSheetName = "Offer_Letter_Matrix";

  const sourceSheet = workbook.getWorksheet(sourceSheetName);

  // ✅ Guard source sheet
  if (!sourceSheet) {
    throw new Error(`Source sheet '${sourceSheetName}' not found.`);
  }

  const data = sourceSheet.getUsedRange().getValues();
  const headers = data[0];

  const workerIdx = headers.indexOf("Worker Name");
  const dateIdx = headers.indexOf("Service Date");
  const visitIdx = headers.indexOf("Visit Type");
  const pointsIdx = headers.indexOf("Points");

  // ================================
  // ✅ Helpers
  // ================================
  
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
    return name.replace(/[\\/*?:[\]]/g, "").substring(0, 31);
  }

  // ================================
  // ✅ FIXED: Load Thresholds SAFELY (line 33 bug fixed)
  // ================================
  const matrixSheet = workbook.getWorksheet(matrixSheetName);

  let thresholdMap = new Map<string, number>();

  if (matrixSheet) {
    const matrixData = matrixSheet.getUsedRange().getValues();
    const matrixHeaders = matrixData[0];

    const nameIdx = matrixHeaders.indexOf("Employee Name");
    const thresholdIdx = matrixHeaders.indexOf("Threshold (pt)");

    for (let i = 1; i < matrixData.length; i++) {
      thresholdMap.set(
        String(matrixData[i][nameIdx]),
        Number(matrixData[i][thresholdIdx])
      );
    }
  } else {
    console.log(
      `Sheet '${matrixSheetName}' not found. Defaulting thresholds to 0.`
    );
  }

  // ================================
  // ✅ Group by worker (O(N))
  // ================================
  const workerMap = new Map<string, (string | number | boolean)[][]>();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const worker = String(row[workerIdx]);

    if (!workerMap.has(worker)) {
      workerMap.set(worker, []);
    }

    workerMap.get(worker)!.push(row);
  }

  const workers = Array.from(workerMap.keys()).sort();

  // ================================
  // ✅ Process each worker
  // ================================
  for (const worker of workers) {
    const safeName = safeSheetName(worker);

    let sheet = workbook.getWorksheet(safeName);
    if (!sheet) {
      sheet = workbook.addWorksheet(safeName);
    } else {
      sheet.getUsedRange()?.clear();
    }

    const rows = workerMap.get(worker)!;

    // ✅ Sort safely
    rows.sort((a, b) =>
      parseExcelDate(a[dateIdx]) - parseExcelDate(b[dateIdx])
    );

    // ================================
    // ✅ Group by date
    // ================================
    const grouped = new Map<string, (string | number | boolean)[][]>();

    for (const r of rows) {
      const d = formatDate(r[dateIdx]);

      if (!grouped.has(d)) {
        grouped.set(d, []);
      }

      grouped.get(d)!.push(r);
    }

    const sortedDates = Array.from(grouped.keys()).sort()

    // ================================
    // ✅ Compute totals
    // ================================
    let totalPoints = 0;
    for (const r of rows) {
      totalPoints += Number(r[pointsIdx]) || 0;
    }

    const threshold = thresholdMap.get(worker) ?? 0;

    // ================================
    // ✅ Build EXACT report structure
    // ================================
    let output: (string | number)[][] = [];

    // ✅ Header (matches your Excel file)
    output.push([`${worker} Weekly Points Report`]);
    output.push([
      `Bonus per visit for anything over ${threshold} points.`
    ]);
    output.push(["Week 04/12/26 - 04/18/26"]);
    output.push([]);

    // ✅ Summary
    output.push(["Bonus Points", 0]);
    output.push(["Bonus Amount", 0]);
    output.push([]);

    // ✅ Table
    output.push(["Row Labels", "Sum of Points"]);

    // ✅ Worker total row (critical fix vs your old script)
    output.push([worker, totalPoints]);

    // ================================
    // ✅ Render data (NO date subtotals)
    // ================================
    for (const date of sortedDates) {
      output.push([date, ""]);

      const visitRows = grouped.get(date)!;

      for (const r of visitRows) {
        output.push([
          `  ${r[visitIdx]}`,
          Number(r[pointsIdx]) || 0
        ]);
      }
    }

    // ================================
    // ✅ Write to sheet
    // ================================
    const range = sheet.getRangeByIndexes(
      0,
      0,
      output.length,
      output[0].length
    );

    range.setValues(output);

    // ================================
    // ✅ Formatting
    // ================================
    for (let i = 0; i < output.length; i++) {
      const label = String(output[i][0]);

      if (
        i === 0 || // title
        label === "Row Labels" ||
        label === worker ||
        !label.startsWith(" ")
      ) {
        sheet.getRange(`A${i + 1}:B${i + 1}`)
          .getFormat()
          .getFont()
          .setBold(true);
      }
    }

    sheet.getUsedRange().getFormat().autofitColumns();
  }
}