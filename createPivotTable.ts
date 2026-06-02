
function main(workbook: ExcelScript.Workbook) {
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
