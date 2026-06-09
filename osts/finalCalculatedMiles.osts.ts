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
