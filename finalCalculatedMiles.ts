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
    console.log("No Worksheet found containing 'Detailed Mileage' in its name.");
    throw new Error('No worksheet found containing "Detailed Mileage".');
  }
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    console.log("No data found in worksheet.");
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
  let finalMilesCol = headers.findIndex(
    h => String(h).trim().toLowerCase() === "final calculated miles"
  );
  if (calculatedMilesCol === -1) {
    console.log('Column "Calculated Miles" not found.');
    throw new Error('Column "Calculated Miles" not found.');
  }
  if (adjustmentCol === -1) {
    console.log('Column "Adjustment" not found.');
    throw new Error('Column "Adjustment" not found.');
  }
  if (finalMilesCol !== -1) {
  // update existing column
  for (let i = 1; i < values.length; i++) {
      const calculatedMiles = Number(values[i][calculatedMilesCol]) || 0;
      const adjustment = Number(values[i][adjustmentCol]) || 0;
      values[i][finalMilesCol] = Number((calculatedMiles + adjustment).toFixed(2));
    }
  } 
  else {
  // create new column
  const finalColIndex = adjustmentCol + 1;
  values[0].splice(finalColIndex, 0, "Final Calculated Miles");
  for (let i = 1; i < values.length; i++) {
    const calculatedMiles = Number(values[i][calculatedMilesCol]) || 0;
    const adjustment = Number(values[i][adjustmentCol]) || 0;
    values[i].splice(finalColIndex, 0, Number((calculatedMiles + adjustment).toFixed(2)));
  }
  finalMilesCol = finalColIndex;
}
// write once
const targetRange = sheet.getRangeByIndexes(0, 0, values.length, values[0].length);
targetRange.setValues(values);
// format
targetRange.getColumn(finalMilesCol).setNumberFormat("0.00");
}
