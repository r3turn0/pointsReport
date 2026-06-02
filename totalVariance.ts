function main(workbook: ExcelScript.Workbook) {
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
