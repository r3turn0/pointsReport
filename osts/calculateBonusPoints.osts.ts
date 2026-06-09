function main(workbook: ExcelScript.Workbook) {
  const sheets = workbook.getWorksheets();

  sheets.forEach((sheet) => {
    const usedRange = sheet.getUsedRange();
    if (!usedRange) return;

    const values = usedRange.getValues();

    // --- Headers ---
    const headers = values[0] as string[];

    const expectedCol = headers.indexOf("Expected Points");
    const totalCol = headers.indexOf("Total Points Earned");
    const workerCol = headers.indexOf("Worker Name");

    if (expectedCol === -1 || totalCol === -1) return;

    // --- Add columns if missing ---
    let bonusPointsCol = headers.indexOf("Bonus Points");
    let bonusAmountCol = headers.indexOf("Bonus Amount");

    if (bonusPointsCol === -1) {
      bonusPointsCol = headers.length;
      sheet.getCell(0, bonusPointsCol).setValue("Bonus Points");
    }

    if (bonusAmountCol === -1) {
      bonusAmountCol = headers.length + 1;
      sheet.getCell(0, bonusAmountCol).setValue("Bonus Amount");
    }

    // --- Get worker name (assume same for sheet) ---
    const workerName = values[1]?.[workerCol] as string;

    // --- Bonus rate lookup (Offer Letter Terms logic) ---
    const bonusRate = getBonusRate(workerName);

    // ✅ Set A2 value (no formula)
    sheet.getRange("A2").setValue(bonusRate);

    // --- Process rows ---
    for (let i = 1; i < values.length; i++) {
      const expected = Number(values[i][expectedCol]) || 0;
      const total = Number(values[i][totalCol]) || 0;

      // ✅ Bonus Points calculation
      const bonusPoints = Math.max(0, total - expected);

      // ✅ Bonus Amount calculation
      const bonusAmount = bonusPoints * bonusRate;

      // Write values (NOT formulas)
      sheet.getCell(i, bonusPointsCol).setValue(bonusPoints);
      sheet.getCell(i, bonusAmountCol).setValue(bonusAmount);
    }
  });
}

// --- Offer Letter Terms Mapping ---
function getBonusRate(workerName: string): number {
  if (!workerName) return 80;

  // Example logic — adjust to your real Offer Letter Terms
  if (workerName.includes("(RN Supervisor)")) return 85;

  if (
    workerName.includes("(RN)") ||
    workerName.includes("(PT)") ||
    workerName.includes("(OT)") ||
    workerName.includes("(PTA)") ||
    workerName.includes("(LPN)")
  ) {
    return 80;
  }

  // Default fallback
  return 80;
}