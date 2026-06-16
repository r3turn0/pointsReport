function main(workbook: ExcelScript.Workbook) {

    const offerSheet = workbook.getWorksheet("Offer Letter Terms");
    const offerData = offerSheet.getUsedRange().getValues();

    const headers = offerData[0].map(h => String(h).trim());

    const workerNameCol = headers.indexOf("Worker Name");
    const bonusRateCol = 5; // Column F
    const thresholdCol = 6; // Column G
    const visitServiceCol = headers.indexOf("Service");        // adjust if needed
    const visitPointsCol = headers.indexOf("Visit Points");    // adjust if needed

    function cleanName(name: string): string {
        return name.replace(/\s*\(.*?\)/g, "").trim().toUpperCase();
    }

    // ✅ Build worker lookup
    const workerMap: {
        [key: string]: { rate: number, threshold: number }
    } = {};

    // ✅ Build service → visit points lookup
    const serviceMap: { [key: string]: number } = {};

    for (let i = 1; i < offerData.length; i++) {

        const rawName = String(offerData[i][workerNameCol]);
        const name = cleanName(rawName);

        const rate = Number(offerData[i][bonusRateCol]) || 0;
        const threshold = Number(offerData[i][thresholdCol]) || 80;

        if (name) {
            workerMap[name] = { rate, threshold };
        }

        // ✅ Service mapping
        const service = String(offerData[i][visitServiceCol] || "")
            .trim()
            .toUpperCase();

        const points = Number(offerData[i][visitPointsCol]) || 0;

        if (service && points) {
            serviceMap[service] = points;
        }
    }

    // ✅ Loop worker sheets
    workbook.getWorksheets().forEach(sheet => {

        if (sheet.getName() === "Offer Letter Terms") return;

        const range = sheet.getUsedRange();
        if (!range) return;

        const values = range.getValues();
        if (values.length < 2) return;

        const headers = values[0].map(h => String(h).trim());

        const workerCol = headers.indexOf("Worker Name");
        const serviceCol = headers.indexOf("Service");
        const pointsCol = headers.indexOf("Points");

        if (workerCol === -1 || serviceCol === -1 || pointsCol === -1) {
            return;
        }

        // ✅ One worker per sheet
        const workerName = cleanName(String(values[1][workerCol]));

        const workerData = workerMap[workerName];

        if (!workerData) {
            sheet.getRange("A1").setValue(`❌ Worker not found`);
            return;
        }

        const { rate, threshold } = workerData;

        // ✅ FIX A2 (this WILL update)
        sheet.getRange("A2").setValue(
            `Bonus per visit for anything over ${threshold} points.`
        );

        let totalPoints = 0;

        // ✅ Update Points column based on service
        for (let i = 1; i < values.length; i++) {

            const rawService = String(values[i][serviceCol] || "");

            // Extract service type after dash
            const serviceType = rawService.split("-").pop()?.trim().toUpperCase();

            if (!serviceType) continue;

            const newPoints = serviceMap[serviceType];

            if (newPoints !== undefined) {
                sheet.getCell(i, pointsCol).setValue(newPoints);
                totalPoints += newPoints;
            }
        }

        // ✅ Calculate Bonus
        const bonusPoints = Math.max(0, totalPoints - threshold);
        const bonusAmount = bonusPoints * rate;

        // ✅ Write results (top of sheet)
        sheet.getRange("B1").setValue(`Total Points: ${totalPoints}`);
        sheet.getRange("B2").setValue(`Bonus Points: ${bonusPoints}`);
        sheet.getRange("B3").setValue(`Bonus Amount: $${bonusAmount}`);

        // ✅ Optional: write to columns if they exist
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

        for (let i = 1; i < values.length; i++) {
            sheet.getCell(i, bonusPointsCol).setValue(bonusPoints);
            sheet.getCell(i, bonusAmountCol).setValue(bonusAmount);
        }
    });
}
