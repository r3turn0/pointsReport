const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function readDirectory(filePath) {
    try {
        return new Promise((resolve, reject) => {
            fs.readdir('.', (err, files) => {
                if (err) {
                    console.error('Error reading directory:', err);
                    reject(err);
                    return;
                }
                const regex = /points[-_\s]+report/i;
                const matchingFiles = files.filter(file => regex.test(file));
                if (matchingFiles.length === 0) {
                    console.log('No matching files found.');
                    reject(new Error('No matching files found.'));
                } else {
                    console.log('Matching files:', matchingFiles);
                    resolve(matchingFiles);
                }   
            }); 
        });
    }
    catch (error) {
        console.error('Error reading directory:', error);
    }
}
async function getColumnIndex(sheet, columnName) {
    const headerRow = sheet.getRow(1);
    for (let col = 1; col <= headerRow.cellCount; col++) {
    const value = headerRow.getCell(col).value;
        if (typeof value === 'string' && value.trim().toLowerCase() === columnName.toLowerCase()) {
            return col;
        }
    }
  throw new Error(`Column not found: "${columnName}"`);
}
// Function to add "File Calculated Miles" column and populate it with N2 + R2 values
async function addN2R2Column(fileName, worksheetName) {
    try {
        const workbook = new ExcelJS.Workbook();
        const ext = path.extname(fileName).toLowerCase();
        if (ext !== '.xlsx') {
            console.error('Unsupported file format:', ext);
            return;
        }
        else {
            await workbook.xlsx.readFile(fileName);
            console.log('File read successfully:', fileName, 'looking for worksheet:', worksheetName);
            const regex = new RegExp(worksheetName, 'i');
            // const worksheet = workbook.worksheets.forEach(ws => {
            //     console.log('Worksheet name:', ws.name)
            //     if(ws.name && regex.test(ws.name)) {
            //         console.log('Found matching worksheet:', ws.name);
            //         return ws;
            //     }
            //     else {
            //         //console.log('No matching worksheet found in file:', fileName);
            //     }
            // });
            const worksheet = workbook.worksheets.find(ws => ws.name && regex.test(ws.name));
            if (!worksheet) {
                console.error('Worksheet not found:', worksheetName, 'in file:', fileName);
                return;
            }
            // Identify column indexes
            console.log('', 'Identifying column indexes for N2, R2, and Adjustment in worksheet:', worksheet.name);
            const headers = worksheet.getRow(1).values;
            const n2Idx = await getColumnIndex(worksheet, 'Calculated Miles');
            const r2Idx = await getColumnIndex(worksheet, 'Adjustment');
            const adjustmentIdx = await getColumnIndex(worksheet, 'Adjustment');
            const finalMilesIdx = await getColumnIndex(worksheet, 'Final Calculated Miles').catch(() => null);
            if(!finalMilesIdx) {                 
                console.log('Final Calculated Miles column not found, will be added after Adjustment column.');
                console.log({ n2Idx, r2Idx, adjustmentIdx });
                [n2Idx, r2Idx, adjustmentIdx].forEach(idx => {
                if (!Number.isInteger(idx) || idx < 1) {
                    throw new Error(`Invalid column index detected: ${idx}`);
                }
                });
                // Insert column after Adjustment
                worksheet.spliceColumns(adjustmentIdx + 1, 0, ['Final Calculated Miles']);
                // Recalculate indexes after insert
                const newN2Col = n2Idx >= adjustmentIdx ? n2Idx + 1 : n2Idx;
                const newR2Col = r2Idx >= adjustmentIdx ? r2Idx : n2Idx + 4;
                console.log('New column indices:', { newN2Col, newR2Col });
                worksheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1) return;
                    const n2 = Number(row.getCell(newN2Col).value || 0);
                    const r2 = Number(row.getCell(newR2Col).value || 0);
                    row.getCell(adjustmentIdx + 1).value = Number((n2 + r2).toFixed(2)) > 0 ? Number((n2 + r2).toFixed(2)) : 0.00;
                });
            }
            else {
                console.log('Final Calculated Miles column found at index:', finalMilesIdx);
                worksheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1) return;
                    const n2 = Number(row.getCell(n2Idx).value || 0);
                    const r2 = Number(row.getCell(r2Idx).value || 0);
                    row.getCell(finalMilesIdx).value = Number((n2 + r2).toFixed(2)) > 0 ? Number((n2 + r2).toFixed(2)) : 0.00;
                });
            }
            await workbook.xlsx.writeFile(fileName);
            console.log('File updated successfully:', fileName);
        }
    }
    catch(error) {
        console.error('Error processing file:', error);
    }
}
// Main function to read directory and process each file
async function main() {
    try {
        const files = await readDirectory();
        for (const file of files) {
            await addN2R2Column(file, 'Detailed Mileage Analysis');
        }
    }
    catch (error) {
        console.error('Error in main function:', error);
    }
}
// Run the main function
main();