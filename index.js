const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const onedrive = require('onedrive-api');
const { ConfidentialClientApplication } = require("@azure/msal-node");
const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch"); // required for Graph client in Node
require('dotenv').config();
const axios = require("axios");
const qs = require("qs");
require('dotenv').config();

const tenantId = process.env.TENANTID;
const clientId = process.env.CLIENTID;
const clientSecret = process.env.CLIENTSECRETVALUE;

async function getToken() {
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    try {
        return new Promise(async (resolve, reject) => {
            const response = await axios.post(tokenUrl, qs.stringify({
                client_id: clientId,
                scope: "https://graph.microsoft.com/.default",
                client_secret: clientSecret,
                grant_type: "client_credentials",
            }),
            { 
                headers: { "Content-Type": "application/x-www-form-urlencoded" 

            }});
            if(response.data && response.data.access_token) {
                resolve(response.data.access_token);
            }
            else {
                reject(new Error('Failed to acquire access token. No access_token in response.'));
            }
        });
    }
    catch(error) {
        console.error('Error acquiring access token:', error.response ? error.response.data : error.message || error);
    }
}
async function getUsers() {
    try {
        return new Promise(async (resolve, reject) => {
            const token = await getToken();
        const res = await axios.get("https://graph.microsoft.com/v1.0/users", {
            headers: {
                Authorization: `Bearer ${token}`
            }
        });
    
        if(res.status !== 200 || !res.data || !res.data.value) {
            reject(new Error(`Failed to fetch users. Status: ${res.status}, Response: ${JSON.stringify(res.data)}`));
        }
        else {
            console.log(res.data.value.map(u => ({
                    name: u.displayName,
                    upn: u.userPrincipalName,
                    id: u.id
                })));
                resolve(res.data.value);
            }
        });
    }
    catch(error) {
        console.error('Error fetching users from Graph API:', error.response ? error.response.data : error.message || error);
    }
}
async function listFiles() {
    const token = await getToken();
    try {
        return new Promise(async(resolve, reject) => {
            const res = await axios.get(`https://graph.microsoft.com/v1.0/users/${process.env.UPN}/drive/root/children`, {
            headers: {
                Authorization: `Bearer ${token}`
            }
        });
        if(!res.data || !res.data.value) {
            reject(new Error('No files data returned from Graph API.'));
        }
        else {
            console.log("Files:", res.data.value.map(f => f.name));
            resolve(res.data.value);
        }
      }
    );
  } catch (err) {
    console.error("REAL ERROR:", err.response?.data || err.message);
  }
}
//getUsers();
//listFiles();

// Alternative using MSAL and Graph client
// const cca = new ConfidentialClientApplication({
//   auth: {
//     clientId: process.env.CLIENTID,
//     authority: `https://login.microsoftonline.com/${process.env.TENANTID}`,
//     clientSecret: process.env.CLIENTSECRETVALUE
//   }
// });
// async function getAccessToken() {
//     try {
//         return new Promise(async (resolve,reject)=> {
//             const result = await cca.acquireTokenByClientCredential({
//                 scopes: ["https://graph.microsoft.com/.default"]
//             });
//             if(result && result.accessToken) {
//                 resolve(result.accessToken);
//             }            else {
//                 reject(new Error('Failed to acquire access token.'));
//             }
//         });
//     }
//     catch(e) {
//         console.error('Error acquiring access token:', e);
//     }
// }
// async function listOneDriveFiles(userUpn, folderPath) {
//     try {
//         return new Promise(async (resolve, reject) => {
//             const accessToken = await getAccessToken();
//             const graphClient = Client.init({authProvider: (done) => done(null, accessToken)});
//             let endpoint = folderPath ? `/users/${userUpn}/drive/root:/${encodeURI(folderPath)}:/children` : `/users/${userUpn}/drive/root/children`;
//             console.log("Graph endpoint:", endpoint);
//             const response = await graphClient.api(endpoint).get();
//             if (!response.value || response.value.length === 0) {
//                 console.warn("No files returned. Check folder path or permissions.");
//             }
//             console.log("Files:", response.value.map(f => f.name));
//             resolve(response.value);
//         });

//     } catch (error) {
//         console.error("Error listing OneDrive files:", error.message || error);
//         throw error;
//     }
// }
// async function oneDriveDownloadFile(userUpn, pathToFile) {
//     try {
//         const accessToken = await getAccessToken();
//         const graphClient = Client.init({authProvider: done => done(null, accessToken)});
//         return new Promise(async (resolve, reject) => {
//             // Normalize path (no root:, no trailing :)
//             const normalizedPath = pathToFile
//                 .replace(/^root:\//, "")
//                 .replace(/^\/+/, "")
//                 .replace(/:+$/, "");
//             // Graph download endpoint
//             const downloadUrl = `/users/${userUpn}/drive/root:/${normalizedPath}:/content`;
//             const stream = await graphClient.api(downloadUrl).getStream();
//             if(stream && typeof stream.pipe === 'function') {
//                 resolve(stream); // Node.js readable stream
//             }
//             else {
//                 reject(new Error('Failed to download file stream from OneDrive.'));
//             }
//         });
//     }
//     catch(error) {

//     }
// }
// async function findAndProcessFiles() {
//     try {
//         const files = await listOneDriveFiles(process.env.UPN,"Documents/Automation/PointsReport");
//         const regex = /points[-_\s]+report/i;
//         const matchingFiles = files.filter(file => regex.test(file.name));
//         if (matchingFiles.length === 0) {
//             console.log('No matching files found in OneDrive.');
//             return;
//         }
//         else {
//             console.log('Matching files in OneDrive:', matchingFiles.map(f => f.name));
//         }
//     }
//     catch(error) {
//         console.error('Error finding files in OneDrive:', error);
//     }
// }
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
//main();