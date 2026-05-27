const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const XLSX = require('xlsx');

const CLIENT_ID = '747489e9-bf4b-405c-9d4d-bc483eebfd2f';
const TENANT_ID = 'b4295bc1-8dd4-40af-935d-820db1079364';
const CLIENT_SECRET = process.env.CLIENT_SECRET; // from Github Actions secret
const USER_EMAIL = 'khoi.nguyen@new-solution.eu';
const GROUP_ID = '8ff3b231-2dcb-42d0-9ebd-146db72d8e07'; // NSL | Assessment Forms

if (!CLIENT_SECRET) {
    console.error("Missing CLIENT_SECRET environment variable.");
    process.exit(1);
}

const msalConfig = {
    auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        clientSecret: CLIENT_SECRET,
    }
};

const cca = new ConfidentialClientApplication(msalConfig);
const authProvider = {
    getAccessToken: async () => {
        const authResponse = await cca.acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default'],
        });
        return authResponse.accessToken;
    }
};
const client = Client.initWithMiddleware({ authProvider });

async function getAllFiles(driveId, folderId = 'root', allFiles = []) {
    try {
        const res = await client.api(`/drives/${driveId}/items/${folderId}/children`).get();
        for (const item of res.value) {
            if (item.folder) {
                await getAllFiles(driveId, item.id, allFiles);
            } else {
                allFiles.push(item);
            }
        }
    } catch (e) {
        console.error("Error reading folder:", e.message);
    }
    return allFiles;
}

async function syncAssessments() {
    console.log("Starting sync...");
    try {
        // 1. Get Group Drive
        const drives = await client.api(`/groups/${GROUP_ID}/drives`).get();
        if (drives.value.length === 0) throw new Error("No drive found for group.");
        const driveId = drives.value[0].id;

        // 2. Get all files recursively in Group Drive
        console.log("Scanning drive for Form files...");
        const allItems = await getAllFiles(driveId);
        const formFiles = allItems.filter(f => f.name.includes('NSL Assessment Cent') && f.name.endsWith('.xlsx'));
        console.log(`Found ${formFiles.length} Assessment files.`);

        let allWorkbooksData = [];
        let columnCounts = {};

        // 3. Download and parse each file
        for (const file of formFiles) {
            console.log(`Downloading ${file.name}...`);
            const fileItem = await client.api(`/drives/${driveId}/items/${file.id}`).get();
            const downloadUrl = fileItem['@microsoft.graph.downloadUrl'];
            
            const response = await fetch(downloadUrl);
            const arrayBuffer = await response.arrayBuffer();
            const buffer = Buffer.from(arrayBuffer);
            
            const wb = XLSX.read(buffer, { type: 'buffer' });
            const sheetName = wb.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
            
            let profession = file.name.replace('NSL Assessment Centre', '').replace('NSL Assessment Center', '').replace('.xlsx', '').trim();
            if (profession.startsWith('-')) profession = profession.substring(1).trim();
            if (!profession) profession = 'Khác';

            if (rows.length > 0) {
                const headerRow = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 })[0];
                if (headerRow) {
                    const headers = headerRow.map(h => typeof h === 'string' ? h.replace(/\r\n/g, '').trim() : h);
                    headers.forEach(h => {
                        if (h) columnCounts[h] = (columnCounts[h] || 0) + 1;
                    });
                }
            }
            allWorkbooksData.push({ profession, rows });
        }

        if (formFiles.length === 0) {
            console.log("No files to sync.");
            process.exit(0);
        }

        // 4. Determine allowed columns based on exact user request
        let allowedColumns = Object.keys(columnCounts).filter(h => {
            const lowerH = h.toLowerCase().trim();
            if (lowerH.startsWith('points -') || lowerH.startsWith('feedback -')) return false;
            
            const exactMatches = ['nghề nghiệp', 'id', 'start time', 'completion time', 'total points'];
            if (exactMatches.includes(lowerH)) return true;

            const partialMatches = [
                'tên đầy đủ',
                'mã số học viên',
                'ngày sinh',
                'địa chỉ email chính',
                'công việc chính mà bạn quan tâm',
                'số điện thoại'
            ];
            return partialMatches.some(pattern => lowerH.includes(pattern));
        });

        // 5. Build Master Data
        let masterData = [];
        for (const wbData of allWorkbooksData) {
            for (const row of wbData.rows) {
                let newRow = { 'Nghề nghiệp': wbData.profession };
                for (const key of Object.keys(row)) {
                    const cleanKey = key.replace(/\r\n/g, '').trim();
                    if (allowedColumns.includes(cleanKey)) {
                        newRow[cleanKey] = row[key];
                    }
                }
                masterData.push(newRow);
            }
        }

        console.log(`Total rows combined: ${masterData.length}`);

        // 6. Create Excel Buffer
        const newWb = XLSX.utils.book_new();
        const newWs = XLSX.utils.json_to_sheet(masterData);
        XLSX.utils.book_append_sheet(newWb, newWs, 'Master Responses');
        const outBuffer = XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' });

        // 7. Upload to User's OneDrive (MANAGEMENT/Assessment)
        console.log("Uploading Master_Assessment_Responses.xlsx to OneDrive...");
        await client.api(`/users/${USER_EMAIL}/drive/root:/MANAGEMENT/Assessment/Master_Assessment_Responses.xlsx:/content`)
            .put(outBuffer);
            
        console.log("Sync completed successfully!");

    } catch (error) {
        console.error("Sync failed:", error);
        process.exit(1);
    }
}

syncAssessments();
