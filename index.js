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

async function syncAssessments() {
    console.log("Starting sync...");
    try {
        // 1. Get Group Drive
        const drives = await client.api(`/groups/${GROUP_ID}/drives`).get();
        if (drives.value.length === 0) throw new Error("No drive found for group.");
        const driveId = drives.value[0].id;

        // 2. Get files in Group Drive
        const filesRes = await client.api(`/drives/${driveId}/root/children`).get();
        const formFiles = filesRes.value.filter(f => f.name.includes('NSL Assessment Cent') && f.name.endsWith('.xlsx'));
        console.log(`Found ${formFiles.length} Assessment files.`);

        let allWorkbooksData = [];
        let columnCounts = {};

        // 3. Download and parse each file
        for (const file of formFiles) {
            console.log(`Downloading ${file.name}...`);
            // Get download URL
            const fileItem = await client.api(`/drives/${driveId}/items/${file.id}`).get();
            const downloadUrl = fileItem['@microsoft.graph.downloadUrl'];
            
            const response = await fetch(downloadUrl);
            const arrayBuffer = await response.arrayBuffer();
            const buffer = Buffer.from(arrayBuffer);
            
            const wb = XLSX.read(buffer, { type: 'buffer' });
            const sheetName = wb.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
            
            // Extract profession from file name
            let profession = file.name.replace('NSL Assessment Centre', '').replace('NSL Assessment Center', '').replace('.xlsx', '').trim();
            if (profession.startsWith('-')) profession = profession.substring(1).trim();
            if (!profession) profession = 'Khác';

            // Count columns for Lite version logic
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

        // 4. Determine allowed columns for Lite version
        const MIN_FILE_COUNT = Math.max(1, formFiles.length - 1); // e.g. 11 out of 12
        let allowedColumns = Object.keys(columnCounts).filter(h => {
            if (h.startsWith('Points') || h.startsWith('Feedback') || h.includes('Quiz feedback') || h === 'Total points' || h === 'Grade posted time') return false;
            return columnCounts[h] >= MIN_FILE_COUNT || h.toLowerCase().includes('điện thoại');
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
        XLSX.utils.book_append_sheet(newWb, newWs, 'Master Responses Lite');
        const outBuffer = XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' });

        // 7. Upload to User's OneDrive (MANAGEMENT/Assessment)
        console.log("Uploading Master_Assessment_Responses_Lite.xlsx to OneDrive...");
        await client.api(`/users/${USER_EMAIL}/drive/root:/MANAGEMENT/Assessment/Master_Assessment_Responses_Lite.xlsx:/content`)
            .put(outBuffer);
            
        console.log("Sync completed successfully!");

    } catch (error) {
        console.error("Sync failed:", error);
        process.exit(1);
    }
}

syncAssessments();
