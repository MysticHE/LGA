/**
 * Shared Excel Graph API Operations
 * Centralized Excel update functions to prevent code duplication and enable queuing
 */

/**
 * Get Excel column letter from index (A, B, C, ... Z, AA, AB, etc.)
 * @param {number} colIndex - 0-based column index
 * @returns {string} Excel column letter(s)
 */
function getExcelColumnLetter(colIndex) {
    let result = '';
    while (colIndex >= 0) {
        result = String.fromCharCode((colIndex % 26) + 65) + result;
        colIndex = Math.floor(colIndex / 26) - 1;
    }
    return result;
}

/**
 * Update lead data in Excel using Microsoft Graph API
 * @param {object} graphClient - Microsoft Graph client
 * @param {string} email - Email address to find and update
 * @param {object} updates - Fields to update
 * @returns {Promise<boolean>} Success status
 */
async function updateLeadViaGraphAPI(graphClient, email, updates) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        console.log(`🔍 Graph API: Searching for ${email} in Excel file...`);
        
        // Get the Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            console.log(`❌ Master file not found: ${masterFileName}`);
            return false;
        }

        const fileId = files.value[0].id;
        
        // Get worksheets to find the correct sheet name
        const worksheets = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets`)
            .get();
            
        const leadsSheet = worksheets.value.find(sheet => 
            sheet.name === 'Leads' || sheet.name.toLowerCase().includes('lead')
        ) || worksheets.value[0];
        
        if (!leadsSheet) {
            console.log(`❌ No leads worksheet found in ${masterFileName}`);
            return false;
        }
        
        // Get worksheet data to find the email
        const usedRange = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets('${leadsSheet.name}')/usedRange`)
            .get();
        
        if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
            console.log(`❌ No data found in worksheet ${leadsSheet.name}`);
            return false;
        }
        
        const headers = usedRange.values[0];
        const rows = usedRange.values.slice(1);
        
        // Debug: Log all headers found
        console.log(`📋 Excel headers found: ${headers.join(', ')}`);
        
        // Find email column and target row
        const emailColumnIndex = headers.findIndex(header => 
            header && typeof header === 'string' && 
            header.toLowerCase().includes('email') && 
            !header.toLowerCase().includes('date')
        );
        
        if (emailColumnIndex === -1) {
            console.log(`❌ Email column not found in ${leadsSheet.name}`);
            return false;
        }
        
        let targetRowIndex = -1;
        for (let i = 0; i < rows.length; i++) {
            const rowEmail = rows[i][emailColumnIndex];
            if (rowEmail && rowEmail.toLowerCase().trim() === email.toLowerCase().trim()) {
                targetRowIndex = i + 2; // +2 for 1-based and header row
                console.log(`📍 Found lead at row ${targetRowIndex}`);
                break;
            }
        }
        
        if (targetRowIndex === -1) {
            console.log(`❌ Lead with email ${email} not found in Excel file`);
            return false;
        }
        
        // Helper function for flexible header matching
        function findHeaderIndex(headers, fieldName) {
            // Normalize field name for comparison
            const normalizeString = (str) => str.toString().toLowerCase().trim().replace(/[\s_-]+/g, '');
            const normalizedField = normalizeString(fieldName);
            
            // Try exact match first
            let colIndex = headers.findIndex(h => h === fieldName);
            if (colIndex !== -1) return colIndex;
            
            // Try case-insensitive exact match
            colIndex = headers.findIndex(h => 
                h && h.toString().toLowerCase().trim() === fieldName.toLowerCase().trim()
            );
            if (colIndex !== -1) return colIndex;
            
            // Try normalized comparison (ignore spaces, underscores, hyphens)
            colIndex = headers.findIndex(h => 
                h && normalizeString(h) === normalizedField
            );
            if (colIndex !== -1) return colIndex;
            
            // For Email_Count, also try variations like "Email Count", "EmailCount", etc.
            if (fieldName === 'Email_Count') {
                const variations = ['Email Count', 'EmailCount', 'Email-Count', 'email_count', 'email count', 'emailcount'];
                for (const variation of variations) {
                    colIndex = headers.findIndex(h => 
                        h && normalizeString(h) === normalizeString(variation)
                    );
                    if (colIndex !== -1) return colIndex;
                }
            }
            
            return -1;
        }
        
        // Update each field
        console.log(`🔄 Updating ${Object.keys(updates).length} fields: ${Object.keys(updates).join(', ')}`);
        for (const [field, value] of Object.entries(updates)) {
            const colIndex = findHeaderIndex(headers, field);
            if (colIndex !== -1) {
                const cellAddress = `${getExcelColumnLetter(colIndex)}${targetRowIndex}`;
                
                try {
                    await graphClient
                        .api(`/me/drive/items/${fileId}/workbook/worksheets('${leadsSheet.name}')/range(address='${cellAddress}')`)
                        .patch({
                            values: [[value]]
                        });
                    
                    console.log(`📝 Updated ${field} = ${value} at ${cellAddress} (header: ${headers[colIndex]})`);
                } catch (cellError) {
                    console.error(`❌ Failed to update ${field} at ${cellAddress}:`, cellError.message);
                }
            } else {
                console.warn(`⚠️ Header not found for field: ${field}. Available headers: ${headers.join(', ')}`);
            }
        }
        
        return true;
        
    } catch (error) {
        console.error('❌ Update lead via Graph API error:', error.message);
        return false;
    }
}

/**
 * Get all leads from Excel using Microsoft Graph API
 * @param {object} graphClient - Microsoft Graph client
 * @returns {Promise<Array|null>} Array of lead objects or null if failed
 */
async function getLeadsViaGraphAPI(graphClient) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Get the Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            return null;
        }

        const fileId = files.value[0].id;
        
        // Get worksheets to find the correct sheet name
        const worksheets = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets`)
            .get();
            
        const leadsSheet = worksheets.value.find(sheet => 
            sheet.name === 'Leads' || sheet.name.toLowerCase().includes('lead')
        ) || worksheets.value[0];
        
        if (!leadsSheet) {
            return null;
        }
        
        // Get worksheet data
        const usedRange = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets('${leadsSheet.name}')/usedRange`)
            .get();
        
        if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
            return [];
        }
        
        // Convert to lead objects
        const headers = usedRange.values[0];
        const rows = usedRange.values.slice(1);
        
        return rows.map(row => {
            const lead = {};
            headers.forEach((header, index) => {
                lead[header] = row[index] || '';
            });
            return lead;
        }).filter(lead => lead.Email);
        
    } catch (error) {
        console.error('❌ Get leads via Graph API error:', error.message);
        return null;
    }
}

module.exports = {
    updateLeadViaGraphAPI,
    getLeadsViaGraphAPI,
    getExcelColumnLetter
};