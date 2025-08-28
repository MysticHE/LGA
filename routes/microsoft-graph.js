const express = require('express');
const XLSX = require('xlsx');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const router = express.Router();

// Apply delegated Graph authentication middleware to protected routes
// Note: /test route is excluded from authentication requirement

/**
 * OneDrive Excel Integration with Table Append Functionality
 * Manages lead data in Microsoft 365 Excel files using table-based operations
 * - Automatically creates tables if they don't exist
 * - Appends data to existing tables without overwriting
 * - Uses Microsoft Graph Excel API for table operations
 */

// Configuration for Excel table operations
const EXCEL_CONFIG = {
    MASTER_FILE_PATH: '/LGA-Leads/LGA-Master-Email-List.xlsx',
    WORKSHEET_NAME: 'Leads',
    TABLE_NAME: 'LeadsTable',
    COLUMN_MAPPING: {
        'Name': 'name',
        'Title': 'title', 
        'Company Name': 'organization_name',
        'Company Website': 'organization_website_url',
        'Size': 'estimated_num_employees',
        'Email': 'email',
        'Email Verified': 'email_verified',
        'LinkedIn URL': 'linkedin_url',
        'Industry': 'industry',
        'Location': 'country',
        'Notes': 'notes',
        'Conversion Status': 'conversion_status'
    }
};

// Append leads to Excel table in OneDrive (replaces create-excel)
router.post('/onedrive/append-to-table', requireDelegatedAuth, async (req, res) => {
    try {
        const { leads, filename, folderPath = '/LGA-Leads', useCustomFile = false } = req.body;

        if (!leads || !Array.isArray(leads) || leads.length === 0) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Leads array is required and must not be empty'
            });
        }

        console.log(`📊 Appending ${leads.length} leads to Excel table in OneDrive...`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Determine target file path
        let targetFilePath;
        if (useCustomFile && filename) {
            const cleanFolderPath = folderPath.startsWith('/') ? folderPath.substring(1) : folderPath;
            targetFilePath = cleanFolderPath ? `${cleanFolderPath}/${filename}` : filename;
        } else {
            targetFilePath = EXCEL_CONFIG.MASTER_FILE_PATH.substring(1); // Remove leading slash
        }

        console.log(`📁 Target file: ${targetFilePath}`);

        // Check if file exists
        const fileInfo = await getOneDriveFileInfo(graphClient, targetFilePath);
        let fileId;
        
        if (!fileInfo) {
            // File doesn't exist, create it with initial table
            console.log(`🆕 Creating new Excel file with table: ${targetFilePath}`);
            fileId = await createExcelFileWithTable(graphClient, targetFilePath, leads);
            
            res.json({
                success: true,
                action: 'created',
                filename: targetFilePath.split('/').pop(),
                folderPath: '/' + targetFilePath.substring(0, targetFilePath.lastIndexOf('/')),
                leadsCount: leads.length,
                fileId: fileId,
                tableCreated: true,
                metadata: {
                    uploadedAt: new Date().toISOString(),
                    location: 'Microsoft OneDrive'
                }
            });
            return;
        }
        
        fileId = fileInfo.id;
        console.log(`✅ Found existing file with ID: ${fileId}`);

        // Check if table exists in the worksheet
        const tableInfo = await getExcelTableInfo(graphClient, fileId, EXCEL_CONFIG.WORKSHEET_NAME, EXCEL_CONFIG.TABLE_NAME);
        
        if (!tableInfo) {
            // Table doesn't exist, create it
            console.log(`🆕 Creating table '${EXCEL_CONFIG.TABLE_NAME}' in worksheet '${EXCEL_CONFIG.WORKSHEET_NAME}'`);
            await createExcelTable(graphClient, fileId, EXCEL_CONFIG.WORKSHEET_NAME, EXCEL_CONFIG.TABLE_NAME, leads);
        } else {
            // Table exists, append data
            console.log(`➕ Appending data to existing table '${EXCEL_CONFIG.TABLE_NAME}'`);
            await appendDataToExcelTable(graphClient, fileId, EXCEL_CONFIG.TABLE_NAME, leads);
        }

        res.json({
            success: true,
            action: 'appended',
            filename: targetFilePath.split('/').pop(),
            folderPath: '/' + targetFilePath.substring(0, targetFilePath.lastIndexOf('/')),
            leadsCount: leads.length,
            fileId: fileId,
            tableExists: !!tableInfo,
            metadata: {
                updatedAt: new Date().toISOString(),
                location: 'Microsoft OneDrive'
            }
        });

    } catch (error) {
        console.error('OneDrive Excel append error:', error);
        res.status(500).json({
            error: 'OneDrive Excel Append Error',
            message: 'Failed to append data to Excel table in OneDrive',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Legacy endpoint - redirects to new append functionality
router.post('/onedrive/create-excel', requireDelegatedAuth, async (req, res) => {
    console.log('⚠️ Using legacy create-excel endpoint, redirecting to append-to-table...');
    req.body.useCustomFile = true; // Allow custom filename
    return require('./microsoft-graph').post('/onedrive/append-to-table')(req, res);
});

// Update Excel table with email tracking data using Graph API
router.post('/onedrive/update-excel-tracking', requireDelegatedAuth, async (req, res) => {
    try {
        const { fileId, leadEmail, trackingData } = req.body;

        if (!fileId || !leadEmail || !trackingData) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'File ID, lead email, and tracking data are required'
            });
        }

        console.log(`📊 Updating Excel table tracking for ${leadEmail}...`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Update tracking data using table operations
        const success = await updateLeadTrackingInTable(
            graphClient, 
            fileId, 
            EXCEL_CONFIG.TABLE_NAME, 
            leadEmail, 
            trackingData
        );

        if (!success) {
            return res.status(404).json({
                error: 'Lead Not Found',
                message: 'Lead email not found in Excel table'
            });
        }

        console.log(`✅ Excel table tracking updated for ${leadEmail}`);

        res.json({
            success: true,
            message: `Tracking data updated for ${leadEmail}`,
            updatedAt: new Date().toISOString()
        });

    } catch (error) {
        console.error('Excel table tracking update error:', error);
        res.status(500).json({
            error: 'Excel Table Update Error',
            message: 'Failed to update Excel table tracking data',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// List OneDrive files
router.get('/onedrive/files', requireDelegatedAuth, async (req, res) => {
    try {
        const { folderPath = '/LGA-Leads' } = req.query;

        console.log(`📂 Listing OneDrive files in ${folderPath}...`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Handle path correctly for OneDrive API
        const cleanPath = folderPath.startsWith('/') ? folderPath.substring(1) : folderPath;
        
        const files = await graphClient
            .api(`/me/drive/root:/${cleanPath}:/children`)
            .filter("file ne null")
            .select('id,name,size,createdDateTime,lastModifiedDateTime,webUrl')
            .get();

        res.json({
            success: true,
            folderPath: folderPath,
            count: files.value.length,
            files: files.value.map(file => ({
                id: file.id,
                name: file.name,
                size: file.size,
                createdAt: file.createdDateTime,
                modifiedAt: file.lastModifiedDateTime,
                webUrl: file.webUrl
            }))
        });

    } catch (error) {
        console.error('OneDrive file listing error:', error);
        res.status(500).json({
            error: 'OneDrive Error',
            message: 'Failed to list OneDrive files',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Helper function to create OneDrive folder
async function createOneDriveFolder(client, folderPath) {
    try {
        // Handle root path correctly - remove leading slash for OneDrive API
        const cleanPath = folderPath.startsWith('/') ? folderPath.substring(1) : folderPath;
        
        if (!cleanPath) {
            console.log(`📂 Root folder - no creation needed`);
            return;
        }
        
        // Check if folder exists
        await client.api(`/me/drive/root:/${cleanPath}`).get();
        console.log(`📂 Folder ${folderPath} already exists`);
    } catch (error) {
        if (error.code === 'itemNotFound') {
            const folderName = folderPath.split('/').pop();
            const parentPath = folderPath.substring(0, folderPath.lastIndexOf('/'));
            
            // Create folder in correct location
            if (parentPath && parentPath !== '/') {
                // Create in subfolder
                const cleanParentPath = parentPath.startsWith('/') ? parentPath.substring(1) : parentPath;
                await client.api(`/me/drive/root:/${cleanParentPath}:/children`).post({
                    name: folderName,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'rename'
                });
            } else {
                // Create in root
                await client.api(`/me/drive/root/children`).post({
                    name: folderName,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'rename'
                });
            }
            
            console.log(`📂 Created folder: ${folderPath}`);
        } else {
            throw error;
        }
    }
}

// Helper function to upload file to OneDrive with retry logic
async function uploadToOneDrive(client, fileBuffer, filename, folderPath, maxRetries = 3) {
    // Handle path correctly for OneDrive API
    let uploadUrl;
    const cleanPath = folderPath.startsWith('/') ? folderPath.substring(1) : folderPath;
    
    if (cleanPath) {
        // Upload to subfolder
        uploadUrl = `/me/drive/root:/${cleanPath}/${filename}:/content`;
    } else {
        // Upload to root
        uploadUrl = `/me/drive/root:/${filename}:/content`;
    }
    
    console.log(`📤 Uploading to URL: ${uploadUrl}`);
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const result = await client.api(uploadUrl).put(fileBuffer);
            
            console.log(`📤 Uploaded file: ${filename} to ${folderPath} (attempt ${attempt})`);
            
            return {
                id: result.id,
                name: result.name,
                webUrl: result.webUrl,
                size: result.size
            };
        } catch (error) {
            const isLockError = error.statusCode === 423 || 
                               error.code === 'resourceLocked' || 
                               error.code === 'notAllowed';
            
            if (isLockError && attempt < maxRetries) {
                const waitTime = Math.pow(2, attempt - 1) * 2000; // 2s, 4s, 8s
                console.log(`🔒 File is locked, waiting ${waitTime/1000}s before retry (attempt ${attempt}/${maxRetries})`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
                continue;
            }
            
            console.error('OneDrive upload error:', error);
            if (isLockError) {
                const customError = new Error(`File is locked. Please close the Excel file in OneDrive/Excel and try again. (Attempted ${maxRetries} times)`);
                customError.originalError = error;
                customError.isLockError = true;
                throw customError;
            }
            throw error;
        }
    }
}

// Test Excel table append functionality
router.post('/onedrive/test-table-append', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('🧪 Testing Excel table append functionality...');
        
        // Sample test data
        const testLeads = [
            {
                name: 'Test User 1',
                title: 'Test Manager',
                organization_name: 'Test Company 1',
                email: 'test1@example.com',
                industry: 'Testing',
                country: 'Test Country'
            },
            {
                name: 'Test User 2',
                title: 'Test Director',
                organization_name: 'Test Company 2',
                email: 'test2@example.com',
                industry: 'Testing',
                country: 'Test Country'
            }
        ];
        
        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        const targetFilePath = 'LGA-Leads/Test-Table-Append.xlsx';
        
        // Check if file exists
        const fileInfo = await getOneDriveFileInfo(graphClient, targetFilePath);
        let result;
        
        if (!fileInfo) {
            console.log('🆕 Creating new test file with table...');
            const fileId = await createExcelFileWithTable(graphClient, targetFilePath, testLeads);
            result = {
                success: true,
                action: 'created',
                fileId: fileId,
                filename: 'Test-Table-Append.xlsx',
                leadsCount: testLeads.length,
                message: 'New Excel file created with table'
            };
        } else {
            console.log('➕ Appending to existing test file...');
            await appendDataToExcelTable(graphClient, fileInfo.id, EXCEL_CONFIG.TABLE_NAME, testLeads);
            result = {
                success: true,
                action: 'appended',
                fileId: fileInfo.id,
                filename: 'Test-Table-Append.xlsx',
                leadsCount: testLeads.length,
                message: 'Data appended to existing table'
            };
        }
        
        console.log('✅ Test completed successfully!');
        res.json(result);
        
    } catch (error) {
        console.error('❌ Table append test failed:', error);
        res.status(500).json({
            success: false,
            error: 'Table Append Test Failed',
            message: error.message,
            details: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
    }
});

// Test Microsoft Graph connection - handles both authenticated and unauthenticated requests"}
router.get('/test', async (req, res) => {
    try {
        const sessionId = req.headers['x-session-id'] || req.query.sessionId;
        
        if (!sessionId) {
            return res.json({
                success: false,
                message: 'Authentication required for Microsoft Graph integration',
                authRequired: true,
                loginUrl: '/auth/login'
            });
        }

        const authProvider = getDelegatedAuthProvider();
        
        if (!authProvider.isAuthenticated(sessionId)) {
            return res.json({
                success: false,
                message: 'Session not authenticated',
                authRequired: true,
                loginUrl: '/auth/login'
            });
        }

        const testResult = await authProvider.testConnection(sessionId);
        
        if (testResult.success) {
            res.json({
                success: true,
                message: 'Microsoft Graph connection successful',
                user: testResult.user,
                email: testResult.email,
                accessType: 'Delegated',
                permissions: 'User.Read, Files.ReadWrite.All, Mail.Send, Mail.ReadWrite'
            });
        } else {
            res.status(401).json({
                success: false,
                message: 'Microsoft Graph connection failed',
                error: testResult.error,
                authRequired: true
            });
        }
    } catch (error) {
        console.error('Microsoft Graph test error:', error);
        res.status(500).json({
            success: false,
            message: 'Microsoft Graph test failed',
            error: error.message
        });
    }
});

// =============================================================================
// NEW HELPER FUNCTIONS FOR EXCEL TABLE OPERATIONS
// =============================================================================

/**
 * Get OneDrive file information
 * @param {Object} client - Microsoft Graph client
 * @param {string} filePath - OneDrive file path (without leading slash)
 * @returns {Object|null} File information or null if not found
 */
async function getOneDriveFileInfo(client, filePath) {
    try {
        const response = await client.api(`/me/drive/root:/${filePath}`).get();
        console.log(`✅ Found file: ${filePath} (ID: ${response.id})`);
        return {
            id: response.id,
            name: response.name,
            webUrl: response.webUrl,
            size: response.size
        };
    } catch (error) {
        if (error.code === 'itemNotFound') {
            console.log(`⚠️ File not found: ${filePath}`);
            return null;
        }
        console.error(`❌ Error checking file ${filePath}:`, error);
        throw error;
    }
}

/**
 * Get Excel table information from worksheet
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} worksheetName - Worksheet name
 * @param {string} tableName - Table name
 * @returns {Object|null} Table information or null if not found
 */
async function getExcelTableInfo(client, fileId, worksheetName, tableName) {
    try {
        const response = await client
            .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/tables/${tableName}`)
            .get();
        
        console.log(`✅ Found table '${tableName}' in worksheet '${worksheetName}'`);
        return {
            id: response.id,
            name: response.name,
            range: response.range
        };
    } catch (error) {
        if (error.code === 'itemNotFound' || error.code === 'InvalidArgument') {
            console.log(`⚠️ Table '${tableName}' not found in worksheet '${worksheetName}'`);
            return null;
        }
        console.error(`❌ Error checking table:`, error);
        throw error;
    }
}

/**
 * Create Excel file with initial table
 * @param {Object} client - Microsoft Graph client
 * @param {string} filePath - OneDrive file path
 * @param {Array} leads - Initial lead data
 * @returns {string} File ID
 */
async function createExcelFileWithTable(client, filePath, leads) {
    try {
        // Prepare initial data with all required columns
        const normalizedLeads = leads.slice(0, 5).map(lead => normalizeLeadData(lead));
        
        // Create Excel workbook using XLSX
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(normalizedLeads);
        
        // Set column widths
        ws['!cols'] = getColumnWidths();
        XLSX.utils.book_append_sheet(wb, ws, EXCEL_CONFIG.WORKSHEET_NAME);
        
        // Convert to buffer
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        // Create folder structure if needed
        const folderPath = filePath.substring(0, filePath.lastIndexOf('/'));
        if (folderPath) {
            await createOneDriveFolder(client, '/' + folderPath);
        }
        
        // Upload file to OneDrive
        const uploadResult = await client
            .api(`/me/drive/root:/${filePath}:/content`)
            .put(excelBuffer);
        
        const fileId = uploadResult.id;
        console.log(`✅ Created Excel file: ${filePath} (ID: ${fileId})`);
        
        // Create table in the uploaded file
        await createExcelTableInFile(client, fileId, EXCEL_CONFIG.WORKSHEET_NAME, EXCEL_CONFIG.TABLE_NAME, normalizedLeads);
        
        // If we have more than 5 leads, append the remaining ones
        if (leads.length > 5) {
            const remainingLeads = leads.slice(5).map(lead => normalizeLeadData(lead));
            await appendDataToExcelTable(client, fileId, EXCEL_CONFIG.TABLE_NAME, remainingLeads);
        }
        
        return fileId;
        
    } catch (error) {
        console.error(`❌ Error creating Excel file with table:`, error);
        throw error;
    }
}

/**
 * Create Excel table in existing file
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} worksheetName - Worksheet name
 * @param {string} tableName - Table name
 * @param {Array} initialData - Initial data for table
 */
async function createExcelTableInFile(client, fileId, worksheetName, tableName, initialData) {
    try {
        // Calculate table range based on data
        const headers = Object.keys(normalizeLeadData({})); // Get all column headers
        const numCols = headers.length;
        const numRows = initialData.length + 1; // +1 for header row
        
        // Convert column number to Excel column letter
        const endCol = getExcelColumnLetter(numCols);
        const tableRange = `A1:${endCol}${numRows}`;
        
        console.log(`🆕 Creating table '${tableName}' with range ${tableRange}`);
        
        // Create table using Graph API
        const tableRequest = {
            address: tableRange,
            hasHeaders: true,
            name: tableName
        };
        
        await client
            .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/tables/add`)
            .post(tableRequest);
        
        console.log(`✅ Created table '${tableName}' in worksheet '${worksheetName}'`);
        
    } catch (error) {
        console.error(`❌ Error creating table in file:`, error);
        throw error;
    }
}

/**
 * Create Excel table (wrapper for existing file)
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} worksheetName - Worksheet name
 * @param {string} tableName - Table name
 * @param {Array} leads - Lead data to determine structure
 */
async function createExcelTable(client, fileId, worksheetName, tableName, leads) {
    try {
        // Add some initial data to the worksheet first
        const normalizedLeads = leads.slice(0, 3).map(lead => normalizeLeadData(lead));
        
        // Clear and populate worksheet with initial data
        await populateWorksheetWithData(client, fileId, worksheetName, normalizedLeads);
        
        // Create table
        await createExcelTableInFile(client, fileId, worksheetName, tableName, normalizedLeads);
        
        // Append remaining data if any
        if (leads.length > 3) {
            const remainingLeads = leads.slice(3).map(lead => normalizeLeadData(lead));
            await appendDataToExcelTable(client, fileId, tableName, remainingLeads);
        }
        
    } catch (error) {
        console.error(`❌ Error creating Excel table:`, error);
        throw error;
    }
}

/**
 * Populate worksheet with initial data
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} worksheetName - Worksheet name
 * @param {Array} data - Data to populate
 */
async function populateWorksheetWithData(client, fileId, worksheetName, data) {
    try {
        if (!data || data.length === 0) return;
        
        // Get headers from the first data row
        const headers = Object.keys(data[0]);
        
        // Prepare data array with headers
        const tableData = [headers]; // Header row
        data.forEach(row => {
            const dataRow = headers.map(header => String(row[header] || ''));
            tableData.push(dataRow);
        });
        
        // Calculate range
        const numCols = headers.length;
        const numRows = tableData.length;
        const endCol = getExcelColumnLetter(numCols);
        const range = `A1:${endCol}${numRows}`;
        
        // Update worksheet range
        await client
            .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/range(address='${range}')`)
            .patch({
                values: tableData
            });
        
        console.log(`✅ Populated worksheet '${worksheetName}' with ${data.length} data rows`);
        
    } catch (error) {
        console.error(`❌ Error populating worksheet:`, error);
        throw error;
    }
}

/**
 * Append data to Excel table
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} tableName - Table name
 * @param {Array} leads - Lead data to append
 */
async function appendDataToExcelTable(client, fileId, tableName, leads) {
    try {
        if (!leads || leads.length === 0) {
            console.log(`⚠️ No data to append to table '${tableName}'`);
            return;
        }
        
        console.log(`➕ Appending ${leads.length} rows to table '${tableName}'`);
        
        // Normalize lead data
        const normalizedLeads = leads.map(lead => normalizeLeadData(lead));
        
        // Get table columns to ensure proper order
        const columnsResponse = await client
            .api(`/me/drive/items/${fileId}/workbook/tables/${tableName}/columns`)
            .get();
        
        const columns = columnsResponse.value;
        const headers = columns.map(col => col.name);
        
        console.log(`📋 Table headers: ${headers.join(', ')}`);
        
        // Prepare data rows in correct column order
        const tableRows = normalizedLeads.map(lead => {
            return headers.map(header => String(lead[header] || ''));
        });
        
        // Append rows to table using Graph API
        const appendRequest = {
            values: tableRows
        };
        
        await client
            .api(`/me/drive/items/${fileId}/workbook/tables/${tableName}/rows/add`)
            .post(appendRequest);
        
        console.log(`✅ Successfully appended ${tableRows.length} rows to table '${tableName}'`);
        
    } catch (error) {
        console.error(`❌ Error appending data to table:`, error);
        throw error;
    }
}

/**
 * Update lead tracking data in table
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} tableName - Table name
 * @param {string} leadEmail - Lead email to find
 * @param {Object} trackingData - Tracking data to update
 * @returns {boolean} Success status
 */
async function updateLeadTrackingInTable(client, fileId, tableName, leadEmail, trackingData) {
    try {
        // Get all table data
        const rowsResponse = await client
            .api(`/me/drive/items/${fileId}/workbook/tables/${tableName}/rows`)
            .get();
        
        const rows = rowsResponse.value;
        
        // Get column information
        const columnsResponse = await client
            .api(`/me/drive/items/${fileId}/workbook/tables/${tableName}/columns`)
            .get();
        
        const columns = columnsResponse.value;
        const headers = columns.map(col => col.name);
        const emailColIndex = headers.indexOf('Email');
        
        if (emailColIndex === -1) {
            throw new Error('Email column not found in table');
        }
        
        // Find the row with matching email
        let targetRowIndex = -1;
        for (let i = 0; i < rows.length; i++) {
            const rowValues = rows[i].values[0]; // Row values are in nested array
            if (rowValues[emailColIndex] && rowValues[emailColIndex].toLowerCase() === leadEmail.toLowerCase()) {
                targetRowIndex = i;
                break;
            }
        }
        
        if (targetRowIndex === -1) {
            console.log(`⚠️ Lead email '${leadEmail}' not found in table`);
            return false;
        }
        
        // Prepare updated row data
        const currentRow = rows[targetRowIndex].values[0];
        const updatedRow = [...currentRow];
        
        // Update tracking fields
        const updateFields = {
            'Email Sent': trackingData.sent ? 'Yes' : 'No',
            'Email Status': trackingData.status || 'Sent',
            'Sent Date': trackingData.sentDate || '',
            'Read Date': trackingData.readDate || '',
            'Reply Date': trackingData.replyDate || '',
            'Last Updated': new Date().toISOString()
        };
        
        // Apply updates
        Object.keys(updateFields).forEach(fieldName => {
            const colIndex = headers.indexOf(fieldName);
            if (colIndex !== -1) {
                updatedRow[colIndex] = updateFields[fieldName];
            }
        });
        
        // Update the row using Graph API
        const rowId = rows[targetRowIndex].id;
        await client
            .api(`/me/drive/items/${fileId}/workbook/tables/${tableName}/rows/${rowId}`)
            .patch({
                values: [updatedRow]
            });
        
        console.log(`✅ Updated tracking data for '${leadEmail}' in table '${tableName}'`);
        return true;
        
    } catch (error) {
        console.error(`❌ Error updating lead tracking in table:`, error);
        throw error;
    }
}

/**
 * Normalize lead data to standard format
 * @param {Object} lead - Raw lead data
 * @returns {Object} Normalized lead data
 */
function normalizeLeadData(lead) {
    return {
        'Name': lead.name || '',
        'Title': lead.title || '',
        'Company Name': lead.organization_name || lead.company || '',
        'Company Website': lead.organization_website_url || lead.website || '',
        'Size': lead.estimated_num_employees || lead.size || '',
        'Email': lead.email || '',
        'Email Verified': lead.email_verified || 'N',
        'LinkedIn URL': lead.linkedin_url || lead.linkedin || '',
        'Industry': lead.industry || '',
        'Location': lead.country || lead.location || '',
        'Notes': lead.notes || '',
        'Conversion Status': lead.conversion_status || 'Pending',
        'Email Sent': '',
        'Email Status': 'Not Sent',
        'Sent Date': '',
        'Read Date': '',
        'Reply Date': '',
        'Last Updated': new Date().toISOString()
    };
}

/**
 * Get Excel column letter from number (A, B, C, ... Z, AA, AB, etc.)
 * @param {number} columnNumber - Column number (1-based)
 * @returns {string} Excel column letter
 */
function getExcelColumnLetter(columnNumber) {
    let result = '';
    while (columnNumber > 0) {
        const remainder = (columnNumber - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        columnNumber = Math.floor((columnNumber - 1) / 26);
    }
    return result;
}

/**
 * Get column widths for Excel formatting
 * @returns {Array} Column width configuration
 */
function getColumnWidths() {
    return [
        {width: 20}, // Name
        {width: 25}, // Title
        {width: 30}, // Company Name
        {width: 35}, // Company Website
        {width: 15}, // Size
        {width: 30}, // Email
        {width: 15}, // Email Verified
        {width: 40}, // LinkedIn URL
        {width: 20}, // Industry
        {width: 15}, // Location
        {width: 50}, // Notes
        {width: 18}, // Conversion Status
        {width: 12}, // Email Sent
        {width: 15}, // Email Status
        {width: 20}, // Sent Date
        {width: 20}, // Read Date
        {width: 20}, // Reply Date
        {width: 20}  // Last Updated
    ];
}

module.exports = router;