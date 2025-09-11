const express = require('express');
const XLSX = require('xlsx');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const { getExcelColumnLetter } = require('../utils/excelGraphAPI');
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

        // Check if ANY table exists in the workbook (FIXED: prevents table overlap error)
        const existingTable = await getAnyExistingTable(graphClient, fileId, EXCEL_CONFIG.WORKSHEET_NAME);
        
        let tableInfo = null;
        let targetWorksheet = EXCEL_CONFIG.WORKSHEET_NAME;
        
        if (existingTable) {
            console.log(`✅ Found existing table: '${existingTable.name}' in worksheet '${existingTable.worksheet}'`);
            console.log(`📊 Table has columns: ${existingTable.columns.join(', ')}`);
            
            if (existingTable.isCompatible) {
                console.log(`✅ Table is compatible with lead data - will append to existing table`);
            } else {
                console.log(`⚠️ Table compatibility unclear - will still use it to prevent overlap error`);
            }
            
            tableInfo = existingTable;
            targetWorksheet = existingTable.worksheet || EXCEL_CONFIG.WORKSHEET_NAME;
        } else {
            console.log(`✅ No existing tables found - safe to create new table`);
        }
        
        if (!tableInfo) {
            // No table exists - safe to create new table without overlap
            console.log(`🆕 Creating new table '${EXCEL_CONFIG.TABLE_NAME}' in worksheet '${targetWorksheet}'`);
            await createExcelTable(graphClient, fileId, targetWorksheet, EXCEL_CONFIG.TABLE_NAME, leads);
        } else {
            // Table exists - append to it instead of creating new one (FIXES OVERLAP ERROR)
            console.log(`➕ Appending ${leads.length} leads to existing table '${tableInfo.name}' in '${targetWorksheet}'`);
            
            try {
                await appendDataToExcelTableWithRetry(graphClient, fileId, tableInfo.name, leads);
                console.log(`✅ Successfully appended data to existing table - no overlap error!`);
            } catch (appendError) {
                console.error(`❌ Failed to append to existing table '${tableInfo.name}': ${appendError.message}`);
                
                // If table is incompatible, offer to create new table in different worksheet
                if (appendError.message?.includes('column') || appendError.message?.includes('schema')) {
                    console.log(`🔄 Table schema incompatible - creating new table in different worksheet`);
                    const newWorksheet = targetWorksheet === 'Leads' ? 'NewLeads' : 'Leads2';
                    await createExcelTable(graphClient, fileId, newWorksheet, EXCEL_CONFIG.TABLE_NAME + '_New', leads);
                } else {
                    throw appendError;
                }
            }
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
        
        // Handle specific table overlap error
        if (error.code === 'InvalidArgument' && 
            error.message?.includes("A table can't overlap another table")) {
            
            console.error(`❌ TABLE OVERLAP ERROR: ${error.message}`);
            res.status(400).json({
                error: 'Table Overlap Error',
                message: 'Cannot create table because it would overlap with an existing table. The uploaded Excel file already contains a table structure.',
                solution: 'This should not happen with the new fix. Please report this error.',
                originalError: error.message,
                troubleshooting: {
                    suggestion: 'Try uploading the file again, or check if the Excel file has multiple conflicting tables',
                    technical: 'The system should now detect existing tables and append to them instead of creating new ones'
                }
            });
        } else {
            // Handle other errors normally
            res.status(500).json({
                error: 'OneDrive Excel Append Error',
                message: 'Failed to append data to Excel table in OneDrive',
                details: process.env.NODE_ENV === 'development' ? error.message : undefined
            });
        }
    }
});

// Legacy endpoint - redirects to new append functionality
router.post('/onedrive/create-excel', requireDelegatedAuth, async (req, res) => {
    console.log('📊 Creating Excel file with table structure...');
    req.body.useCustomFile = true; // Allow custom filename
    
    // Forward to append-to-table endpoint
    try {
        const { leads, filename, folderPath = '/LGA-Leads', useCustomFile = true } = req.body;

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
            // Table doesn't exist - convert uploaded Excel to table format
            console.log(`🔄 Excel file found but no table structure - converting to table format`);
            console.log(`🆕 Creating table '${EXCEL_CONFIG.TABLE_NAME}' in worksheet '${EXCEL_CONFIG.WORKSHEET_NAME}'`);
            await createExcelTable(graphClient, fileId, EXCEL_CONFIG.WORKSHEET_NAME, EXCEL_CONFIG.TABLE_NAME, leads);
        } else {
            // Table exists, but verify the exact name/ID before appending
            console.log(`🔍 Verifying existing table name before appending...`);
            const verifiedTableName = await verifyTableExistsWithPolling(graphClient, fileId, EXCEL_CONFIG.TABLE_NAME, 2);
            
            if (!verifiedTableName) {
                throw new Error(`Table '${EXCEL_CONFIG.TABLE_NAME}' verification failed`);
            }
            
            console.log(`➕ Appending data to verified table '${verifiedTableName}'`);
            await appendDataToExcelTableWithRetry(graphClient, fileId, verifiedTableName, leads);
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
        console.error('Excel creation error:', error);
        res.status(500).json({
            error: 'OneDrive Excel Create Error',
            message: 'Failed to create/append data to Excel table in OneDrive',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
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

// Convert existing Excel file to table format (MIGRATION ENDPOINT)
router.post('/onedrive/convert-to-table', requireDelegatedAuth, async (req, res) => {
    try {
        const { filePath, worksheetName = 'Leads', tableName = 'LeadsTable', forceConvert = false } = req.body;
        
        if (!filePath) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'File path is required'
            });
        }
        
        console.log(`🔄 Converting Excel file to table format: ${filePath}`);
        
        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Clean file path
        const cleanFilePath = filePath.startsWith('/') ? filePath.substring(1) : filePath;
        
        // Check if file exists
        const fileInfo = await getOneDriveFileInfo(graphClient, cleanFilePath);
        if (!fileInfo) {
            return res.status(404).json({
                error: 'File Not Found',
                message: `Excel file not found: ${filePath}`
            });
        }
        
        // Check if table already exists
        const tableInfo = await getExcelTableInfo(graphClient, fileInfo.id, worksheetName, tableName);
        
        if (tableInfo && !forceConvert) {
            return res.json({
                success: true,
                action: 'already_exists',
                message: `Table '${tableName}' already exists in worksheet '${worksheetName}'`,
                fileId: fileInfo.id,
                filename: fileInfo.name,
                tableId: tableInfo.id
            });
        }
        
        if (tableInfo && forceConvert) {
            console.log(`⚠️ Table exists but forceConvert=true, will recreate table`);
            // Delete existing table first
            try {
                await graphClient
                    .api(`/me/drive/items/${fileInfo.id}/workbook/tables/${tableName}`)
                    .delete();
                console.log(`🗑️ Deleted existing table '${tableName}'`);
            } catch (deleteError) {
                console.log(`⚠️ Could not delete existing table: ${deleteError.message}`);
            }
        }
        
        // Get existing data from worksheet
        const existingData = await getWorksheetData(graphClient, fileInfo.id, worksheetName);
        
        if (!existingData || existingData.length === 0) {
            return res.json({
                success: false,
                action: 'no_data',
                message: `No data found in worksheet '${worksheetName}' to convert`,
                fileId: fileInfo.id,
                filename: fileInfo.name
            });
        }
        
        // Convert existing data to table
        await convertExistingDataToTable(graphClient, fileInfo.id, worksheetName, tableName, existingData);
        
        console.log(`✅ Successfully converted file to table format`);
        
        res.json({
            success: true,
            action: 'converted',
            message: `Successfully converted ${existingData.length} rows to table format`,
            fileId: fileInfo.id,
            filename: fileInfo.name,
            filePath: filePath,
            worksheetName: worksheetName,
            tableName: tableName,
            rowsConverted: existingData.length,
            convertedAt: new Date().toISOString()
        });
        
    } catch (error) {
        console.error('❌ Table conversion failed:', error);
        res.status(500).json({
            success: false,
            error: 'Table Conversion Failed',
            message: error.message,
            details: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
    }
});

// Test Excel table append functionality
router.post('/onedrive/test-table-append', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('🧪 Testing Excel table append functionality with conversion...');
        
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
        return {
            id: response.id,
            name: response.name,
            webUrl: response.webUrl,
            size: response.size
        };
    } catch (error) {
        if (error.code === 'itemNotFound') {
            return null;
        }
        console.error(`❌ Error checking file ${filePath}:`, error);
        throw error;
    }
}

/**
 * Verify table exists with polling mechanism (proper solution for table creation timing)
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} expectedTableName - Expected table name
 * @param {number} maxAttempts - Maximum polling attempts
 * @returns {string|null} Actual table name/ID or null if not found
 */
async function verifyTableExistsWithPolling(client, fileId, expectedTableName, maxAttempts = 3) {
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            console.log(`🔍 Table verification attempt ${attempt}/${maxAttempts}...`);
            
            // Get all tables in the workbook
            const tablesResponse = await client
                .api(`/me/drive/items/${fileId}/workbook/tables`)
                .get();
            
            const tables = tablesResponse.value;
            console.log(`📊 Found ${tables.length} table(s) in workbook: ${tables.map(t => t.name).join(', ')}`);
            
            // Look for our expected table by name, or accept any table if only one exists
            let targetTable = tables.find(table => 
                table.name === expectedTableName || 
                table.name.toLowerCase() === expectedTableName.toLowerCase()
            );
            
            // If expected table not found but exactly one table exists, use it (Graph API naming issue)
            if (!targetTable && tables.length === 1) {
                targetTable = tables[0];
                console.log(`🔄 Expected '${expectedTableName}' but found '${targetTable.name}' - using actual table name`);
            }
            
            if (targetTable) {
                console.log(`✅ Table found: '${targetTable.name}' (ID: ${targetTable.id})`);
                return targetTable.name; // Return the actual name as stored in Excel
            }
            
            if (attempt < maxAttempts) {
                const waitTime = attempt * 3000; // 3s, 6s, 9s delays
                console.log(`⏳ Table not found yet, waiting ${waitTime/1000}s before next attempt...`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
            }
            
        } catch (error) {
            console.error(`❌ Table verification attempt ${attempt} failed:`, error.message);
            
            if (attempt < maxAttempts) {
                const waitTime = attempt * 3000;
                console.log(`⏳ Error occurred, waiting ${waitTime/1000}s before retry...`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
            }
        }
    }
    
    console.error(`❌ Table '${expectedTableName}' not found after ${maxAttempts} verification attempts`);
    return null;
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
        if (error.code === 'itemNotFound' || 
            error.code === 'InvalidArgument' || 
            error.message?.includes("doesn't exist") ||
            error.message?.includes('ItemNotFound')) {
            console.log(`⚠️ Table '${tableName}' not found in worksheet '${worksheetName}' - will convert to table format`);
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
        console.log(`🆕 Creating Excel file with uniform table processing for ${leads.length} leads`);
        
        // Create proper Excel workbook with all sheets (Leads, Templates, Campaign_History)
        const ExcelProcessor = require('../utils/excelProcessor');
        const excelProcessor = new ExcelProcessor();
        const wb = excelProcessor.createMasterFile([]);
        
        // Get headers for table creation (needed for createEmptyExcelTable)
        const sampleNormalized = normalizeLeadData({});
        const headers = Object.keys(sampleNormalized);
        
        // Note: createMasterFile() already creates all three sheets with proper structure
        
        // Convert to buffer
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        // Create folder structure if needed
        const folderPath = filePath.substring(0, filePath.lastIndexOf('/'));
        if (folderPath) {
            await createOneDriveFolder(client, '/' + folderPath);
        }
        
        // Upload empty file to OneDrive
        const uploadResult = await client
            .api(`/me/drive/root:/${filePath}:/content`)
            .put(excelBuffer);
        
        const fileId = uploadResult.id;
        console.log(`✅ Created empty Excel file: ${filePath} (ID: ${fileId})`);
        
        // Create empty table structure with headers only
        await createEmptyExcelTable(client, fileId, EXCEL_CONFIG.WORKSHEET_NAME, EXCEL_CONFIG.TABLE_NAME, headers);
        
        // Now append ALL leads uniformly via Graph API table operations
        if (leads.length > 0) {
            console.log(`🔍 Verifying empty table exists before appending all ${leads.length} leads...`);
            
            // Verify table existence 
            const verifiedTableName = await verifyTableExistsWithPolling(client, fileId, EXCEL_CONFIG.TABLE_NAME, 3);
            
            if (!verifiedTableName) {
                throw new Error(`Table '${EXCEL_CONFIG.TABLE_NAME}' was not found after creation`);
            }
            
            console.log(`✅ Empty table verified as '${verifiedTableName}' - proceeding with uniform append of all leads`);
            
            // Process all leads uniformly through Graph API table append
            await appendDataToExcelTableWithRetry(client, fileId, verifiedTableName, leads, 5);
        }
        
        return fileId;
        
    } catch (error) {
        console.error(`❌ Error creating Excel file with uniform table processing:`, error);
        throw error;
    }
}

/**
 * Create empty Excel table with headers only (for uniform data processing)
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} worksheetName - Worksheet name
 * @param {string} tableName - Table name
 * @param {Array} headers - Column headers
 */
async function createEmptyExcelTable(client, fileId, worksheetName, tableName, headers) {
    try {
        console.log(`🗂️ Creating empty table '${tableName}' with ${headers.length} columns`);
        
        // Write only the header row to establish table structure
        await client
            .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/range(address='A1:${getExcelColumnLetter(headers.length - 1)}1')`)
            .patch({
                values: [headers]
            });
        
        // Create table from header row only (A1 to last header column)
        const tableRange = `A1:${getExcelColumnLetter(headers.length - 1)}1`;
        
        const tableRequest = {
            address: tableRange,
            hasHeaders: true,
            name: tableName
        };
        
        await client
            .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/tables/add`)
            .post(tableRequest);
        
        console.log(`✅ Created empty table '${tableName}' with headers: ${headers.join(', ')}`);
        
    } catch (error) {
        console.error(`❌ Error creating empty table:`, error);
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
        const endCol = getExcelColumnLetter(numCols - 1);
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
 * Create Excel table for existing file (handles existing data safely)
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} worksheetName - Worksheet name
 * @param {string} tableName - Table name
 * @param {Array} newLeads - New lead data to append
 */
async function createExcelTable(client, fileId, worksheetName, tableName, newLeads) {
    try {
        console.log(`🔄 Creating table '${tableName}' for existing file...`);
        
        // Step 1: Check if worksheet has existing data
        const existingData = await getWorksheetData(client, fileId, worksheetName);
        
        if (existingData && existingData.length > 0) {
            console.log(`📊 Found ${existingData.length} existing rows in worksheet`);
            
            // Step 2: Convert existing data to table format
            await convertExistingDataToTable(client, fileId, worksheetName, tableName, existingData);
            
            // Step 3: Append new data to the newly created table  
            if (newLeads && newLeads.length > 0) {
                console.log(`➕ Appending ${newLeads.length} new leads to converted table`);
                await appendDataToExcelTableWithRetry(client, fileId, tableName, newLeads);
            }
        } else {
            console.log(`📝 No existing data found, creating empty table and appending all leads uniformly`);
            
            // Get standard column structure
            const sampleNormalized = normalizeLeadData({});
            const headers = Object.keys(sampleNormalized);
            
            // Create empty table with headers only
            await createEmptyExcelTable(client, fileId, worksheetName, tableName, headers);
            
            // Append all new leads uniformly via Graph API
            if (newLeads.length > 0) {
                console.log(`⏳ Waiting briefly for empty table to be ready...`);
                await new Promise(resolve => setTimeout(resolve, 2000));
                await appendDataToExcelTableWithRetry(client, fileId, tableName, newLeads);
            }
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
        const endCol = getExcelColumnLetter(numCols - 1);
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
 * Append data to Excel table with retry logic
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} tableName - Table name
 * @param {Array} leads - Lead data to append
 * @param {number} maxRetries - Maximum retry attempts
 */
async function appendDataToExcelTableWithRetry(client, fileId, tableName, leads, maxRetries = 3) {
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            console.log(`🔄 Append attempt ${attempt}/${maxRetries} for table '${tableName}'`);
            await appendDataToExcelTable(client, fileId, tableName, leads);
            console.log(`✅ Append successful on attempt ${attempt}`);
            return;
        } catch (error) {
            console.log(`❌ Append attempt ${attempt} failed: ${error.message}`);
            
            if (attempt === maxRetries) {
                console.error(`❌ All ${maxRetries} append attempts failed`);
                throw error;
            }
            
            // For new table creation issues, wait longer between retries
            let waitTime;
            if (error.code === 'ItemNotFound' || error.message?.includes("doesn't exist")) {
                // Table indexing issue - use longer delays
                waitTime = Math.min(attempt * 5000, 15000); // 5s, 10s, 15s for table indexing
                console.log(`🔄 Table indexing issue detected - waiting ${waitTime/1000}s before retry...`);
            } else {
                // Other issues - standard exponential backoff
                waitTime = Math.min(attempt * 3000, 10000); // 3s, 6s, 9s (max 10s)
                console.log(`⏳ Waiting ${waitTime/1000}s before retry...`);
            }
            await new Promise(resolve => setTimeout(resolve, waitTime));
        }
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
        
        // Normalize lead data
        const normalizedLeads = leads.map(lead => normalizeLeadData(lead));
        
        // Get table columns to ensure proper order
        const columnsResponse = await client
            .api(`/me/drive/items/${fileId}/workbook/tables/${tableName}/columns`)
            .get();
        
        const columns = columnsResponse.value;
        const headers = columns.map(col => col.name);
        
        // Prepare data rows in correct column order with proper data types
        const tableRows = normalizedLeads.map(lead => {
            return headers.map(header => {
                const value = lead[header];
                // Handle different data types appropriately
                return formatCellValueForExcel(header, value);
            });
        });
        
        // Append rows to table using Graph API
        const appendRequest = {
            values: tableRows
        };
        
        await client
            .api(`/me/drive/items/${fileId}/workbook/tables/${tableName}/rows/add`)
            .post(appendRequest);
        
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
            'Sent Date': trackingData.sentDate || '',
            'Read Date': trackingData.readDate || '',
            'Reply Date': trackingData.replyDate || '',
            'Last Updated': require('../utils/dateFormatter').getCurrentFormattedDate()
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
        // Basic lead information
        'Name': lead['Name'] || lead.name || '',
        'Title': lead['Title'] || lead.title || '',
        'Company Name': lead['Company Name'] || lead.organization_name || lead.company || '',
        'Company Website': lead['Company Website'] || lead.organization_website_url || lead.website || '',
        'Size': lead['Size'] || lead.estimated_num_employees || lead.size || '',
        'Email': lead['Email'] || lead.email || '',
        'Email Verified': lead['Email Verified'] || lead.email_verified || 'Y',
        'LinkedIn URL': lead['LinkedIn URL'] || lead.linkedin_url || lead.linkedin || '',
        'Industry': lead['Industry'] || lead.industry || '',
        'Location': lead['Location'] || lead.country || lead.location || '',
        'Last Updated': lead['Last Updated'] || require('../utils/dateFormatter').getCurrentFormattedDate(),
        
        // Email automation columns
        'AI_Generated_Email': lead['AI_Generated_Email'] || '',
        'Status': lead['Status'] || 'New',
        'Campaign_Stage': lead['Campaign_Stage'] || 'First_Contact',
        'Template_Used': lead['Template_Used'] || '',
        'Last_Email_Date': lead['Last_Email_Date'] || '',
        'Next_Email_Date': lead['Next_Email_Date'] || '',
        'Follow_Up_Days': lead['Follow_Up_Days'] || 7,
        'Email_Count': lead['Email_Count'] || 0,
        'Read_Date': lead['Read_Date'] || '',
        'Reply_Date': lead['Reply_Date'] || '',
        'Email Bounce': lead['Email Bounce'] || 'No', // Initialize bounce status
        
        // Legacy compatibility columns removed
        'Sent Date': lead['Sent Date'] || ''
    };
}

/**
 * Get existing data from worksheet
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} worksheetName - Worksheet name
 * @returns {Array|null} Existing data or null if no data
 */
async function getWorksheetData(client, fileId, worksheetName) {
    try {
        console.log(`🔍 Checking for existing data in worksheet '${worksheetName}'...`);
        
        // Get used range of the worksheet
        const rangeResponse = await client
            .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/usedRange`)
            .get();
        
        if (!rangeResponse || !rangeResponse.values || rangeResponse.values.length === 0) {
            console.log(`📄 Worksheet '${worksheetName}' is empty`);
            return null;
        }
        
        const values = rangeResponse.values;
        console.log(`📊 Found ${values.length} rows in worksheet`);
        
        // Convert array of arrays to array of objects
        if (values.length < 2) {
            // Only header row or no data
            console.log(`⚠️ Only header row found, treating as empty`);
            return null;
        }
        
        const headers = values[0]; // First row as headers
        const dataRows = values.slice(1); // Rest as data
        
        const existingData = dataRows.map(row => {
            const dataObj = {};
            headers.forEach((header, index) => {
                dataObj[header] = row[index] || '';
            });
            return dataObj;
        });
        
        console.log(`✅ Processed ${existingData.length} existing rows`);
        return existingData;
        
    } catch (error) {
        if (error.code === 'itemNotFound' || error.code === 'InvalidArgument') {
            console.log(`📄 Worksheet '${worksheetName}' appears to be empty or invalid range`);
            return null;
        }
        console.error(`❌ Error reading worksheet data:`, error);
        throw error;
    }
}

/**
 * Convert existing data to table format
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} worksheetName - Worksheet name
 * @param {string} tableName - Table name
 * @param {Array} existingData - Existing data from worksheet
 */
async function convertExistingDataToTable(client, fileId, worksheetName, tableName, existingData) {
    try {
        console.log(`🔄 Converting ${existingData.length} existing rows to table format...`);
        
        // Step 1: Normalize existing data to match our expected structure
        const normalizedExistingData = existingData.map(row => {
            // Try to map existing columns to our standard structure
            const normalizedRow = normalizeLeadData(row);
            
            // Preserve any additional columns that exist in the original data
            Object.keys(row).forEach(key => {
                if (!normalizedRow.hasOwnProperty(key) && row[key]) {
                    normalizedRow[key] = row[key];
                }
            });
            
            return normalizedRow;
        });
        
        // Step 2: Get all unique headers from both existing and standard structure
        const allHeaders = new Set();
        normalizedExistingData.forEach(row => {
            Object.keys(row).forEach(key => allHeaders.add(key));
        });
        
        const finalHeaders = Array.from(allHeaders);
        console.log(`📋 Table configured with ${finalHeaders.length} columns`);
        
        // Step 3: Clear the worksheet and rewrite with normalized data
        console.log(`🧹 Clearing worksheet to prepare for table creation...`);
        
        // Clear the used range
        try {
            await client
                .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/usedRange`)
                .patch({ values: [[]] }); // Clear content
        } catch (clearError) {
            console.log(`⚠️ Could not clear worksheet (might already be empty): ${clearError.message}`);
        }
        
        // Step 4: Write headers and data in table format
        const tableData = [finalHeaders]; // Header row
        normalizedExistingData.forEach(row => {
            const dataRow = finalHeaders.map(header => String(row[header] || ''));
            tableData.push(dataRow);
        });
        
        // Calculate range for the table
        const numCols = finalHeaders.length;
        const numRows = tableData.length;
        const endCol = getExcelColumnLetter(numCols - 1);
        const tableRange = `A1:${endCol}${numRows}`;
        
        console.log(`📝 Writing ${normalizedExistingData.length} rows to table range...`);
        
        // Write the data to worksheet
        await client
            .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/range(address='${tableRange}')`)
            .patch({
                values: tableData
            });
        
        // Step 5: Create table from the written data
        console.log(`🗂️ Creating table '${tableName}' from existing data...`);
        
        const tableRequest = {
            address: tableRange,
            hasHeaders: true,
            name: tableName
        };
        
        await client
            .api(`/me/drive/items/${fileId}/workbook/worksheets/${worksheetName}/tables/add`)
            .post(tableRequest);
        
        console.log(`✅ Successfully converted ${normalizedExistingData.length} existing rows to table '${tableName}'`);
        
    } catch (error) {
        console.error(`❌ Error converting existing data to table:`, error);
        throw error;
    }
}


/**
 * Format cell value with proper data type for Excel table operations
 * @param {string} header - Column header name
 * @param {*} value - Cell value
 * @returns {*} Properly formatted value for Excel
 */
function formatCellValueForExcel(header, value) {
    // Handle null/undefined values
    if (value === null || value === undefined) {
        return '';
    }
    
    // Date fields - ensure proper ISO format for Excel date recognition
    const dateFields = ['Last Updated', 'Last_Email_Date', 'Next_Email_Date', 'Read_Date', 'Reply_Date', 'Sent Date'];
    if (dateFields.includes(header)) {
        if (!value || value === '') return '';
        
        try {
            // Ensure valid date format for Excel
            const date = new Date(value);
            if (isNaN(date.getTime())) return '';
            return date.toISOString().split('T')[0]; // YYYY-MM-DD format
        } catch {
            return '';
        }
    }
    
    // Number fields - ensure proper numeric type
    const numberFields = ['Follow_Up_Days', 'Email_Count'];
    if (numberFields.includes(header)) {
        const numValue = parseInt(value, 10);
        return isNaN(numValue) ? 0 : numValue;
    }
    
    // Boolean-like fields - standardize to Yes/No
    const booleanFields = ['Email Verified'];
    if (booleanFields.includes(header)) {
        if (!value || value === '') return 'No';
        const val = String(value).toLowerCase();
        if (val === 'true' || val === 'yes' || val === 'y' || val === '1') return 'Yes';
        return 'No';
    }
    
    // Email fields - ensure lowercase normalization
    if (header === 'Email') {
        return String(value || '').toLowerCase().trim();
    }
    
    // All other text fields - ensure string type and trim whitespace
    return String(value || '').trim();
}

/**
 * Get column widths for Excel formatting
 * @returns {Array} Column width configuration
 */
function getColumnWidths() {
    return [
        // Basic lead information
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
        {width: 20}, // Last Updated
        
        // Email automation columns
        {width: 50}, // AI_Generated_Email
        {width: 15}, // Status
        {width: 20}, // Campaign_Stage
        {width: 20}, // Template_Used
        {width: 18}, // Last_Email_Date
        {width: 18}, // Next_Email_Date
        {width: 15}, // Follow_Up_Days
        {width: 12}, // Email_Count
        {width: 18}, // Read_Date
        {width: 18}, // Reply_Date
        
        // Legacy compatibility columns removed
        {width: 20}  // Sent Date
    ];
}

/**
 * Get ANY existing table in the workbook (fixes table overlap issue)
 * @param {Object} client - Microsoft Graph client
 * @param {string} fileId - OneDrive file ID
 * @param {string} preferredWorksheet - Preferred worksheet name
 * @returns {Object|null} Table information or null if no tables found
 */
async function getAnyExistingTable(client, fileId, preferredWorksheet = 'Leads') {
    try {
        console.log(`🔍 Looking for ANY existing table in workbook...`);
        
        // Get all tables in the workbook first
        const tablesResponse = await client
            .api(`/me/drive/items/${fileId}/workbook/tables`)
            .get();
        
        const tables = tablesResponse.value;
        console.log(`📊 Found ${tables.length} existing table(s): ${tables.map(t => t.name).join(', ')}`);
        
        if (tables.length === 0) {
            console.log(`✅ No existing tables found - safe to create new table`);
            return null;
        }
        
        // If we have tables, find the best one to use
        // Priority: 1) Table in preferred worksheet, 2) Any table with headers that match lead data
        for (const table of tables) {
            try {
                // Get table details including worksheet
                const tableDetails = await client
                    .api(`/me/drive/items/${fileId}/workbook/tables/${table.name}`)
                    .get();
                
                // Get columns to check if it's compatible with lead data
                const columnsResponse = await client
                    .api(`/me/drive/items/${fileId}/workbook/tables/${table.name}/columns`)
                    .get();
                
                const columns = columnsResponse.value;
                const headers = columns.map(col => col.name);
                
                // Check if this table has lead-compatible columns (Email, Name, Company, etc.)
                const leadColumns = ['Email', 'Name', 'Company', 'Title'];
                const hasLeadColumns = leadColumns.some(col => 
                    headers.some(header => header.toLowerCase().includes(col.toLowerCase()))
                );
                
                if (hasLeadColumns) {
                    console.log(`✅ Found compatible existing table: '${table.name}' with columns: ${headers.join(', ')}`);
                    return {
                        id: tableDetails.id,
                        name: tableDetails.name,
                        range: tableDetails.range,
                        worksheet: tableDetails.worksheet?.name || 'Unknown',
                        columns: headers,
                        isCompatible: true
                    };
                }
            } catch (tableError) {
                console.log(`⚠️ Could not check table '${table.name}': ${tableError.message}`);
                continue;
            }
        }
        
        // If no compatible table found, return the first table anyway (user can decide)
        if (tables.length > 0) {
            const firstTable = tables[0];
            console.log(`⚠️ Found table '${firstTable.name}' but compatibility unclear - will use anyway to avoid overlap`);
            
            try {
                const tableDetails = await client
                    .api(`/me/drive/items/${fileId}/workbook/tables/${firstTable.name}`)
                    .get();
                
                const columnsResponse = await client
                    .api(`/me/drive/items/${fileId}/workbook/tables/${firstTable.name}/columns`)
                    .get();
                
                return {
                    id: tableDetails.id,
                    name: tableDetails.name,
                    range: tableDetails.range,
                    worksheet: tableDetails.worksheet?.name || 'Unknown',
                    columns: columnsResponse.value.map(col => col.name),
                    isCompatible: false
                };
            } catch (detailError) {
                console.log(`⚠️ Could not get details for table '${firstTable.name}': ${detailError.message}`);
                return {
                    id: firstTable.id,
                    name: firstTable.name,
                    range: null,
                    worksheet: 'Unknown',
                    columns: [],
                    isCompatible: false
                };
            }
        }
        
        return null;
        
    } catch (error) {
        console.error(`❌ Error checking for existing tables:`, error);
        // Don't throw error - just return null to allow table creation
        return null;
    }
}

module.exports = router;