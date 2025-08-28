const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const EmailContentProcessor = require('../utils/emailContentProcessor');
const router = express.Router();

// Configure multer for file uploads
const upload = multer({
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    },
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
            file.mimetype === 'application/vnd.ms-excel') {
            cb(null, true);
        } else {
            cb(new Error('Only Excel files (.xlsx, .xls) are allowed'), false);
        }
    }
});

// Initialize processors
const excelProcessor = new ExcelProcessor();
const emailContentProcessor = new EmailContentProcessor();

/**
 * Email Automation Master List Management
 * Handles Excel file operations, lead management, and campaign coordination
 */

// Upload and merge Excel file with master list
router.post('/master-list/upload', requireDelegatedAuth, upload.single('excelFile'), async (req, res) => {
    try {
        console.log('📤 Starting Excel file upload and merge...');

        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: 'No Excel file provided'
            });
        }

        console.log(`📊 Processing uploaded file: ${req.file.originalname} (${req.file.size} bytes)`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Parse uploaded Excel file
        const uploadedLeads = excelProcessor.parseUploadedFile(req.file.buffer);

        if (uploadedLeads.length === 0) {
            return res.status(400).json({
                success: false,
                message: 'No valid leads found in uploaded file'
            });
        }

        // Check if master file exists, create if not
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        let masterWorkbook;
        let existingData = [];

        console.log('🔍 Looking for existing master file in OneDrive...');

        try {
            // Try to download existing master file
            const files = await graphClient
                .api(`/me/drive/root:${masterFolderPath}:/children`)
                .filter(`name eq '${masterFileName}'`)
                .get();

            console.log(`📂 Found ${files.value.length} files in ${masterFolderPath} folder`);

            if (files.value.length > 0) {
                console.log('📋 Found existing master file, downloading...');
                const fileContent = await graphClient
                    .api(`/me/drive/items/${files.value[0].id}/content`)
                    .get();
                
                masterWorkbook = excelProcessor.bufferToWorkbook(fileContent);
                
                // Get existing leads data
                const leadsSheet = masterWorkbook.Sheets['Leads'];
                if (leadsSheet) {
                    existingData = XLSX.utils.sheet_to_json(leadsSheet);
                    console.log(`📊 Found ${existingData.length} existing leads in master file`);
                } else {
                    console.log('⚠️ Master file exists but has no Leads sheet');
                }
            } else {
                console.log('📋 No master file found, creating new one...');
                masterWorkbook = excelProcessor.createMasterFile();
            }
        } catch (error) {
            console.error('❌ Error accessing master file:', error.message);
            if (error.code === 'itemNotFound' || error.message.includes('not found')) {
                console.log('📋 Folder or file not found - creating new master file');
            } else {
                console.log('📋 Creating new master file due to access issue:', error.message);
            }
            masterWorkbook = excelProcessor.createMasterFile();
        }

        // Merge uploaded leads with existing data
        const mergeResults = excelProcessor.mergeLeadsWithMaster(uploadedLeads, existingData);

        if (mergeResults.newLeads.length === 0) {
            return res.json({
                success: true,
                message: 'No new leads to add - all leads already exist',
                totalProcessed: mergeResults.totalProcessed,
                newLeads: 0,
                duplicates: mergeResults.duplicates.length,
                duplicateDetails: mergeResults.duplicates
            });
        }

        // Update master file with new leads
        const updatedWorkbook = excelProcessor.updateMasterFileWithLeads(masterWorkbook, mergeResults.newLeads);

        // Create folder if it doesn't exist
        await createOneDriveFolder(graphClient, masterFolderPath);

        // Save updated master file to OneDrive
        const masterBuffer = excelProcessor.workbookToBuffer(updatedWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, masterFileName, masterFolderPath);

        console.log(`✅ Master file updated: ${mergeResults.newLeads.length} new leads added`);
        console.log(`🎉 Upload completed successfully - Master file ready in OneDrive`);

        res.json({
            success: true,
            message: 'Excel file uploaded and merged successfully',
            totalProcessed: mergeResults.totalProcessed,
            newLeads: mergeResults.newLeads.length,
            duplicates: mergeResults.duplicates.length,
            duplicateDetails: mergeResults.duplicates,
            masterFile: {
                name: masterFileName,
                location: masterFolderPath,
                totalLeads: existingData.length + mergeResults.newLeads.length
            }
        });

    } catch (error) {
        console.error('❌ Excel upload error:', error);
        
        // Handle specific error types
        if (error.isLockError) {
            res.status(423).json({
                success: false,
                message: 'File is currently locked',
                error: error.message,
                errorType: 'FILE_LOCKED',
                userMessage: 'The Excel file is currently open in OneDrive or Excel. Please close it and try again.',
                details: process.env.NODE_ENV === 'development' ? error.stack : undefined
            });
        } else {
            res.status(500).json({
                success: false,
                message: 'Failed to upload and process Excel file',
                error: error.message,
                details: process.env.NODE_ENV === 'development' ? error.stack : undefined
            });
        }
    }
});

// Get master list data
router.get('/master-list/data', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('📋 Retrieving master list data...');

        const { limit = 100, offset = 0, status, campaign_stage } = req.query;

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.json({
                success: true,
                data: [],
                total: 0,
                message: 'No master file found'
            });
        }

        // Get leads data
        const leadsSheet = masterWorkbook.Sheets['Leads'];
        let leadsData = XLSX.utils.sheet_to_json(leadsSheet);

        // Filter data if parameters provided
        if (status) {
            leadsData = leadsData.filter(lead => lead.Status === status);
        }
        if (campaign_stage) {
            leadsData = leadsData.filter(lead => lead.Campaign_Stage === campaign_stage);
        }

        // Apply pagination
        const total = leadsData.length;
        const paginatedData = leadsData.slice(parseInt(offset), parseInt(offset) + parseInt(limit));

        res.json({
            success: true,
            data: paginatedData,
            total: total,
            limit: parseInt(limit),
            offset: parseInt(offset),
            hasMore: (parseInt(offset) + parseInt(limit)) < total
        });

    } catch (error) {
        console.error('❌ Master list retrieval error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve master list',
            error: error.message
        });
    }
});

// Get master list statistics
router.get('/master-list/stats', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('📊 Calculating master list statistics...');

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.json({
                success: true,
                data: {
                    totalLeads: 0,
                    dueToday: 0,
                    emailsSent: 0,
                    emailsRead: 0,
                    repliesReceived: 0,
                    statusBreakdown: {}
                }
            });
        }

        // Calculate statistics
        const stats = excelProcessor.getMasterFileStats(masterWorkbook);

        res.json({
            success: true,
            data: stats
        });

    } catch (error) {
        console.error('❌ Statistics calculation error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to calculate statistics',
            error: error.message
        });
    }
});

// Update lead information
router.put('/master-list/lead/:email', requireDelegatedAuth, async (req, res) => {
    try {
        const { email } = req.params;
        const updates = req.body;

        console.log(`📝 Updating lead: ${email}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Update lead in master file
        const updatedWorkbook = excelProcessor.updateLeadInMaster(masterWorkbook, email, updates);

        // Save updated file
        const masterBuffer = excelProcessor.workbookToBuffer(updatedWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

        console.log(`✅ Lead updated: ${email}`);

        res.json({
            success: true,
            message: `Lead ${email} updated successfully`,
            updatedFields: Object.keys(updates)
        });

    } catch (error) {
        console.error('❌ Lead update error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to update lead',
            error: error.message
        });
    }
});

// Get leads due for email today
router.get('/master-list/due-today', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('📅 Getting leads due for email today...');

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.json({
                success: true,
                data: [],
                total: 0
            });
        }

        // Get leads due today
        const dueLeads = excelProcessor.getLeadsDueToday(masterWorkbook);

        res.json({
            success: true,
            data: dueLeads,
            total: dueLeads.length
        });

    } catch (error) {
        console.error('❌ Due leads retrieval error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve due leads',
            error: error.message
        });
    }
});

// Export master list to Excel
router.get('/master-list/export', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('📥 Exporting master list...');

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Convert to buffer and send
        const buffer = excelProcessor.workbookToBuffer(masterWorkbook);
        
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const filename = `LGA-Master-Email-List-Export-${timestamp}.xlsx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buffer);

    } catch (error) {
        console.error('❌ Export error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to export master list',
            error: error.message
        });
    }
});

// DEBUG: Inspect master file contents
router.get('/debug/master-file', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('🔍 DEBUG: Inspecting master file contents...');

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Inspect the workbook structure
        const sheets = Object.keys(masterWorkbook.Sheets);
        const debugInfo = {
            sheets: sheets,
            sheetDetails: {}
        };

        // Inspect each sheet
        sheets.forEach(sheetName => {
            const sheet = masterWorkbook.Sheets[sheetName];
            const range = sheet['!ref'];
            const data = XLSX.utils.sheet_to_json(sheet);
            
            debugInfo.sheetDetails[sheetName] = {
                range: range,
                rowCount: data.length,
                firstRow: data[0] || null,
                lastRow: data[data.length - 1] || null,
                sampleData: data.slice(0, 3)
            };
        });

        res.json({
            success: true,
            debugInfo: debugInfo
        });

    } catch (error) {
        console.error('❌ Debug inspection error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to inspect master file',
            error: error.message
        });
    }
});

// DEBUG: Test Excel file creation locally
router.post('/debug/test-excel-creation', async (req, res) => {
    try {
        console.log('🔍 DEBUG: Testing Excel file creation...');

        // Create a test Excel file with sample data
        const testLeads = [
            {
                Name: 'Test User 1',
                Email: 'test1@example.com',
                'Company Name': 'Test Company 1',
                Title: 'CEO'
            },
            {
                Name: 'Test User 2', 
                Email: 'test2@example.com',
                'Company Name': 'Test Company 2',
                Title: 'CTO'
            }
        ];

        const excelProcessor = new ExcelProcessor();
        const normalizedLeads = excelProcessor.normalizeLeadsData(testLeads);
        
        console.log('🔍 DEBUG: Normalized test leads:', normalizedLeads);

        // Create master file
        const masterWorkbook = excelProcessor.createMasterFile();
        
        // Update with test leads
        const updatedWorkbook = excelProcessor.updateMasterFileWithLeads(masterWorkbook, normalizedLeads);
        
        // Convert to buffer and inspect
        const buffer = excelProcessor.workbookToBuffer(updatedWorkbook);
        
        console.log('🔍 DEBUG: Created Excel buffer size:', buffer.length);

        // Read the buffer back to verify
        const verifyWorkbook = excelProcessor.bufferToWorkbook(buffer);
        const leadsSheet = verifyWorkbook.Sheets['Leads'];
        const verifyData = XLSX.utils.sheet_to_json(leadsSheet);

        res.json({
            success: true,
            testResults: {
                originalLeads: testLeads,
                normalizedLeads: normalizedLeads,
                bufferSize: buffer.length,
                verifiedData: verifyData,
                sheetRange: leadsSheet['!ref'],
                firstCell: leadsSheet['A1'],
                secondCell: leadsSheet['B1']
            }
        });

    } catch (error) {
        console.error('❌ Excel creation test error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to test Excel creation',
            error: error.message,
            stack: error.stack
        });
    }
});

// Send email to specific lead
router.post('/send-email/:email', requireDelegatedAuth, async (req, res) => {
    try {
        const { email } = req.params;
        const { emailChoice, customTemplate } = req.body;

        console.log(`📧 Sending email to: ${email} using ${emailChoice}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file to get lead data and templates
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Get lead data
        const leadsSheet = masterWorkbook.Sheets['Leads'];
        const leadsData = XLSX.utils.sheet_to_json(leadsSheet);
        const lead = leadsData.find(l => l.Email.toLowerCase() === email.toLowerCase());

        if (!lead) {
            return res.status(404).json({
                success: false,
                message: 'Lead not found'
            });
        }

        // Get templates
        const templates = excelProcessor.getTemplates(masterWorkbook);

        // Process email content
        const emailContent = await emailContentProcessor.processEmailContent(
            lead, 
            emailChoice || lead.Email_Choice || 'AI_Generated', 
            templates
        );

        // Validate email content
        const validation = emailContentProcessor.validateEmailContent(emailContent);
        if (!validation.isValid) {
            return res.status(400).json({
                success: false,
                message: 'Invalid email content',
                errors: validation.errors
            });
        }

        // Send email using Microsoft Graph
        const emailMessage = {
            subject: emailContent.subject,
            body: {
                contentType: 'HTML',
                content: emailContentProcessor.convertToHTML(emailContent)
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: lead.Email,
                        name: lead.Name
                    }
                }
            ]
        };

        await graphClient.api('/me/sendMail').post({
            message: emailMessage,
            saveToSentItems: true
        });

        // Update lead status in master file
        const updates = {
            Status: 'Sent',
            Last_Email_Date: new Date().toISOString().split('T')[0],
            Email_Count: (lead.Email_Count || 0) + 1,
            Template_Used: emailContent.contentType,
            Email_Content_Sent: emailContent.subject + '\n\n' + emailContent.body,
            Next_Email_Date: excelProcessor.calculateNextEmailDate(new Date(), lead.Follow_Up_Days || 7),
            'Email Sent': 'Yes',
            'Email Status': 'Sent',
            'Sent Date': new Date().toISOString()
        };

        const updatedWorkbook = excelProcessor.updateLeadInMaster(masterWorkbook, email, updates);
        const masterBuffer = excelProcessor.workbookToBuffer(updatedWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

        console.log(`✅ Email sent successfully to: ${email}`);

        res.json({
            success: true,
            message: `Email sent successfully to ${email}`,
            emailContent: {
                subject: emailContent.subject,
                contentType: emailContent.contentType,
                variables: emailContent.variables
            },
            leadUpdates: updates
        });

    } catch (error) {
        console.error('❌ Email sending error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to send email',
            error: error.message
        });
    }
});

// Helper function to download master file
async function downloadMasterFile(graphClient) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        console.log(`📥 Attempting to download master file from: ${masterFolderPath}`);
        
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();
            
        console.log(`📋 Found ${files.value.length} files matching master file name`);

        if (files.value.length === 0) {
            console.log('📋 Master file not found in OneDrive');
            return null;
        }

        console.log(`📥 Downloading master file: ${files.value[0].name} (${files.value[0].size} bytes)`);
        console.log(`🔍 DEBUG: File ID: ${files.value[0].id}`);
        console.log(`🔍 DEBUG: Download URL: /me/drive/items/${files.value[0].id}/content`);
        
        // Try multiple methods to download the file
        let fileContent = null;
        
        try {
            // Method 1: Direct content download
            console.log('🔍 DEBUG: Trying Method 1 - Direct content download');
            fileContent = await graphClient
                .api(`/me/drive/items/${files.value[0].id}/content`)
                .get();
                
            console.log(`🔍 DEBUG: Method 1 result - Content type: ${typeof fileContent}`);
            console.log(`🔍 DEBUG: Method 1 result - Content length: ${fileContent ? fileContent.length : 'undefined'}`);
        } catch (error) {
            console.log('⚠️ Method 1 failed:', error.message);
        }
        
        // If Method 1 failed or returned undefined, try Method 2
        if (!fileContent || fileContent.length === undefined) {
            try {
                console.log('🔍 DEBUG: Trying Method 2 - Stream download');
                const response = await graphClient
                    .api(`/me/drive/items/${files.value[0].id}/content`)
                    .getStream();
                    
                // Convert stream to buffer
                const chunks = [];
                for await (const chunk of response) {
                    chunks.push(chunk);
                }
                fileContent = Buffer.concat(chunks);
                
                console.log(`🔍 DEBUG: Method 2 result - Buffer length: ${fileContent.length}`);
            } catch (error) {
                console.log('⚠️ Method 2 failed:', error.message);
            }
        }
        
        // If both methods failed, try Method 3
        if (!fileContent || fileContent.length === undefined) {
            try {
                console.log('🔍 DEBUG: Trying Method 3 - Alternative path');
                fileContent = await graphClient
                    .api(`/me/drive/root:${masterFolderPath}/${masterFileName}:/content`)
                    .get();
                    
                console.log(`🔍 DEBUG: Method 3 result - Content type: ${typeof fileContent}`);
                console.log(`🔍 DEBUG: Method 3 result - Content length: ${fileContent ? fileContent.length : 'undefined'}`);
            } catch (error) {
                console.log('⚠️ Method 3 failed:', error.message);
            }
        }

        if (!fileContent) {
            console.error('❌ All download methods failed');
            return null;
        }

        console.log(`🔍 DEBUG: Final content type: ${typeof fileContent}`);
        console.log(`🔍 DEBUG: Content is Buffer: ${Buffer.isBuffer(fileContent)}`);
        console.log(`🔍 DEBUG: Content is ArrayBuffer: ${fileContent instanceof ArrayBuffer}`);

        // Handle different content types from Graph API
        let buffer;
        if (Buffer.isBuffer(fileContent)) {
            buffer = fileContent;
        } else if (fileContent instanceof ArrayBuffer) {
            buffer = Buffer.from(fileContent);
        } else if (typeof fileContent === 'string') {
            buffer = Buffer.from(fileContent, 'binary');
        } else {
            console.error('❌ Unexpected file content type:', typeof fileContent);
            return null;
        }

        console.log(`✅ Master file downloaded successfully - converted to buffer of ${buffer.length} bytes`);
        return excelProcessor.bufferToWorkbook(buffer);
    } catch (error) {
        console.error('❌ Master file download error:', error);
        return null;
    }
}

// Helper function to create OneDrive folder
async function createOneDriveFolder(client, folderPath) {
    try {
        console.log(`📂 Checking/creating OneDrive folder: ${folderPath}`);
        
        // Check if folder exists using the correct API path
        await client.api(`/me/drive/root:${folderPath}`).get();
        console.log(`📂 Folder ${folderPath} already exists`);
    } catch (error) {
        if (error.code === 'itemNotFound') {
            const folderName = folderPath.split('/').filter(Boolean).pop(); // Get last non-empty part
            
            console.log(`📂 Creating folder: ${folderName} in root directory`);
            
            // Create folder in root directory
            await client.api(`/me/drive/root/children`).post({
                name: folderName,
                folder: {},
                '@microsoft.graph.conflictBehavior': 'rename'
            });
            
            console.log(`✅ Created folder: ${folderPath}`);
        } else {
            console.error(`❌ Error checking/creating folder ${folderPath}:`, error);
            throw error;
        }
    }
}

// Helper function to upload file to OneDrive with retry logic
async function uploadToOneDrive(client, fileBuffer, filename, folderPath, maxRetries = 3) {
    // Construct the correct OneDrive API path
    const uploadUrl = `/me/drive/root:${folderPath}/${filename}:/content`;
    
    console.log(`📤 Uploading file: ${filename} to ${folderPath}`);
    console.log(`📤 Using OneDrive API URL: ${uploadUrl}`);
    console.log(`📊 File size: ${fileBuffer.length} bytes`);
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const result = await client.api(uploadUrl).put(fileBuffer);
            
            console.log(`✅ Successfully uploaded: ${filename} to ${folderPath}`);
            console.log(`📋 File details: ID=${result.id}, Size=${result.size} bytes`);
            console.log(`🔗 OneDrive URL: ${result.webUrl}`);
            
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
            
            console.error('❌ OneDrive upload error:', error);
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

module.exports = router;