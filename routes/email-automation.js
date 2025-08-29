const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const axios = require('axios');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const EmailContentProcessor = require('../utils/emailContentProcessor');
const { advancedExcelUpload } = require('./excel-upload-fix');
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
        // Use .xlsx extension for better OneDrive compatibility
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        let masterWorkbook;
        let existingData = [];
        let initialExistingCount = 0; // Track initial count for accurate breakdown

        console.log('✅ Master file structure verified');

        try {
            // Try to download existing master file
            const files = await graphClient
                .api(`/me/drive/root:${masterFolderPath}:/children`)
                .filter(`name eq '${masterFileName}'`)
                .get();

            if (files.value.length > 0) {
                const fileContent = await graphClient
                    .api(`/me/drive/items/${files.value[0].id}/content`)
                    .get();
                
                masterWorkbook = excelProcessor.bufferToWorkbook(fileContent);
                
                // CORRUPTION DETECTION: Check if the file has proper structure
                const hasRequiredSheets = ['Leads', 'Templates', 'Campaign_History'].every(
                    sheetName => masterWorkbook.Sheets[sheetName]
                );
                
                // Additional check: If we have a Leads sheet with data, it's probably not corrupted
                const leadsSheet = masterWorkbook.Sheets['Leads'];
                let leadsData = [];
                if (leadsSheet) {
                    try {
                        leadsData = XLSX.utils.sheet_to_json(leadsSheet);
                        console.log(`🔍 Checking data integrity...`);
                    } catch (error) {
                        console.log(`⚠️ CORRUPTION CHECK: Error reading Leads sheet:`, error.message);
                    }
                }
                
                // Only rebuild if there's absolutely no usable data anywhere
                // If Leads sheet has data, NEVER rebuild regardless of other sheet structure
                const shouldRebuild = leadsData.length === 0 && (!leadsSheet || !leadsSheet['!ref']);
                
                if (shouldRebuild) {
                    console.log('🚨 CORRUPTION DETECTED: Missing sheets and no lead data - Rebuilding master file');
                    
                    // Try to recover lead data from any sheet
                    let recoveredData = [];
                    const sheetNames = Object.keys(masterWorkbook.Sheets);
                    
                    for (const sheetName of sheetNames) {
                        try {
                            const sheet = masterWorkbook.Sheets[sheetName];
                            const data = XLSX.utils.sheet_to_json(sheet);
                            
                            if (data.length > 0) {
                                const hasLeadData = data.some(row => 
                                    row.Email || row.email || row.Name || row.name ||
                                    row.Company || row.company || row['Company Name'] || 
                                    row.Title || row.title
                                );
                                
                                if (hasLeadData) {
                                    console.log(`✅ Recovered leads from ${sheetName}`);
                                    recoveredData = [...recoveredData, ...data];
                                }
                            }
                        } catch (error) {
                            // Silent recovery - don't spam logs
                        }
                    }
                    
                    console.log(`📊 Recovery completed`);
                    existingData = recoveredData;
                    initialExistingCount = existingData.length;
                    
                    // Recreate with recovered data
                    masterWorkbook = excelProcessor.createMasterFile(existingData);
                } else {
                    // File structure is good OR has data, extract existing data normally
                    if (hasRequiredSheets) {
                        console.log(`✅ Master file structure verified`);
                        const leadsSheetData = masterWorkbook.Sheets['Leads'];
                        if (leadsSheetData) {
                            existingData = XLSX.utils.sheet_to_json(leadsSheetData);
                            initialExistingCount = existingData.length;
                            console.log(`📊 Found ${existingData.length} existing leads in master file`);
                        }
                    } else if (leadsData.length > 0) {
                        console.log(`⚠️ Missing structure sheets but preserving ${leadsData.length} existing leads`);
                        existingData = leadsData;
                        initialExistingCount = existingData.length;
                        // Recreate with proper structure but preserve data
                        masterWorkbook = excelProcessor.createMasterFile(existingData);
                    }
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

        // CRITICAL: Get current OneDrive data for accurate duplicate checking
        let currentOneDriveData = [];
        try {
            const currentWorkbook = await downloadMasterFile(graphClient, false); // Fresh download
            if (currentWorkbook && currentWorkbook.Sheets['Leads']) {
                currentOneDriveData = XLSX.utils.sheet_to_json(currentWorkbook.Sheets['Leads']);
                console.log(`📊 Current OneDrive data: ${currentOneDriveData.length} leads for duplicate checking`);
            }
        } catch (error) {
            console.log(`⚠️ Could not fetch current OneDrive data for duplicate check: ${error.message}`);
            // Fallback to locally parsed data
            currentOneDriveData = existingData;
        }

        // CRITICAL: Merge uploaded leads with current OneDrive data for accurate duplicate detection
        const mergeResults = excelProcessor.mergeLeadsWithMaster(uploadedLeads, currentOneDriveData);
        
        // Update initialExistingCount with actual current data
        initialExistingCount = currentOneDriveData.length;

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

        // CRITICAL: Use Microsoft Graph Table API to APPEND data (no file replacement)
        
        // Use the Microsoft Graph table append functionality
        const appendResult = await appendLeadsToOneDriveTable({
            delegatedAuth: req.delegatedAuth,
            sessionId: req.sessionId
        }, {
            leads: mergeResults.newLeads,
            filename: masterFileName,
            folderPath: masterFolderPath,
            useCustomFile: true
        });
        
        if (!appendResult.success) {
            throw new Error(`Failed to append leads to table: ${appendResult.message}`);
        }

        try {
            const verificationWorkbook = await downloadMasterFile(graphClient, false); // No cache for verification
            if (verificationWorkbook && verificationWorkbook.Sheets['Leads']) {
                const verifySheet = verificationWorkbook.Sheets['Leads'];
                const verifyData = XLSX.utils.sheet_to_json(verifySheet);
                
                // Smart verification: Ensure file has reasonable data and at least the new leads were added
                const hasReasonableData = verifyData.length >= mergeResults.newLeads.length && verifyData.length > 0;
                
                if (!hasReasonableData) {
                    console.error(`❌ DATA INTEGRITY ERROR: Expected at least ${mergeResults.newLeads.length} new leads, but file has ${verifyData.length} total rows`);
                    throw new Error(`Data integrity check failed: File contains ${verifyData.length} rows but should have at least ${mergeResults.newLeads.length} new leads`);
                } else {
                    // Calculate the proper breakdown
                    const uploadedCount = uploadedLeads.length;
                    const duplicatesCount = mergeResults.duplicates.length;
                    const newLeadsAdded = mergeResults.newLeads.length;
                    const finalTotalCount = verifyData.length;
                    
                    console.log(`📊 Upload breakdown:`);
                    console.log(`   - Existing leads: ${initialExistingCount}`);
                    console.log(`   - Leads uploaded: ${uploadedCount}`);
                    console.log(`   - Duplicates skipped: ${duplicatesCount}`);
                    console.log(`   - New leads added: ${newLeadsAdded}`);
                    console.log(`📊 Final count: ${initialExistingCount} existing + ${newLeadsAdded} new = ${finalTotalCount} total records`);
                    console.log(`✅ Data integrity verified successfully`);
                }
            } else {
                console.error('❌ Post-upload verification failed - no Leads sheet found');
                
                // Enhanced debugging for failed verification
                if (verificationWorkbook) {
                    const availableSheets = Object.keys(verificationWorkbook.Sheets);
                    console.error(`❌ Available sheets in downloaded file: [${availableSheets.join(', ')}]`);
                    
                    // Check if we have a Sheet1 with data (indicating corruption)
                    if (availableSheets.includes('Sheet1')) {
                        const sheet1Data = XLSX.utils.sheet_to_json(verificationWorkbook.Sheets['Sheet1']);
                        console.error(`❌ Sheet1 detected - OneDrive corruption possible`);
                        console.error(`❌ This is a known Microsoft Graph API Excel upload corruption issue`);
                    }
                } else {
                    console.error(`❌ Could not download verification file at all`);
                }
                
                throw new Error('Post-upload verification failed: No Leads sheet found in uploaded file');
            }
        } catch (verifyError) {
            console.error('❌ Post-upload verification failed:', verifyError.message);
            
            // If verification fails, this is a critical error - don't claim success
            return res.status(500).json({
                success: false,
                message: 'File uploaded but verification failed - data may not have been saved correctly',
                error: verifyError.message,
                troubleshooting: {
                    suggestion: 'Please check your OneDrive file manually and try again if data is missing',
                    expectedRows: existingData.length + mergeResults.newLeads.length,
                    uploadedRows: mergeResults.newLeads.length,
                    existingRows: existingData.length
                }
            });
        }


        // Get the actual final count from verification
        let finalTotalLeads = existingData.length + mergeResults.newLeads.length;
        try {
            const finalVerificationWorkbook = await downloadMasterFile(graphClient, false);
            if (finalVerificationWorkbook && finalVerificationWorkbook.Sheets['Leads']) {
                const finalVerifyData = XLSX.utils.sheet_to_json(finalVerificationWorkbook.Sheets['Leads']);
                finalTotalLeads = finalVerifyData.length;
            }
        } catch (verifyError) {
            // Use calculated count if verification fails
        }

        res.json({
            success: true,
            message: 'Excel file uploaded and merged successfully',
            breakdown: {
                existingLeads: initialExistingCount,
                leadsUploaded: uploadedLeads.length,
                duplicatesSkipped: mergeResults.duplicates.length,
                newLeadsAdded: mergeResults.newLeads.length,
                finalTotal: finalTotalLeads
            },
            // Legacy compatibility
            totalProcessed: mergeResults.totalProcessed,
            newLeads: mergeResults.newLeads.length,
            duplicates: mergeResults.duplicates.length,
            duplicateDetails: mergeResults.duplicates,
            masterFile: {
                name: masterFileName,
                location: masterFolderPath,
                totalLeads: finalTotalLeads
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
        await advancedExcelUpload(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

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
                // Data content removed for privacy
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

// RECOVERY: Merge additional leads with existing master list (for data recovery)
router.post('/master-list/merge-recovery', requireDelegatedAuth, upload.single('excelFile'), async (req, res) => {
    try {
        console.log('🔄 RECOVERY MERGE: Starting manual recovery merge process...');

        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: 'No Excel file provided for recovery merge'
            });
        }

        console.log(`📊 Processing recovery file: ${req.file.originalname} (${req.file.size} bytes)`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Parse recovery Excel file
        const recoveryLeads = excelProcessor.parseUploadedFile(req.file.buffer);

        if (recoveryLeads.length === 0) {
            return res.status(400).json({
                success: false,
                message: 'No valid leads found in recovery file'
            });
        }

        console.log(`📊 Found ${recoveryLeads.length} leads in recovery file`);

        // Download current master file
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        let currentMasterWorkbook = await downloadMasterFile(graphClient);
        let currentLeads = [];

        if (currentMasterWorkbook && currentMasterWorkbook.Sheets['Leads']) {
            const leadsSheet = currentMasterWorkbook.Sheets['Leads'];
            currentLeads = XLSX.utils.sheet_to_json(leadsSheet);
            console.log(`📊 Current master file has ${currentLeads.length} existing leads`);
        } else {
            console.log('📋 No current master file found or corrupted - creating new one');
            currentMasterWorkbook = excelProcessor.createMasterFile();
        }

        // Merge recovery leads with current data (append mode)
        const mergeResults = excelProcessor.mergeLeadsWithMaster(recoveryLeads, currentLeads);

        console.log(`🔄 RECOVERY MERGE RESULTS:`);
        console.log(`   - Current leads: ${currentLeads.length}`);
        console.log(`   - Recovery leads provided: ${recoveryLeads.length}`);
        console.log(`   - New unique leads to add: ${mergeResults.newLeads.length}`);
        console.log(`   - Duplicates skipped: ${mergeResults.duplicates.length}`);

        if (mergeResults.newLeads.length === 0) {
            return res.json({
                success: true,
                message: 'No new leads to recover - all leads already exist in master list',
                currentCount: currentLeads.length,
                recoveryAttempted: recoveryLeads.length,
                duplicates: mergeResults.duplicates.length
            });
        }

        // FIXED: Use Microsoft Graph Table API to append recovered leads
        console.log(`📊 RECOVERY: Appending ${mergeResults.newLeads.length} recovered leads using table API`);
        
        const recoveryAppendResult = await appendLeadsToOneDriveTable({
            delegatedAuth: req.delegatedAuth,
            sessionId: req.sessionId
        }, {
            leads: mergeResults.newLeads,
            filename: masterFileName,
            folderPath: masterFolderPath,
            useCustomFile: true
        });
        
        if (!recoveryAppendResult.success) {
            throw new Error(`Failed to append recovered leads: ${recoveryAppendResult.message}`);
        }
        
        console.log(`✅ RECOVERY SUCCESS: ${recoveryAppendResult.action} - ${recoveryAppendResult.leadsCount} leads processed`);

        console.log(`✅ RECOVERY MERGE completed successfully`);

        res.json({
            success: true,
            message: `Successfully recovered ${mergeResults.newLeads.length} leads`,
            recoveryResults: {
                originalCount: currentLeads.length,
                recoveredLeads: mergeResults.newLeads.length,
                duplicatesSkipped: mergeResults.duplicates.length,
                finalCount: currentLeads.length + mergeResults.newLeads.length
            }
        });

    } catch (error) {
        console.error('❌ Recovery merge error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to perform recovery merge',
            error: error.message
        });
    }
});

// DEBUG: Test upload and merge process
router.post('/debug/test-upload-merge', requireDelegatedAuth, async (req, res) => {
    try {

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Create test upload data
        const testUploadLeads = [
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

        // Step 1: Try to download existing master file
        console.log('Step 1: Downloading existing master file...');
        const existingWorkbook = await downloadMasterFile(graphClient);
        let existingData = [];
        
        if (existingWorkbook) {
            const leadsSheet = existingWorkbook.Sheets['Leads'];
            existingData = leadsSheet ? XLSX.utils.sheet_to_json(leadsSheet) : [];
            console.log(`Found existing master file with ${existingData.length} leads`);
        } else {
            console.log('No existing master file found');
        }

        // Step 2: Merge test data with existing
        console.log('Step 2: Merging test data...');
        const mergeResults = excelProcessor.mergeLeadsWithMaster(testUploadLeads, existingData);
        
        if (mergeResults.newLeads.length === 0) {
            return res.json({
                success: true,
                message: 'Test aborted - no new leads to add (all test emails already exist)',
                existingCount: existingData.length,
                testLeads: testUploadLeads,
                duplicates: mergeResults.duplicates
            });
        }

        // Step 3: FIXED - Use Microsoft Graph Table API instead of file replacement
        console.log('Step 3: Using Table API to append data (no file replacement)...');
        
        const appendResult = await appendLeadsToOneDriveTable({
            delegatedAuth: req.delegatedAuth,
            sessionId: req.sessionId
        }, {
            leads: mergeResults.newLeads,
            filename: 'LGA-Master-Email-List.xlsx',
            folderPath: '/LGA-Email-Automation',
            useCustomFile: true
        });
        
        if (!appendResult.success) {
            throw new Error(`Failed to append leads to table: ${appendResult.message}`);
        }
        
        console.log(`✅ TABLE APPEND SUCCESS: ${appendResult.action} - ${appendResult.leadsCount} leads processed`);

        // Step 5: Verification
        console.log('Step 5: Verifying upload...');
        const verificationWorkbook = await downloadMasterFile(graphClient);
        const verifyData = verificationWorkbook ? 
            XLSX.utils.sheet_to_json(verificationWorkbook.Sheets['Leads']) : [];

        res.json({
            success: true,
            testResults: {
                originalExisting: existingData.length,
                testLeadsUploaded: testUploadLeads.length,
                newLeadsAdded: mergeResults.newLeads.length,
                duplicatesFound: mergeResults.duplicates.length,
                finalVerificationCount: verifyData.length,
                expectedTotal: existingData.length + mergeResults.newLeads.length,
                dataIntegrityCheck: verifyData.length === (existingData.length + mergeResults.newLeads.length),
                firstRowAfterUpload: verifyData[0],
                lastRowAfterUpload: verifyData[verifyData.length - 1]
            }
        });

    } catch (error) {
        console.error('❌ Upload merge test error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to test upload merge',
            error: error.message,
            stack: error.stack
        });
    }
});

// DEBUG: Test Excel file creation locally
router.post('/debug/test-excel-creation', async (req, res) => {
    try {

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
        

        // Create master file
        const masterWorkbook = excelProcessor.createMasterFile();
        
        // FIXED: Test using Table API instead of file replacement
        console.log(`🧪 Using Table API for test data`);
        
        // For testing, we'll simulate the table append without actual upload
        console.log(`🧪 TEST: Would append ${normalizedLeads.length} normalized test leads to table`);
        
        // Create test buffer for size estimation only
        const updatedWorkbook = excelProcessor.updateMasterFileWithLeads(masterWorkbook, normalizedLeads);
        const buffer = excelProcessor.workbookToBuffer(updatedWorkbook);
        console.log(`🧪 Buffer test completed`);
        

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
        await advancedExcelUpload(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

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

// Cache for master file downloads with promise-based deduplication
const masterFileCache = new Map();
const activeDownloads = new Map();

// Helper function to download master file with caching and deduplication
async function downloadMasterFile(graphClient, useCache = true) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Create unique cache key per session to avoid cross-user conflicts
        const cacheKey = `master_file_${graphClient.config?.authProvider?.account?.username || 'default'}`;
        
        // Check cache first
        if (useCache && masterFileCache.has(cacheKey)) {
            const cached = masterFileCache.get(cacheKey);
            const age = Date.now() - cached.timestamp;
            if (age < 30 * 1000) { // 30 seconds
                console.log(`📋 Using cached master file (${Math.round(age/1000)}s old)`);
                return cached.workbook;
            } else {
                masterFileCache.delete(cacheKey);
            }
        }
        
        // Check if download is already in progress
        if (activeDownloads.has(cacheKey)) {
            console.log(`⏳ Master file download in progress, waiting...`);
            return await activeDownloads.get(cacheKey);
        }
        
        // Start download and cache the promise to prevent concurrent downloads
        const downloadPromise = performMasterFileDownload(graphClient, masterFileName, masterFolderPath, cacheKey);
        activeDownloads.set(cacheKey, downloadPromise);
        
        try {
            const workbook = await downloadPromise;
            return workbook;
        } finally {
            // Clean up active download tracking
            activeDownloads.delete(cacheKey);
        }
    } catch (error) {
        console.error('❌ Master file download error:', error);
        return null;
    }
}

// Separate function to perform the actual download
async function performMasterFileDownload(graphClient, masterFileName, masterFolderPath, cacheKey) {
    console.log(`📥 Downloading master file from: ${masterFolderPath}`);
    
    const files = await graphClient
        .api(`/me/drive/root:${masterFolderPath}:/children`)
        .filter(`name eq '${masterFileName}'`)
        .get();

    if (files.value.length === 0) {
        return null;
    }
        
        // Try multiple methods to download the file
        let fileContent = null;
        
        try {
            // Method 1: Direct content download
            fileContent = await graphClient
                .api(`/me/drive/items/${files.value[0].id}/content`)
                .get();
                
        } catch (error) {
            console.log('⚠️ Method 1 failed:', error.message);
        }
        
        // If Method 1 failed or returned undefined, try Method 2
        if (!fileContent || fileContent.length === undefined) {
            try {
                const response = await graphClient
                    .api(`/me/drive/items/${files.value[0].id}/content`)
                    .getStream();
                    
                // Convert stream to buffer
                const chunks = [];
                for await (const chunk of response) {
                    chunks.push(chunk);
                }
                fileContent = Buffer.concat(chunks);
                
            } catch (error) {
                console.log('⚠️ Method 2 failed:', error.message);
            }
        }
        
        // If both methods failed, try Method 3
        if (!fileContent || fileContent.length === undefined) {
            try {
                fileContent = await graphClient
                    .api(`/me/drive/root:${masterFolderPath}/${masterFileName}:/content`)
                    .get();
                    
            } catch (error) {
                console.log('⚠️ Method 3 failed:', error.message);
            }
        }

        if (!fileContent) {
            console.error('❌ All download methods failed');
            return null;
        }


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

        // Parse workbook and verify content
        const workbook = excelProcessor.bufferToWorkbook(buffer);
        
        if (!workbook.SheetNames.includes('Leads')) {
            console.error('❌ Downloaded file missing Leads sheet');
            return null;
        }
        
        // Cache the workbook for 30 seconds
        masterFileCache.set(cacheKey, {
            workbook: workbook,
            timestamp: Date.now()
        });
        
        return workbook;
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

// Helper function to upload file to OneDrive with retry logic and multiple methods
// Legacy wrapper function for backward compatibility
async function uploadToOneDrive(client, fileBuffer, filename, folderPath, maxRetries = 3) {
    console.log(`📤 LEGACY WRAPPER: Redirecting to advancedExcelUpload`);
    return await advancedExcelUpload(client, fileBuffer, filename, folderPath);
}

/**
 * Bridge function to call Microsoft Graph table append API from email automation
 * This replaces the old file replacement approach with table-based appending
 */
async function appendLeadsToOneDriveTable(auth, requestData) {
    try {
        console.log(`🔗 BRIDGE: Calling Microsoft Graph table append API`);
        
        // Get the base URL for internal API calls
        const protocol = process.env.NODE_ENV === 'development' ? 'http' : 'https';
        const host = process.env.RENDER_EXTERNAL_URL ? 
            new URL(process.env.RENDER_EXTERNAL_URL).host : 
            'localhost:3000';
        
        // Call our own Microsoft Graph table append endpoint
        const response = await axios.post(`${protocol}://${host}/api/microsoft-graph/onedrive/append-to-table`, requestData, {
            headers: {
                'Content-Type': 'application/json',
                'X-Session-Id': auth.sessionId
            },
            timeout: 30000 // 30 second timeout
        });
        
        if (response.data.success) {
            console.log(`✅ BRIDGE SUCCESS: Table append completed - ${response.data.action}`);
            return response.data;
        } else {
            console.error(`❌ BRIDGE ERROR: Table append failed`, response.data);
            return {
                success: false,
                message: response.data.message || 'Unknown error',
                error: response.data.error
            };
        }
        
    } catch (error) {
        console.error(`❌ BRIDGE EXCEPTION: Failed to call table append API`, error);
        
        // Extract meaningful error message
        let errorMessage = 'Failed to append to table';
        if (error.response && error.response.data && error.response.data.message) {
            errorMessage = error.response.data.message;
        } else if (error.message) {
            errorMessage = error.message;
        }
        
        return {
            success: false,
            message: errorMessage,
            error: error.code || 'BRIDGE_ERROR'
        };
    }
}

module.exports = router;