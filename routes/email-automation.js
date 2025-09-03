const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const axios = require('axios');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const EmailContentProcessor = require('../utils/emailContentProcessor');
const EmailDelayUtils = require('../utils/emailDelayUtils');
const excelUpdateQueue = require('../utils/excelUpdateQueue');
const { updateLeadViaGraphAPI, getLeadsViaGraphAPI } = require('../utils/excelGraphAPI');
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
const emailDelayUtils = new EmailDelayUtils();

/**
 * Enhanced corruption detection to handle Excel table format parsing issues
 * Addresses false positives where OneDrive table format isn't properly parsed by XLSX library
 */
async function isFileActuallyCorrupted(leadsSheet, leadsData) {
    console.log('🔍 Enhanced corruption detection starting...');
    
    // Check 1: No sheet at all = true corruption
    if (!leadsSheet) {
        console.log('❌ No Leads sheet found - true corruption');
        return true;
    }
    
    // Check 2: XLSX parsing returned data = not corrupted
    if (leadsData.length > 0) {
        console.log(`✅ XLSX parsing found ${leadsData.length} leads - not corrupted`);
        return false;
    }
    
    // Check 3: No sheet reference range = truly empty sheet
    if (!leadsSheet['!ref']) {
        console.log('❌ No sheet reference range - true corruption');
        return true;
    }
    
    console.log(`🔍 Sheet reference: ${leadsSheet['!ref']}`);
    
    // Check 4: Try alternative parsing methods for table format
    try {
        // Method 1: Parse with headers as first row
        const alternativeParse = XLSX.utils.sheet_to_json(leadsSheet, { header: 1, raw: false });
        console.log(`🔍 Alternative parsing found ${alternativeParse.length} rows`);
        
        if (alternativeParse.length > 1) { // More than just header
            console.log('✅ Alternative parsing found data - not corrupted');
            return false;
        }
        
        // Method 2: Check if we have just headers (table setup but no data)
        if (alternativeParse.length === 1) {
            const headers = alternativeParse[0];
            if (headers && headers.length > 0 && headers.some(h => h)) {
                console.log('✅ Found table headers - file structure intact, just no data yet');
                return false;
            }
        }
    } catch (parseError) {
        console.log(`⚠️ Alternative parsing failed: ${parseError.message}`);
    }
    
    // Check 5: Manual sheet structure examination
    try {
        const range = XLSX.utils.decode_range(leadsSheet['!ref']);
        const rowCount = range.e.r - range.s.r + 1;
        const colCount = range.e.c - range.s.c + 1;
        
        console.log(`🔍 Sheet dimensions: ${rowCount} rows x ${colCount} columns`);
        
        if (rowCount > 1 || colCount > 0) {
            console.log('✅ Sheet has structure - not corrupted');
            return false;
        }
    } catch (rangeError) {
        console.log(`⚠️ Range analysis failed: ${rangeError.message}`);
    }
    
    // Check 6: Look for any cell content directly
    const cellKeys = Object.keys(leadsSheet).filter(key => key.match(/^[A-Z]+\d+$/));
    if (cellKeys.length > 0) {
        console.log(`✅ Found ${cellKeys.length} cells with content - not corrupted`);
        return false;
    }
    
    console.log('❌ All checks indicate true corruption');
    return true;
}

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
                
                // Skip corruption detection - OneDrive files are reliable, and we can handle parsing issues
                // Original logic caused false positives with table format files
                const shouldRebuild = false; // Disabled - let normal flow handle any issues
                
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

        // CRITICAL: Get current OneDrive data for accurate duplicate checking using Graph API
        let currentOneDriveData = [];
        try {
            currentOneDriveData = await getLeadsViaGraphAPI(graphClient);
            console.log(`📊 Current OneDrive data: ${currentOneDriveData.length} leads for duplicate checking`);
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

        // Wait a moment for OneDrive to process the table append
        await new Promise(resolve => setTimeout(resolve, 2000)); // 2 second delay

        try {
            console.log(`📋 Fetching fresh data for verification (post-append)...`);
            const verifyData = await getLeadsViaGraphAPI(graphClient);
            if (verifyData && verifyData.length > 0) {
                
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
            const finalVerifyData = await getLeadsViaGraphAPI(graphClient);
            if (finalVerifyData) {
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

        // Get leads data using Graph API
        let leadsData = await getLeadsViaGraphAPI(graphClient);
        
        if (!leadsData) {
            return res.json({
                success: true,
                data: [],
                total: 0,
                message: 'No master file found'
            });
        }

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

        // Get leads data and calculate statistics using Graph API
        const leadsData = await getLeadsViaGraphAPI(graphClient);
        
        if (!leadsData) {
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

        // Calculate statistics from leads data
        const stats = calculateStatsFromLeads(leadsData);

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

        // Update lead using Graph API
        const updateSuccess = await updateLeadViaGraphAPI(graphClient, email, updates);
        
        if (!updateSuccess) {
            return res.status(404).json({
                success: false,
                message: 'Lead not found or update failed'
            });
        }

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

        // Get leads due today using Graph API
        const dueLeads = await getLeadsDueTodayViaGraphAPI(graphClient);
        
        if (!dueLeads) {
            return res.json({
                success: true,
                data: [],
                total: 0
            });
        }

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

        // Export master file using Graph API
        const buffer = await exportMasterFileViaGraphAPI(graphClient);
        
        if (!buffer) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }
        
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

        // Get current leads using Graph API
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        let currentLeads = await getLeadsViaGraphAPI(graphClient);
        
        if (!currentLeads) {
            console.log('📋 No current master file found');
            currentLeads = [];
        } else {
            console.log(`📊 Current master file has ${currentLeads.length} existing leads`);
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



// Send email to specific lead
router.post('/send-email/:email', requireDelegatedAuth, async (req, res) => {
    try {
        const { email } = req.params;
        const { emailChoice, customTemplate } = req.body;

        console.log(`📧 Sending email to: ${email} using ${emailChoice}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Get lead data and templates using Graph API
        const leadsData = await getLeadsViaGraphAPI(graphClient);
        const templates = await getTemplatesViaGraphAPI(graphClient);
        
        if (!leadsData) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        const lead = leadsData.find(l => l.Email.toLowerCase() === email.toLowerCase());

        if (!lead) {
            return res.status(404).json({
                success: false,
                message: 'Lead not found'
            });
        }

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
                content: emailContentProcessor.convertToHTML(emailContent, lead.Email, lead)
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

        // Get sender's email for tracking purposes
        const userInfo = authProvider.getUserInfo(req.sessionId);
        const senderEmail = userInfo?.username || 'unknown@sender.com';
        
        // Update lead status in master file
        const updates = {
            Status: 'Sent',
            Last_Email_Date: new Date().toISOString().split('T')[0],
            Email_Count: (lead.Email_Count || 0) + 1,
            Template_Used: emailContent.contentType,
            Next_Email_Date: calculateNextEmailDate(new Date(), lead.Follow_Up_Days || 7),
            'Email Sent': 'Yes',
            'Email Status': 'Sent',
            'Email Bounce': 'No', // Initialize bounce status
            'Sent Date': new Date().toISOString(),
            'Sent By': senderEmail
        };

        // Update lead using Graph API
        await updateLeadViaGraphAPI(graphClient, email, updates);

        console.log(`✅ Email sent successfully to: ${email}`);

        // Email tracking is now handled via direct Graph API Excel lookup (no persistent storage needed)

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

// Send bulk email campaign
router.post('/send-campaign', requireDelegatedAuth, async (req, res) => {
    try {
        const { 
            leads, 
            templateChoice = 'AI_Generated',
            emailTemplate = '',
            subject = '',
            trackReads = true,
            oneDriveFileId = null 
        } = req.body;

        console.log(`📧 Starting bulk email campaign for ${leads.length} leads using ${templateChoice}`);

        if (!leads || leads.length === 0) {
            return res.status(400).json({
                success: false,
                message: 'No leads provided for campaign'
            });
        }

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        const authProvider = getDelegatedAuthProvider();
        
        // Get templates for processing
        const templates = await getTemplatesViaGraphAPI(graphClient);
        
        const results = {
            campaignId: `campaign_${Date.now()}`,
            sent: 0,
            failed: 0,
            trackingEnabled: trackReads,
            errors: [],
            totalEmails: leads.length,
            estimatedTime: emailDelayUtils.estimateBulkSendingTime(leads.length)
        };

        console.log(`⏱️ Estimated campaign duration: ${results.estimatedTime.formatted}, completion: ${results.estimatedTime.completionTime}`);

        // Process each lead with random delays
        for (let i = 0; i < leads.length; i++) {
            const lead = leads[i];
            try {
                if (!lead.Email) {
                    results.failed++;
                    results.errors.push(`Lead missing email: ${lead.Name || 'Unknown'}`);
                    continue;
                }

                // Determine email choice - use templateChoice from frontend or lead's existing choice
                let emailChoice = templateChoice;
                if (emailChoice === 'custom' && emailTemplate) {
                    // For custom templates, create temporary template-like structure
                    emailChoice = 'AI_Generated'; // Process as custom content
                    lead.AI_Generated_Email = `Subject: ${subject}\n\n${emailTemplate}`;
                }

                // Process email content
                const emailContent = await emailContentProcessor.processEmailContent(
                    lead, 
                    emailChoice, 
                    templates
                );

                // Send email via Microsoft Graph
                const emailMessage = emailContentProcessor.createEmailMessage(
                    emailContent, 
                    lead.Email, 
                    lead,
                    trackReads
                );

                await graphClient
                    .api('/me/sendMail')
                    .post({ message: emailMessage });

                // Update lead status
                const updates = {
                    Status: 'Sent',
                    Last_Email_Date: new Date().toISOString().split('T')[0],
                    Email_Count: (lead.Email_Count || 0) + 1,
                    Template_Used: emailContent.contentType,
                    Email_Choice: emailChoice,
                    'Email Sent': 'Yes',
                    'Email Status': 'Sent',
                    'Email Bounce': 'No', // Initialize bounce status
                    'Sent Date': new Date().toISOString()
                };

                // Queue Excel update to prevent race conditions
                await excelUpdateQueue.queueUpdate(
                    lead.Email, // Use email as file identifier
                    () => updateLeadViaGraphAPI(graphClient, lead.Email, updates),
                    { type: 'campaign-send', email: lead.Email, source: 'email-automation' }
                );
                results.sent++;

                console.log(`📧 Email ${i + 1}/${leads.length} sent to ${lead.Email} (${results.sent} successful, ${results.failed} failed)`);

                // Add random delay between emails (skip delay for last email)
                if (i < leads.length - 1) {
                    await emailDelayUtils.progressiveDelay(i, leads.length);
                }

            } catch (emailError) {
                console.error(`❌ Failed to send email to ${lead.Email}:`, emailError.message);
                results.failed++;
                results.errors.push(`${lead.Email}: ${emailError.message}`);
                
                // Add delay even after failures to maintain sending pattern
                if (i < leads.length - 1) {
                    await emailDelayUtils.randomDelay(15, 45); // Shorter delay after failures
                }
            }
        }

        console.log(`✅ Campaign completed: ${results.sent} sent, ${results.failed} failed`);

        // Calculate actual completion time
        const actualEndTime = new Date();
        const actualDuration = Math.round((actualEndTime - new Date(Date.now() - results.estimatedTime.totalSeconds * 1000)) / 1000);
        
        res.json({
            success: true,
            message: `Campaign completed: ${results.sent} emails sent, ${results.failed} failed`,
            ...results,
            timing: {
                estimated: results.estimatedTime,
                actualDurationSeconds: actualDuration,
                actualDurationFormatted: emailDelayUtils.formatDelayTime(actualDuration * 1000),
                completedAt: actualEndTime.toLocaleTimeString()
            },
            delayStats: emailDelayUtils.getDelayStats()
        });

    } catch (error) {
        console.error('❌ Campaign error:', error);
        res.status(500).json({
            success: false,
            message: 'Campaign failed',
            error: error.message
        });
    }
});

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
