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

        try {
            // Try to download existing master file
            const files = await graphClient
                .api(`/me/drive/root:${masterFolderPath}:/children`)
                .filter(`name eq '${masterFileName}'`)
                .get();

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
                }
            } else {
                console.log('📋 No master file found, creating new one...');
                masterWorkbook = excelProcessor.createMasterFile();
            }
        } catch (error) {
            if (error.code === 'itemNotFound' || error.message.includes('not found')) {
                console.log('📋 No existing master file found - creating new master file for first time');
            } else {
                console.log('📋 Creating new master file due to issue accessing existing file:', error.message);
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
        res.status(500).json({
            success: false,
            message: 'Failed to upload and process Excel file',
            error: error.message,
            details: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
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
        
        // Handle path correctly for OneDrive API
        const cleanPath = masterFolderPath.startsWith('/') ? masterFolderPath.substring(1) : masterFolderPath;
        
        const files = await graphClient
            .api(`/me/drive/root:/${cleanPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            console.log('📋 Master file not found');
            return null;
        }

        const fileContent = await graphClient
            .api(`/me/drive/items/${files.value[0].id}/content`)
            .get();

        return excelProcessor.bufferToWorkbook(fileContent);
    } catch (error) {
        console.error('❌ Master file download error:', error);
        return null;
    }
}

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

// Helper function to upload file to OneDrive
async function uploadToOneDrive(client, fileBuffer, filename, folderPath) {
    try {
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
        
        const result = await client.api(uploadUrl).put(fileBuffer);
        
        console.log(`📤 Uploaded file: ${filename} to ${folderPath}`);
        
        return {
            id: result.id,
            name: result.name,
            webUrl: result.webUrl,
            size: result.size
        };
    } catch (error) {
        console.error('❌ OneDrive upload error:', error);
        throw error;
    }
}

module.exports = router;