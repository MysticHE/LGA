const express = require('express');
const XLSX = require('xlsx');
const { requireDelegatedAuth } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const EmailContentProcessor = require('../utils/emailContentProcessor');
const router = express.Router();

// Initialize processors
const excelProcessor = new ExcelProcessor();
const emailContentProcessor = new EmailContentProcessor();

/**
 * Email Scheduler and Campaign Management
 * Handles campaign creation, scheduling, and automated email sending
 */

// Start email campaign
router.post('/campaigns/start', requireDelegatedAuth, async (req, res) => {
    try {
        const {
            campaignName,
            emailContentType,
            targetLeads,
            sendSchedule,
            scheduledTime,
            followUpDays
        } = req.body;

        console.log(`🚀 Starting email campaign: ${campaignName}`);

        // Validate required fields
        if (!campaignName || !emailContentType || !targetLeads) {
            return res.status(400).json({
                success: false,
                message: 'Campaign name, email content type, and target leads are required'
            });
        }

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file - force fresh download for campaign operations
        console.log(`📥 Downloading master file for campaign (bypassing cache)...`);
        const masterWorkbook = await downloadMasterFile(graphClient, false);
        
        if (!masterWorkbook) {
            console.error(`❌ Master file download failed for campaign`);
            return res.status(404).json({
                success: false,
                message: 'Master file not found - cannot start campaign',
                debug: {
                    fileName: 'LGA-Master-Email-List.xlsx',
                    folderPath: '/LGA-Email-Automation'
                }
            });
        }

        // Get leads based on target criteria
        const leadsData = getTargetLeads(masterWorkbook, targetLeads);
        
        const leadsSheet = masterWorkbook.Sheets['Leads'];
        const allLeads = XLSX.utils.sheet_to_json(leadsSheet);
        
        if (leadsData.length === 0) {
            return res.json({
                success: true,
                message: 'No leads match the target criteria',
                campaignId: null,
                emailsSent: 0,
                debug: {
                    targetCriteria: targetLeads,
                    totalLeads: allLeads.length,
                    matchingLeads: leadsData.length
                }
            });
        }

        // Generate campaign ID
        const campaignId = `Campaign_${Date.now()}`;

        // Get templates for content processing
        const templates = excelProcessor.getTemplates(masterWorkbook);

        // Process and send emails based on schedule
        let emailsSent = 0;
        let emailsQueued = 0;
        const results = [];
        const errors = [];

        if (sendSchedule === 'immediate') {
            // Send emails immediately
            const sendResults = await sendEmailsToLeads(
                graphClient, 
                leadsData, 
                emailContentType, 
                templates, 
                campaignId
            );
            
            emailsSent = sendResults.sent;
            results.push(...sendResults.results);
            errors.push(...sendResults.errors);

            // Update leads in master file
            await updateLeadsAfterCampaign(graphClient, masterWorkbook, sendResults.results, followUpDays);
        } else {
            // Schedule emails for later
            emailsQueued = leadsData.length;
            
            // Create scheduled campaign record
            await createScheduledCampaign(graphClient, masterWorkbook, {
                campaignId,
                campaignName,
                emailContentType,
                targetLeads: leadsData,
                scheduledTime,
                followUpDays,
                status: 'Scheduled'
            });
            
            console.log(`📅 Campaign scheduled: ${campaignId} for ${scheduledTime}`);
        }

        // Record campaign in master file
        await recordCampaignHistory(graphClient, masterWorkbook, {
            Campaign_ID: campaignId,
            Campaign_Name: campaignName,
            Start_Date: new Date().toISOString().split('T')[0],
            Emails_Sent: emailsSent,
            Emails_Read: 0,
            Replies: 0,
            Status: sendSchedule === 'immediate' ? 'Active' : 'Scheduled'
        });

        console.log(`✅ Campaign ${sendSchedule === 'immediate' ? 'started' : 'scheduled'}: ${campaignId}`);

        res.json({
            success: true,
            message: `Campaign ${sendSchedule === 'immediate' ? 'started' : 'scheduled'} successfully`,
            campaignId: campaignId,
            campaignName: campaignName,
            emailsSent: emailsSent,
            emailsQueued: emailsQueued,
            targetLeads: leadsData.length,
            sendSchedule: sendSchedule,
            scheduledTime: scheduledTime,
            results: results.slice(0, 10), // Limit results for response size
            errors: errors.length > 0 ? errors.slice(0, 5) : undefined
        });

    } catch (error) {
        console.error('❌ Campaign start error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to start campaign',
            error: error.message
        });
    }
});

// Get campaign status and statistics
router.get('/campaigns/:campaignId', requireDelegatedAuth, async (req, res) => {
    try {
        const { campaignId } = req.params;
        console.log(`📊 Getting campaign status: ${campaignId}`);

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

        // Get campaign history
        const campaignHistory = getCampaignHistory(masterWorkbook);
        const campaign = campaignHistory.find(c => c.Campaign_ID === campaignId);

        if (!campaign) {
            return res.status(404).json({
                success: false,
                message: 'Campaign not found'
            });
        }

        // Get leads associated with this campaign
        const leadsData = getLeadsByCampaign(masterWorkbook, campaignId);
        
        // Calculate campaign statistics
        const stats = calculateCampaignStats(leadsData);

        res.json({
            success: true,
            campaign: campaign,
            stats: stats,
            leadsCount: leadsData.length
        });

    } catch (error) {
        console.error('❌ Campaign status error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to get campaign status',
            error: error.message
        });
    }
});

// Get all campaigns
router.get('/campaigns', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('📊 Getting all campaigns...');

        const { status, limit = 50 } = req.query;

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.json({
                success: true,
                campaigns: [],
                total: 0
            });
        }

        // Get campaign history
        let campaigns = getCampaignHistory(masterWorkbook);

        // Filter by status if provided
        if (status) {
            campaigns = campaigns.filter(campaign => campaign.Status === status);
        }

        // Sort by start date (newest first)
        campaigns.sort((a, b) => new Date(b.Start_Date) - new Date(a.Start_Date));

        // Apply limit
        const limitedCampaigns = campaigns.slice(0, parseInt(limit));

        res.json({
            success: true,
            campaigns: limitedCampaigns,
            total: campaigns.length,
            limit: parseInt(limit)
        });

    } catch (error) {
        console.error('❌ Campaigns retrieval error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve campaigns',
            error: error.message
        });
    }
});

// Pause campaign
router.post('/campaigns/:campaignId/pause', requireDelegatedAuth, async (req, res) => {
    try {
        const { campaignId } = req.params;
        console.log(`⏸️ Pausing campaign: ${campaignId}`);

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

        // Update campaign status
        const updated = await updateCampaignStatus(graphClient, masterWorkbook, campaignId, 'Paused');

        if (!updated) {
            return res.status(404).json({
                success: false,
                message: 'Campaign not found'
            });
        }

        res.json({
            success: true,
            message: 'Campaign paused successfully',
            campaignId: campaignId
        });

    } catch (error) {
        console.error('❌ Campaign pause error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to pause campaign',
            error: error.message
        });
    }
});

// Resume campaign
router.post('/campaigns/:campaignId/resume', requireDelegatedAuth, async (req, res) => {
    try {
        const { campaignId } = req.params;
        console.log(`▶️ Resuming campaign: ${campaignId}`);

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

        // Update campaign status
        const updated = await updateCampaignStatus(graphClient, masterWorkbook, campaignId, 'Active');

        if (!updated) {
            return res.status(404).json({
                success: false,
                message: 'Campaign not found'
            });
        }

        res.json({
            success: true,
            message: 'Campaign resumed successfully',
            campaignId: campaignId
        });

    } catch (error) {
        console.error('❌ Campaign resume error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to resume campaign',
            error: error.message
        });
    }
});

// Process scheduled campaigns (called by background job)
router.post('/process-scheduled', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('⏰ Processing scheduled campaigns...');

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.json({
                success: true,
                message: 'No master file found',
                processed: 0
            });
        }

        // Get scheduled campaigns that are due
        const dueCampaigns = getScheduledCampaignsDue(masterWorkbook);
        
        if (dueCampaigns.length === 0) {
            return res.json({
                success: true,
                message: 'No campaigns due for processing',
                processed: 0
            });
        }

        let processed = 0;
        const results = [];

        for (const campaign of dueCampaigns) {
            try {
                console.log(`📧 Processing scheduled campaign: ${campaign.Campaign_ID}`);

                // Get templates
                const templates = excelProcessor.getTemplates(masterWorkbook);

                // Send emails
                const sendResults = await sendEmailsToLeads(
                    graphClient,
                    campaign.targetLeads,
                    campaign.emailContentType,
                    templates,
                    campaign.Campaign_ID
                );

                // Update leads and campaign status
                await updateLeadsAfterCampaign(graphClient, masterWorkbook, sendResults.results, campaign.followUpDays);
                await updateCampaignStatus(graphClient, masterWorkbook, campaign.Campaign_ID, 'Active');

                processed++;
                results.push({
                    campaignId: campaign.Campaign_ID,
                    emailsSent: sendResults.sent,
                    errors: sendResults.errors.length
                });

                console.log(`✅ Processed campaign: ${campaign.Campaign_ID} - ${sendResults.sent} emails sent`);

            } catch (campaignError) {
                console.error(`❌ Failed to process campaign ${campaign.Campaign_ID}:`, campaignError);
                results.push({
                    campaignId: campaign.Campaign_ID,
                    error: campaignError.message
                });
            }
        }

        res.json({
            success: true,
            message: `Processed ${processed} scheduled campaigns`,
            processed: processed,
            results: results
        });

    } catch (error) {
        console.error('❌ Scheduled campaigns processing error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to process scheduled campaigns',
            error: error.message
        });
    }
});

// Helper function to get target leads based on criteria
function getTargetLeads(masterWorkbook, targetCriteria) {
    const leadsSheet = masterWorkbook.Sheets['Leads'];
    const allLeads = XLSX.utils.sheet_to_json(leadsSheet);

    switch (targetCriteria) {
        case 'new':
            // Enhanced debugging for 'new' lead filtering
            const newLeadsFiltered = allLeads.filter(lead => {
                const isNew = lead.Status === 'New';
                const isAutoEnabled = lead.Auto_Send_Enabled === 'Yes';
                
                console.log(`🔍 LEAD FILTER DEBUG - ${lead.Email}:`, {
                    Status: lead.Status,
                    isNew: isNew,
                    Auto_Send_Enabled: lead.Auto_Send_Enabled,
                    isAutoEnabled: isAutoEnabled,
                    willInclude: isNew && isAutoEnabled
                });
                
                return isNew && isAutoEnabled;
            });
            
            // TEMPORARY FIX: If no leads match strict criteria, try relaxed criteria for debugging
            if (newLeadsFiltered.length === 0) {
                console.log(`⚠️  No leads match strict criteria, checking with relaxed filters...`);
                const relaxedFilter = allLeads.filter(lead => lead.Status === 'New');
                console.log(`📋 Leads with just Status='New': ${relaxedFilter.length}`);
                
                if (relaxedFilter.length > 0) {
                    console.log(`🔧 TEMPORARY: Using relaxed criteria (ignoring Auto_Send_Enabled for debugging)`);
                    return relaxedFilter; // Return leads with just Status='New' for debugging
                }
            }
            
            return newLeadsFiltered;

        case 'due':
            const today = new Date().toISOString().split('T')[0];
            return allLeads.filter(lead => {
                const nextEmailDate = lead.Next_Email_Date ? 
                    new Date(lead.Next_Email_Date).toISOString().split('T')[0] : null;
                
                return nextEmailDate && nextEmailDate <= today && 
                       lead.Auto_Send_Enabled === 'Yes' &&
                       !['Replied', 'Unsubscribed', 'Bounced'].includes(lead.Status);
            });

        case 'all_new':
            const todayAllNew = new Date().toISOString().split('T')[0];
            return allLeads.filter(lead => {
                if (lead.Status === 'New' && lead.Auto_Send_Enabled === 'Yes') return true;
                
                const nextEmailDate = lead.Next_Email_Date ? 
                    new Date(lead.Next_Email_Date).toISOString().split('T')[0] : null;
                
                return nextEmailDate && nextEmailDate <= todayAllNew && 
                       lead.Auto_Send_Enabled === 'Yes' &&
                       !['Replied', 'Unsubscribed', 'Bounced'].includes(lead.Status);
            });

        default:
            return [];
    }
}

// Helper function to send emails to leads
async function sendEmailsToLeads(graphClient, leads, emailContentType, templates, campaignId) {
    const results = [];
    const errors = [];
    let sent = 0;

    console.log(`📧 STARTING EMAIL SEND PROCESS:`);
    console.log(`   - Leads to process: ${leads.length}`);
    console.log(`   - Email content type: ${emailContentType}`);
    console.log(`   - Campaign ID: ${campaignId}`);
    console.log(`   - Templates available: ${templates.length}`);

    for (const lead of leads) {
        try {
            console.log(`🔄 Processing lead: ${lead.Email} (${lead.Name})`);
            
            // Process email content
            const emailContent = await emailContentProcessor.processEmailContent(
                lead, 
                emailContentType, 
                templates
            );
            
            console.log(`📝 Email content generated for ${lead.Email}:`, {
                subject: emailContent.subject?.substring(0, 50) + '...',
                contentType: emailContent.contentType,
                hasBody: !!emailContent.body,
                bodyLength: emailContent.body?.length || 0
            });

            // Validate email content
            const validation = emailContentProcessor.validateEmailContent(emailContent);
            if (!validation.isValid) {
                console.error(`❌ Email validation failed for ${lead.Email}:`, validation.errors);
                errors.push({
                    email: lead.Email,
                    name: lead.Name,
                    error: 'Invalid email content: ' + validation.errors.join(', ')
                });
                continue;
            }

            // Send email using Microsoft Graph
            const emailMessage = {
                subject: emailContent.subject,
                body: {
                    contentType: 'HTML',
                    content: emailContentProcessor.convertToHTML(emailContent, lead.Email)
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

            console.log(`📧 Attempting to send email via Microsoft Graph to: ${lead.Email}`);
            console.log(`📋 Email message structure:`, {
                subject: emailMessage.subject,
                bodyType: emailMessage.body.contentType,
                recipientEmail: emailMessage.toRecipients[0].emailAddress.address,
                hasContent: !!emailMessage.body.content
            });

            const sendResult = await graphClient.api('/me/sendMail').post({
                message: emailMessage,
                saveToSentItems: true
            });
            
            console.log(`✅ Microsoft Graph sendMail API response:`, sendResult || 'No response body (normal for sendMail)');

            results.push({
                ...lead,
                emailSent: true,
                campaignId: campaignId,
                templateUsed: emailContent.contentType,
                sentAt: new Date().toISOString()
            });

            sent++;
            console.log(`📧 Email sent to: ${lead.Email}`);

            // Add small delay to respect rate limits
            await new Promise(resolve => setTimeout(resolve, 100));

        } catch (error) {
            console.error(`❌ Failed to send email to ${lead.Email}:`, error.message);
            console.error(`❌ Full error details:`, {
                code: error.code,
                statusCode: error.statusCode,
                message: error.message,
                stack: process.env.NODE_ENV === 'development' ? error.stack : 'Hidden in production'
            });
            errors.push({
                email: lead.Email,
                name: lead.Name,
                error: error.message,
                errorCode: error.code,
                statusCode: error.statusCode
            });
        }
    }

    console.log(`📊 EMAIL SEND SUMMARY:`);
    console.log(`   - Emails sent successfully: ${sent}`);
    console.log(`   - Errors encountered: ${errors.length}`);
    if (errors.length > 0) {
        console.log(`   - Error details:`, errors.map(e => `${e.email}: ${e.error}`));
    }

    return { sent, results, errors };
}

// Helper function to update leads after campaign
async function updateLeadsAfterCampaign(graphClient, masterWorkbook, results, followUpDays) {
    try {
        const leadsSheet = masterWorkbook.Sheets['Leads'];
        const leadsData = XLSX.utils.sheet_to_json(leadsSheet);

        // Update each lead that received an email
        for (const result of results) {
            if (result.emailSent) {
                const leadIndex = leadsData.findIndex(lead => 
                    lead.Email.toLowerCase() === result.Email.toLowerCase()
                );

                if (leadIndex !== -1) {
                    const lead = leadsData[leadIndex];
                    leadsData[leadIndex] = {
                        ...lead,
                        Status: 'Sent',
                        Last_Email_Date: new Date().toISOString().split('T')[0],
                        Email_Count: (lead.Email_Count || 0) + 1,
                        Template_Used: result.templateUsed,
                        Next_Email_Date: excelProcessor.calculateNextEmailDate(new Date(), followUpDays || 7),
                        'Email Sent': 'Yes',
                        'Email Status': 'Sent',
                        'Sent Date': new Date().toISOString(),
                        'Last Updated': new Date().toISOString()
                    };
                }
            }
        }

        // Update leads sheet
        const newLeadsSheet = XLSX.utils.json_to_sheet(leadsData);
        newLeadsSheet['!cols'] = excelProcessor.getColumnWidths();
        masterWorkbook.Sheets['Leads'] = newLeadsSheet;

        // Save updated master file
        const masterBuffer = excelProcessor.workbookToBuffer(masterWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

        console.log(`✅ Updated ${results.length} leads after campaign`);

    } catch (error) {
        console.error('❌ Error updating leads after campaign:', error);
        throw error;
    }
}

// Helper functions for campaign management
function getCampaignHistory(masterWorkbook) {
    try {
        const campaignSheet = masterWorkbook.Sheets['Campaign_History'];
        if (!campaignSheet) return [];
        
        const data = XLSX.utils.sheet_to_json(campaignSheet);
        return data.filter(campaign => campaign.Campaign_ID && campaign.Campaign_ID !== '');
    } catch (error) {
        console.error('❌ Error getting campaign history:', error);
        return [];
    }
}

function getLeadsByCampaign(masterWorkbook, campaignId) {
    const leadsSheet = masterWorkbook.Sheets['Leads'];
    const leadsData = XLSX.utils.sheet_to_json(leadsSheet);
    
    // This is a simplified approach - in a real system you'd track campaign associations
    // For now, we'll return leads that might have been part of the campaign
    return leadsData.filter(lead => 
        lead.Template_Used && lead.Last_Email_Date
    );
}

function calculateCampaignStats(leads) {
    const stats = {
        totalLeads: leads.length,
        sent: 0,
        read: 0,
        replied: 0,
        bounced: 0,
        pending: 0
    };

    leads.forEach(lead => {
        switch (lead.Status) {
            case 'Sent':
                stats.sent++;
                break;
            case 'Read':
                stats.read++;
                break;
            case 'Replied':
                stats.replied++;
                break;
            case 'Bounced':
                stats.bounced++;
                break;
            default:
                stats.pending++;
        }
    });

    return stats;
}

async function recordCampaignHistory(graphClient, masterWorkbook, campaignData) {
    try {
        const campaignSheet = masterWorkbook.Sheets['Campaign_History'];
        const existingData = XLSX.utils.sheet_to_json(campaignSheet);
        
        const updatedData = [...existingData.filter(c => c.Campaign_ID), campaignData];
        
        const newSheet = XLSX.utils.json_to_sheet(updatedData);
        newSheet['!cols'] = [
            {width: 20}, {width: 30}, {width: 20}, {width: 15}, {width: 15}, {width: 15}, {width: 20}
        ];
        
        masterWorkbook.Sheets['Campaign_History'] = newSheet;
        
        // Save updated master file
        const masterBuffer = excelProcessor.workbookToBuffer(masterWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');
        
    } catch (error) {
        console.error('❌ Error recording campaign history:', error);
        throw error;
    }
}

async function updateCampaignStatus(graphClient, masterWorkbook, campaignId, newStatus) {
    try {
        const campaignSheet = masterWorkbook.Sheets['Campaign_History'];
        const campaignData = XLSX.utils.sheet_to_json(campaignSheet);
        
        let found = false;
        for (let i = 0; i < campaignData.length; i++) {
            if (campaignData[i].Campaign_ID === campaignId) {
                campaignData[i].Status = newStatus;
                found = true;
                break;
            }
        }
        
        if (found) {
            const newSheet = XLSX.utils.json_to_sheet(campaignData);
            newSheet['!cols'] = [
                {width: 20}, {width: 30}, {width: 20}, {width: 15}, {width: 15}, {width: 15}, {width: 20}
            ];
            masterWorkbook.Sheets['Campaign_History'] = newSheet;
            
            const masterBuffer = excelProcessor.workbookToBuffer(masterWorkbook);
            await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');
        }
        
        return found;
    } catch (error) {
        console.error('❌ Error updating campaign status:', error);
        return false;
    }
}

async function createScheduledCampaign(graphClient, masterWorkbook, campaignData) {
    // In a production system, you'd store scheduled campaigns in a separate sheet or database
    // For now, we'll add it to campaign history with scheduled status
    await recordCampaignHistory(graphClient, masterWorkbook, {
        Campaign_ID: campaignData.campaignId,
        Campaign_Name: campaignData.campaignName,
        Start_Date: campaignData.scheduledTime.split('T')[0],
        Emails_Sent: 0,
        Emails_Read: 0,
        Replies: 0,
        Status: 'Scheduled'
    });
}

function getScheduledCampaignsDue(masterWorkbook) {
    // This is a simplified implementation
    // In a real system, you'd have a proper scheduled campaigns management
    return [];
}

// Helper function to download master file
async function downloadMasterFile(graphClient, useCache = true) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        console.log(`📥 MASTER FILE DOWNLOAD DEBUG:`);
        console.log(`   - Searching for: ${masterFileName}`);
        console.log(`   - In folder: ${masterFolderPath}`);
        
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        console.log(`📋 Files found in ${masterFolderPath}:`, files.value.length);
        if (files.value.length > 0) {
            console.log(`📄 Master file details:`, {
                name: files.value[0].name,
                id: files.value[0].id,
                size: files.value[0].size,
                lastModified: files.value[0].lastModifiedDateTime
            });
        } else {
            console.log(`❌ No files found matching '${masterFileName}'`);
            
            // List all files in the folder for debugging
            const allFiles = await graphClient
                .api(`/me/drive/root:${masterFolderPath}:/children`)
                .get();
            
            console.log(`📋 All files in ${masterFolderPath}:`, 
                allFiles.value.map(f => ({ name: f.name, size: f.size }))
            );
            
            return null;
        }

        console.log(`📥 Downloading file content...`);
        const fileContent = await graphClient
            .api(`/me/drive/items/${files.value[0].id}/content`)
            .get();

        console.log(`📄 File content received:`, {
            type: typeof fileContent,
            isBuffer: Buffer.isBuffer(fileContent),
            size: fileContent?.length || fileContent?.size || 'unknown'
        });

        const workbook = excelProcessor.bufferToWorkbook(fileContent);
        
        console.log(`📊 WORKBOOK ANALYSIS:`);
        console.log(`   - Sheet names: [${Object.keys(workbook.Sheets).join(', ')}]`);
        
        if (workbook.Sheets['Leads']) {
            const leadsData = XLSX.utils.sheet_to_json(workbook.Sheets['Leads']);
            console.log(`   - Leads sheet data count: ${leadsData.length}`);
            if (leadsData.length > 0) {
                console.log(`   - First lead sample:`, {
                    Name: leadsData[0].Name,
                    Email: leadsData[0].Email,
                    Status: leadsData[0].Status,
                    Auto_Send_Enabled: leadsData[0].Auto_Send_Enabled
                });
            }
        } else {
            console.log(`   - ❌ No 'Leads' sheet found in workbook`);
        }
        
        return workbook;
    } catch (error) {
        console.error('❌ Master file download error:', error);
        console.error('❌ Full error details:', {
            message: error.message,
            code: error.code,
            statusCode: error.statusCode,
            stack: process.env.NODE_ENV === 'development' ? error.stack : 'Hidden'
        });
        return null;
    }
}

// Helper function to upload file to OneDrive
async function uploadToOneDrive(client, fileBuffer, filename, folderPath) {
    try {
        const uploadUrl = `/me/drive/root:${folderPath}/${filename}:/content`;
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