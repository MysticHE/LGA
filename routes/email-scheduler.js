const express = require('express');
const { requireDelegatedAuth } = require('../middleware/delegatedGraphAuth');
const EmailContentProcessor = require('../utils/emailContentProcessor');
const excelUpdateQueue = require('../utils/excelUpdateQueue');
const { getExcelColumnLetter, getLeadsViaGraphAPI, updateLeadViaGraphAPI } = require('../utils/excelGraphAPI');
const CampaignTokenManager = require('../utils/campaignTokenManager');
const router = express.Router();

// Initialize processors
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

        // Get leads data using Graph API
        const allLeads = await getLeadsViaGraphAPI(graphClient);
        
        if (!allLeads) {
            console.error(`❌ Master file data retrieval failed for campaign`);
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
        const leadsData = getTargetLeadsFromData(allLeads, targetLeads);
        
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

        // Get templates using Graph API
        const templates = await getTemplatesViaGraphAPI(graphClient);

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
                campaignId,
                followUpDays
            );
            
            emailsSent = sendResults.sent;
            results.push(...sendResults.results);
            errors.push(...sendResults.errors);

            // Skip bulk updates after campaign - immediate per-email updates are sufficient
            // to prevent duplicate Excel operations and rate limiting
            console.log('📊 Skipping bulk campaign updates - real-time updates already completed');
        } else {
            // Schedule emails for later
            emailsQueued = leadsData.length;
            
            // Create scheduled campaign record
            await createScheduledCampaign(graphClient, {
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
        await recordCampaignHistory(graphClient, {
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
        console.error('❌ Campaign start error:', error.message);
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

        // Get campaign history using Graph API
        const campaignHistory = await getCampaignHistoryViaGraphAPI(graphClient);
        const campaign = campaignHistory.find(c => c.Campaign_ID === campaignId);

        if (!campaign) {
            return res.status(404).json({
                success: false,
                message: 'Campaign not found'
            });
        }

        // Get leads associated with this campaign using Graph API
        const allLeads = await getLeadsViaGraphAPI(graphClient);
        const leadsData = getLeadsByCampaign(allLeads, campaignId);
        
        // Calculate campaign statistics
        const stats = calculateCampaignStats(leadsData);

        res.json({
            success: true,
            campaign: campaign,
            stats: stats,
            leadsCount: leadsData.length
        });

    } catch (error) {
        console.error('❌ Campaign status error:', error.message);
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

        // Get campaign history using Graph API
        let campaigns = await getCampaignHistoryViaGraphAPI(graphClient);
        
        if (!campaigns) {
            return res.json({
                success: true,
                campaigns: [],
                total: 0
            });
        }

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
        console.error('❌ Campaigns retrieval error:', error.message);
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

        // Update campaign status using Graph API
        const updated = await updateCampaignStatusViaGraphAPI(graphClient, campaignId, 'Paused');

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
        console.error('❌ Campaign pause error:', error.message);
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

        // Update campaign status using Graph API
        const updated = await updateCampaignStatusViaGraphAPI(graphClient, campaignId, 'Active');

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
        console.error('❌ Campaign resume error:', error.message);
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

        // Get scheduled campaigns that are due using Graph API
        const dueCampaigns = await getScheduledCampaignsDueViaGraphAPI(graphClient);
        
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

                // Get templates using Graph API
                const templates = await getTemplatesViaGraphAPI(graphClient);

                // Send emails
                const sendResults = await sendEmailsToLeads(
                    graphClient,
                    campaign.targetLeads,
                    campaign.emailContentType,
                    templates,
                    campaign.Campaign_ID,
                    campaign.followUpDays || 7
                );

                // Skip bulk updates - immediate per-email updates already completed
                // Update campaign status only
                await updateCampaignStatusViaGraphAPI(graphClient, campaign.Campaign_ID, 'Active');
                console.log('📊 Skipping scheduled campaign bulk updates - real-time updates sufficient');

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
        console.error('❌ Scheduled campaigns processing error:', error.message);
        res.status(500).json({
            success: false,
            message: 'Failed to process scheduled campaigns',
            error: error.message
        });
    }
});

// New function to get target leads from Graph API data (no workbook parsing needed)
function getTargetLeadsFromData(allLeads, targetCriteria) {
    console.log(`📊 Filtering ${allLeads.length} leads for criteria: ${targetCriteria}`);

    switch (targetCriteria) {
        case 'new':
            // Enhanced debugging for 'new' lead filtering
            const newLeadsFiltered = allLeads.filter(lead => {
                const isNew = lead.Status === 'New';
                
                console.log(`🔍 LEAD FILTER DEBUG - ${lead.Email}:`, {
                    Status: lead.Status,
                    isNew: isNew,
                    willInclude: isNew
                });
                
                return isNew;
            });
            
            
            return newLeadsFiltered;

        case 'due':
            const today = new Date().toISOString().split('T')[0];
            return allLeads.filter(lead => {
                const nextEmailDate = lead.Next_Email_Date ? 
                    new Date(lead.Next_Email_Date).toISOString().split('T')[0] : null;
                
                return nextEmailDate && nextEmailDate <= today && 
                       !['Replied', 'Unsubscribed', 'Bounced'].includes(lead.Status);
            });

        case 'all_new':
            const todayAllNew = new Date().toISOString().split('T')[0];
            return allLeads.filter(lead => {
                if (lead.Status === 'New') return true;
                
                const nextEmailDate = lead.Next_Email_Date ? 
                    new Date(lead.Next_Email_Date).toISOString().split('T')[0] : null;
                
                return nextEmailDate && nextEmailDate <= todayAllNew && 
                       !['Replied', 'Unsubscribed', 'Bounced'].includes(lead.Status);
            });

        default:
            return [];
    }
}

// Helper function to send emails to leads
async function sendEmailsToLeads(graphClient, leads, emailContentType, templates, campaignId, followUpDays = 7) {
    const results = [];
    const errors = [];
    let sent = 0;

    console.log(`📧 STARTING EMAIL SEND PROCESS:`);
    console.log(`   - Leads to process: ${leads.length}`);
    console.log(`   - Email content type: ${emailContentType}`);
    console.log(`   - Campaign ID: ${campaignId}`);
    console.log(`   - Templates available: ${templates.length}`);

    // Initialize campaign token manager for long campaigns
    const campaignTokenManager = new CampaignTokenManager();
    const estimatedDurationMs = leads.length * 60000; // Rough estimate: 1 minute per email
    console.log(`⏱️ Estimated campaign duration: ${Math.round(estimatedDurationMs / 60000)} minutes`);
    
    // Note: sessionId not available in this function - token management handled at higher level

    for (const lead of leads) {
        try {
            console.log(`🔄 Processing lead: ${lead.Email} (${lead.Name})`);
            
            // Process email content
            const emailContent = await emailContentProcessor.processEmailContent(
                lead, 
                emailContentType, 
                templates
            );
            
            console.log(`📝 Email content generated for ${lead.Email}`);

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

            // Send email using Microsoft Graph with token refresh handling
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

            console.log(`📧 Attempting to send email via Microsoft Graph to: ${lead.Email}`);

            // Handle token refresh for long campaigns
            let sendResult;
            try {
                sendResult = await graphClient.api('/me/sendMail').post({
                    message: emailMessage,
                    saveToSentItems: true
                });
            } catch (tokenError) {
                if (tokenError.message.includes('401') || tokenError.message.includes('unauthorized') || 
                    tokenError.message.includes('Authentication expired')) {
                    console.log('🔄 Token expired during campaign, attempting to refresh graph client...');
                    // Note: The calling function should handle token refresh at the session level
                    throw new Error('Token refresh required - campaign should be restarted');
                }
                throw tokenError;
            }
            
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

            // IMMEDIATE Excel update right after email is sent (for real-time tracking)
            console.log(`📊 Updating Excel for ${lead.Email} immediately...`);
            try {
                const updates = {
                    Status: 'Sent',
                    Campaign_Stage: 'Email_Sent',
                    Last_Email_Date: new Date().toISOString().split('T')[0],
                    Next_Email_Date: calculateNextEmailDate(new Date(), followUpDays || 7),
                    Email_Count: (lead.Email_Count || 0) + 1,
                    Template_Used: emailContent.contentType,
                    'Email Sent': 'Yes',
                    'Email Status': 'Sent',
                    'Email Bounce': 'No',
                    Campaign_ID: campaignId
                };

                await excelUpdateQueue.queueUpdate(
                    lead.Email,
                    () => updateLeadViaGraphAPI(graphClient, lead.Email, updates),
                    { 
                        type: 'campaign-send', 
                        email: lead.Email, 
                        source: 'email-scheduler',
                        priority: 'high'
                    }
                );
                console.log(`✅ Excel updated for ${lead.Email} - Status: Sent`);
            } catch (excelError) {
                console.error(`⚠️ Excel update failed for ${lead.Email}: ${excelError.message}`);
                // Continue campaign even if Excel update fails
            }

            console.log(`📊 Campaign progress: ${sent}/${leads.length} sent (${Math.round((sent / leads.length) * 100)}% complete)`);

            // Add progressive delay between emails (skip delay for last email)
            const leadIndex = leads.indexOf(lead);
            console.log(`🔍 DELAY DEBUG: leadIndex=${leadIndex}, totalLeads=${leads.length}, shouldDelay=${leadIndex < leads.length - 1}`);
            if (leadIndex < leads.length - 1) {
                const delaySeconds = Math.floor(Math.random() * (120 - 30 + 1)) + 30; // 30-120 seconds
                console.log(`⏳ Adding ${delaySeconds}s delay before next email...`);
                await new Promise(resolve => setTimeout(resolve, delaySeconds * 1000));
                console.log(`✅ Delay completed - ready for next email`);
            } else {
                console.log(`🏁 Last email - no delay needed`);
            }

        } catch (error) {
            console.error(`❌ Failed to send email to ${lead.Email}:`, error.message);
            console.error(`❌ Full error details:`, {
                code: error.code,
                statusCode: error.statusCode,
                message: error.message,
                stack: process.env.NODE_ENV === 'development' ? error.stack : 'Hidden in production'
            });

            // IMMEDIATE Excel update for failed emails (track attempt and failure reason)
            console.log(`📊 Updating Excel for ${lead.Email} - marking as failed...`);
            try {
                const failedUpdates = {
                    Status: 'Failed',
                    Last_Email_Date: new Date().toISOString().split('T')[0],
                    Email_Count: (lead.Email_Count || 0) + 1,
                    'Email Sent': 'No',
                    'Email Status': 'Failed',
                    'Email Bounce': 'No',
                    'Failed Date': new Date().toISOString(),
                    'Failure Reason': error.message?.substring(0, 255) || 'Unknown error',
                    Campaign_ID: campaignId
                };

                await excelUpdateQueue.queueUpdate(
                    lead.Email,
                    () => updateLeadViaGraphAPI(graphClient, lead.Email, failedUpdates),
                    { 
                        type: 'campaign-failed', 
                        email: lead.Email, 
                        source: 'email-scheduler',
                        priority: 'high'
                    }
                );
                console.log(`✅ Excel updated for ${lead.Email} - Status: Failed`);
            } catch (excelError) {
                console.error(`⚠️ Failed to update Excel for failed email ${lead.Email}: ${excelError.message}`);
            }

            errors.push({
                email: lead.Email,
                name: lead.Name,
                error: error.message,
                errorCode: error.code,
                statusCode: error.statusCode
            });

            console.log(`📊 Campaign progress: ${sent}/${leads.length} sent, ${errors.length} failed (${Math.round(((sent + errors.length) / leads.length) * 100)}% complete)`);
        }
    }

    console.log(`📊 EMAIL SEND SUMMARY:`);
    console.log(`   - Emails sent successfully: ${sent}`);
    console.log(`   - Errors encountered: ${errors.length}`);
    console.log(`   - Excel updates: Immediate per-email (real-time tracking enabled)`);
    if (errors.length > 0) {
        console.log(`   - Error details:`, errors.map(e => `${e.email}: ${e.error}`));
    }

    return { 
        sent, 
        results, 
        errors,
        excelUpdates: {
            realTimeUpdates: true,
            updateMethod: "immediate_per_email",
            queueingEnabled: true,
            totalProcessed: sent + errors.length,
            source: "email-scheduler"
        }
    };
}

// REMOVED: updateLeadsAfterCampaign function 
// This was causing duplicate Excel updates with the same data already updated in real-time
// Real-time per-email updates (lines 527-556) are sufficient and more accurate


function getLeadsByCampaign(allLeads, campaignId) {
    // This is a simplified approach - in a real system you'd track campaign associations
    // For now, we'll return leads that might have been part of the campaign
    return allLeads.filter(lead => 
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

async function recordCampaignHistory(graphClient, campaignData) {
    try {
        console.log(`📝 Recording campaign history: ${campaignData.Campaign_ID}`);
        
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Get Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            console.log('❌ Master file not found for campaign history recording');
            return false;
        }

        const fileId = files.value[0].id;
        
        try {
            // Try to append to Campaign_History worksheet
            // First get existing data to find next row
            const usedRange = await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('Campaign_History')/usedRange`)
                .get();
            
            let nextRow = 2; // Start after headers
            if (usedRange && usedRange.values) {
                nextRow = usedRange.values.length + 1;
            }
            
            // Add new campaign record
            const campaignRow = [
                campaignData.Campaign_ID,
                campaignData.Campaign_Name,
                campaignData.Start_Date,
                campaignData.Emails_Sent,
                campaignData.Emails_Read,
                campaignData.Replies,
                campaignData.Status
            ];
            
            await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('Campaign_History')/range(address='A${nextRow}:G${nextRow}')`)
                .patch({
                    values: [campaignRow]
                });
            
            console.log(`✅ Campaign history recorded at row ${nextRow}`);
            return true;
            
        } catch (error) {
            console.log('⚠️ Campaign_History sheet might not exist, creating headers...');
            // If sheet doesn't exist, create it with headers
            const headers = ['Campaign_ID', 'Campaign_Name', 'Start_Date', 'Emails_Sent', 'Emails_Read', 'Replies', 'Status'];
            
            await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('Campaign_History')/range(address='A1:G1')`)
                .patch({
                    values: [headers]
                });
            
            // Add campaign data
            const campaignRow = [
                campaignData.Campaign_ID,
                campaignData.Campaign_Name,
                campaignData.Start_Date,
                campaignData.Emails_Sent,
                campaignData.Emails_Read,
                campaignData.Replies,
                campaignData.Status
            ];
            
            await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('Campaign_History')/range(address='A2:G2')`)
                .patch({
                    values: [campaignRow]
                });
            
            console.log(`✅ Campaign history recorded with new headers`);
            return true;
        }
        
    } catch (error) {
        console.error('❌ Error recording campaign history:', error.message);
        // Don't throw error - campaign history is not critical for email sending
        return false;
    }
}


async function createScheduledCampaign(graphClient, campaignData) {
    // In a production system, you'd store scheduled campaigns in a separate sheet or database
    // For now, we'll add it to campaign history with scheduled status
    await recordCampaignHistory(graphClient, {
        Campaign_ID: campaignData.campaignId,
        Campaign_Name: campaignData.campaignName,
        Start_Date: campaignData.scheduledTime.split('T')[0],
        Emails_Sent: 0,
        Emails_Read: 0,
        Replies: 0,
        Status: 'Scheduled'
    });
}


// Helper function to get master file data using Graph Workbook API
async function getMasterFileData(graphClient, useCache = true) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        console.log(`📥 USING GRAPH WORKBOOK API (not raw download):`);
        console.log(`   - Searching for: ${masterFileName}`);
        console.log(`   - In folder: ${masterFolderPath}`);
        
        // First, get the file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();
            
        if (files.value.length === 0) {
            console.log(`❌ No files found matching '${masterFileName}'`);
            return null;
        }
        
        const fileId = files.value[0].id;
        console.log(`📄 Found file ID: ${fileId}`);
        
        // Get worksheets using Graph Workbook API
        const worksheets = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets`)
            .get();
            
        console.log(`📊 ACTUAL WORKSHEETS FROM GRAPH API:`);
        worksheets.value.forEach(sheet => {
            console.log(`   - Sheet: "${sheet.name}" (ID: ${sheet.id})`);
        });
        
        // Find the Leads sheet (should now show the real name)
        const leadsSheet = worksheets.value.find(sheet => 
            sheet.name === 'Leads' || sheet.name.toLowerCase().includes('lead')
        );
        
        if (!leadsSheet) {
            console.log(`❌ No Leads sheet found in worksheets`);
            return { worksheets: worksheets.value, leadsData: [] };
        }
        
        console.log(`✅ Found Leads sheet: "${leadsSheet.name}"`);
        
        // Get the actual data from the Leads sheet
        const tableData = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets('${leadsSheet.name}')/usedRange`)
            .get();
            
        // Convert Graph API table data to our expected format
        const leadsData = convertGraphTableToLeads(tableData);
        console.log(`📊 Leads data from Graph API: ${leadsData.length} leads`);
        
        return {
            worksheets: worksheets.value,
            leadsData: leadsData,
            fileId: fileId
        };
        
    } catch (error) {
        console.error('❌ Graph Workbook API error:', error.message);
        console.log('⚠️ Falling back to raw download method...');
        return await downloadMasterFileRaw(graphClient, useCache);
    }
}

// Convert Graph API table format to our lead format
function convertGraphTableToLeads(tableData) {
    if (!tableData || !tableData.values || tableData.values.length <= 1) {
        return [];
    }
    
    const headers = tableData.values[0];
    const rows = tableData.values.slice(1);
    
    return rows.map(row => {
        const lead = {};
        headers.forEach((header, index) => {
            lead[header] = row[index] || '';
        });
        return lead;
    }).filter(lead => lead.Email && lead.Email.trim()); // Only include leads with emails
}

// Fallback function for raw file download (original method)
async function downloadMasterFileRaw(graphClient, useCache = true) {
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

        // Debug the raw buffer to check for corruption or encoding issues
        if (Buffer.isBuffer(fileContent)) {
            const bufferStart = fileContent.slice(0, 100).toString('hex');
            console.log(`🔍 Buffer start (hex):`, bufferStart.substring(0, 40) + '...');
            
            // Check if it starts with Excel file signature
            const isExcelFile = bufferStart.startsWith('504b0304'); // ZIP signature for Excel
            console.log(`📊 Is valid Excel file signature:`, isExcelFile);
        }

        const workbook = excelProcessor.bufferToWorkbook(fileContent);
        
        console.log(`📊 WORKBOOK ANALYSIS:`);
        console.log(`   - Sheet names: [${Object.keys(workbook.Sheets).join(', ')}]`);
        
        // Enhanced debugging for sheet detection
        const sheetNames = Object.keys(workbook.Sheets);
        console.log(`🔍 DETAILED SHEET ANALYSIS:`);
        sheetNames.forEach(name => {
            const sheet = workbook.Sheets[name];
            const data = XLSX.utils.sheet_to_json(sheet);
            console.log(`   - Sheet "${name}": ${data.length} rows`);
            if (data.length > 0) {
                console.log(`     First row sample:`, Object.keys(data[0]).slice(0, 5));
            }
        });
        
        // Try multiple sheet names for lead data analysis (same logic as getTargetLeads)
        let analysisSheet = workbook.Sheets['Leads'] || 
                           workbook.Sheets['Sheet1'] || 
                           workbook.Sheets[Object.keys(workbook.Sheets)[0]];
        
        if (analysisSheet) {
            const sheetNameUsed = Object.keys(workbook.Sheets).find(name => 
                workbook.Sheets[name] === analysisSheet
            );
            const leadsData = XLSX.utils.sheet_to_json(analysisSheet);
            console.log(`   - Using sheet "${sheetNameUsed}" for analysis: ${leadsData.length} leads`);
            if (leadsData.length > 0) {
                console.log(`   - First lead sample:`, {
                    Name: leadsData[0].Name,
                    Email: leadsData[0].Email,
                    Status: leadsData[0].Status
                });
            }
        } else {
            console.log(`   - ❌ No lead sheet found in any format`);
        }
        
        return workbook;
    } catch (error) {
        console.error('❌ Master file download error:', error.message);
        console.error('❌ Full error details:', {
            message: error.message,
            code: error.code,
            statusCode: error.statusCode,
            stack: process.env.NODE_ENV === 'development' ? error.stack : 'Hidden'
        });
        return null;
    }
}


// REMOVED: updateLeadStatusViaGraph function
// This was a duplicate Excel update function that caused the same data to be updated twice
// The system now uses only updateLeadViaGraphAPI from utils/excelGraphAPI.js for all updates

// Helper function to calculate next email date
function calculateNextEmailDate(fromDate, followUpDays) {
    const nextDate = new Date(fromDate);
    nextDate.setDate(nextDate.getDate() + followUpDays);
    return nextDate.toISOString().split('T')[0];
}

// Helper function to get Excel column letter from number (A, B, C, ... Z, AA, AB, etc.)

// Graph API helper functions for migrated functionality


async function getTemplatesViaGraphAPI(graphClient) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Get Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            return [];
        }

        const fileId = files.value[0].id;
        
        try {
            // Try to get Templates worksheet data
            const usedRange = await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('Templates')/usedRange`)
                .get();
            
            if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
                return [];
            }
            
            // Convert to template objects
            const headers = usedRange.values[0];
            const rows = usedRange.values.slice(1);
            
            return rows.map(row => {
                const template = {};
                headers.forEach((header, index) => {
                    template[header] = row[index] || '';
                });
                return template;
            }).filter(template => template.Template_ID);
            
        } catch (error) {
            // If Templates sheet doesn't exist, return empty array
            console.log('Templates sheet not found, returning empty array');
            return [];
        }
        
    } catch (error) {
        console.error('❌ Get templates via Graph API error:', error.message);
        return [];
    }
}

async function getCampaignHistoryViaGraphAPI(graphClient) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Get Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            return null;
        }

        const fileId = files.value[0].id;
        
        try {
            // Try to get Campaign_History worksheet data
            const usedRange = await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('Campaign_History')/usedRange`)
                .get();
            
            if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
                return [];
            }
            
            // Convert to campaign objects
            const headers = usedRange.values[0];
            const rows = usedRange.values.slice(1);
            
            return rows.map(row => {
                const campaign = {};
                headers.forEach((header, index) => {
                    campaign[header] = row[index] || '';
                });
                return campaign;
            }).filter(campaign => campaign.Campaign_ID && campaign.Campaign_ID !== '');
            
        } catch (error) {
            // If Campaign_History sheet doesn't exist, return empty array
            console.log('Campaign_History sheet not found, returning empty array');
            return [];
        }
        
    } catch (error) {
        console.error('❌ Get campaign history via Graph API error:', error.message);
        return null;
    }
}

async function updateCampaignStatusViaGraphAPI(graphClient, campaignId, newStatus) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Get Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            console.log('❌ Master file not found for campaign status update');
            return false;
        }

        const fileId = files.value[0].id;
        
        try {
            // Get Campaign_History worksheet data
            const usedRange = await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('Campaign_History')/usedRange`)
                .get();
            
            if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
                console.log('❌ No campaign history data found');
                return false;
            }
            
            const headers = usedRange.values[0];
            const rows = usedRange.values.slice(1);
            
            // Find Campaign_ID column
            const campaignIdIndex = headers.findIndex(h => h === 'Campaign_ID');
            const statusIndex = headers.findIndex(h => h === 'Status');
            
            if (campaignIdIndex === -1 || statusIndex === -1) {
                console.log('❌ Campaign_ID or Status column not found');
                return false;
            }
            
            // Find the campaign row
            let targetRowIndex = -1;
            for (let i = 0; i < rows.length; i++) {
                if (rows[i][campaignIdIndex] === campaignId) {
                    targetRowIndex = i;
                    break;
                }
            }
            
            if (targetRowIndex === -1) {
                console.log(`❌ Campaign ${campaignId} not found`);
                return false;
            }
            
            // Update the status cell
            const excelRowNumber = targetRowIndex + 2; // +2 for header row and 0-based index
            const statusColumnLetter = getExcelColumnLetter(statusIndex + 1);
            const cellAddress = `${statusColumnLetter}${excelRowNumber}`;
            
            await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('Campaign_History')/range(address='${cellAddress}')`)
                .patch({
                    values: [[newStatus]]
                });
            
            console.log(`✅ Updated campaign ${campaignId} status to ${newStatus}`);
            return true;
            
        } catch (error) {
            console.error('❌ Error updating campaign status:', error.message);
            return false;
        }
        
    } catch (error) {
        console.error('❌ Update campaign status via Graph API error:', error.message);
        return false;
    }
}

async function getScheduledCampaignsDueViaGraphAPI(graphClient) {
    try {
        // Get campaign history to find scheduled campaigns
        const campaigns = await getCampaignHistoryViaGraphAPI(graphClient);
        
        if (!campaigns) {
            return [];
        }
        
        // Filter for scheduled campaigns that are due
        const now = new Date();
        const dueCampaigns = campaigns.filter(campaign => {
            if (campaign.Status !== 'Scheduled') return false;
            
            // Check if the scheduled time has passed
            const scheduledTime = new Date(campaign.Start_Date);
            return scheduledTime <= now;
        });
        
        // In a real system, you'd also need to retrieve the full campaign details
        // including target leads, email content type, etc. For now, return basic info
        return dueCampaigns.map(campaign => ({
            Campaign_ID: campaign.Campaign_ID,
            Campaign_Name: campaign.Campaign_Name,
            Start_Date: campaign.Start_Date,
            // These would need to be stored in the campaign record or a separate sheet
            targetLeads: [], // TODO: Retrieve from campaign details
            emailContentType: 'ai_generated', // TODO: Retrieve from campaign details
            followUpDays: 7 // TODO: Retrieve from campaign details
        }));
        
    } catch (error) {
        console.error('❌ Get scheduled campaigns due via Graph API error:', error.message);
        return [];
    }
}

module.exports = router;