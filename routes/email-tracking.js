const express = require('express');
const XLSX = require('xlsx');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const { advancedExcelUpload } = require('./excel-upload-fix');
const router = express.Router();

// Initialize processors
const excelProcessor = new ExcelProcessor();
const authProvider = getDelegatedAuthProvider();

/**
 * Email Tracking and Webhook Integration
 * Handles email read/reply status updates from Microsoft Graph webhooks
 */

// Webhook validation endpoint
router.get('/webhook/notifications', async (req, res) => {
    const { validationToken } = req.query;
    
    if (validationToken) {
        console.log('📧 Webhook validation requested');
        return res.status(200).send(validationToken);
    }
    
    res.status(400).json({ error: 'No validation token provided' });
});

// Webhook notification endpoint
router.post('/webhook/notifications', async (req, res) => {
    try {
        console.log('📧 Webhook notification received:', JSON.stringify(req.body, null, 2));
        
        const notifications = req.body.value || [];
        
        for (const notification of notifications) {
            await processEmailNotification(notification);
        }
        
        res.status(202).send('Accepted');
        
    } catch (error) {
        console.error('❌ Webhook processing error:', error);
        res.status(500).json({ error: 'Failed to process webhook' });
    }
});

// Create webhook subscription
router.post('/webhook/subscribe', requireDelegatedAuth, async (req, res) => {
    try {
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Subscribe to email read/reply events
        const subscription = {
            changeType: 'updated',
            notificationUrl: `${process.env.RENDER_EXTERNAL_URL || 'http://localhost:3000'}/api/email/webhook/notifications`,
            resource: '/me/messages',
            expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(), // 24 hours
            clientState: 'lga-email-tracking'
        };
        
        const result = await graphClient.api('/subscriptions').post(subscription);
        
        console.log('✅ Webhook subscription created:', result.id);
        
        res.json({
            success: true,
            subscriptionId: result.id,
            expirationDateTime: result.expirationDateTime
        });
        
    } catch (error) {
        console.error('❌ Webhook subscription error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to create webhook subscription',
            error: error.message
        });
    }
});

// Pixel tracking endpoint
router.get('/track-read', async (req, res) => {
    try {
        const { id: trackingId } = req.query;
        
        if (trackingId) {
            console.log(`📧 Email read tracking pixel hit: ${trackingId}`);
            await updateEmailReadStatus(trackingId);
        }
        
        // Return 1x1 transparent pixel
        const pixelBuffer = Buffer.from(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==',
            'base64'
        );
        
        res.set({
            'Content-Type': 'image/png',
            'Content-Length': pixelBuffer.length,
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0'
        });
        
        res.send(pixelBuffer);
        
    } catch (error) {
        console.error('❌ Tracking pixel error:', error);
        res.status(500).send('Error');
    }
});

// Get campaign tracking stats
router.get('/tracking/:campaignId', requireDelegatedAuth, async (req, res) => {
    try {
        const { campaignId } = req.params;
        
        // Get tracking stats for campaign
        const stats = await getCampaignTrackingStats(campaignId);
        
        res.json({
            success: true,
            campaignId: campaignId,
            stats: stats
        });
        
    } catch (error) {
        console.error('❌ Campaign tracking error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to get campaign tracking stats',
            error: error.message
        });
    }
});

// Helper function to process email notifications
async function processEmailNotification(notification) {
    try {
        console.log(`📧 Processing notification for resource: ${notification.resource}`);
        
        // Get the email details from the notification
        const resourceData = notification.resourceData;
        
        if (!resourceData) {
            console.log('⚠️ No resource data in notification');
            return;
        }
        
        // Check if this is a read/reply event
        const isRead = resourceData.isRead === true;
        const hasReply = resourceData.parentFolderId && resourceData.parentFolderId.includes('SentItems');
        
        if (isRead || hasReply) {
            const emailId = resourceData.id;
            const subject = resourceData.subject || '';
            
            console.log(`📧 Email status change: ${subject} - Read: ${isRead}, Reply: ${hasReply}`);
            
            // Update master file with read/reply status
            await updateMasterFileEmailStatus(emailId, subject, isRead, hasReply);
        }
        
    } catch (error) {
        console.error('❌ Notification processing error:', error);
    }
}

// Helper function to update email read status from tracking pixel
async function updateEmailReadStatus(trackingId) {
    try {
        // Parse tracking ID to get email and campaign info
        const [email, timestamp] = trackingId.split('-');
        
        if (!email) {
            console.log('⚠️ Invalid tracking ID format');
            return;
        }
        
        console.log(`📧 Updating read status for email: ${email}`);
        
        // Get all active sessions to update across all users
        const activeSessions = authProvider.getActiveSessions();
        
        for (const sessionId of activeSessions) {
            try {
                const graphClient = await authProvider.getGraphClient(sessionId);
                await updateLeadEmailStatus(graphClient, email, 'Read', new Date().toISOString());
            } catch (error) {
                console.log(`⚠️ Failed to update for session ${sessionId}:`, error.message);
            }
        }
        
    } catch (error) {
        console.error('❌ Read status update error:', error);
    }
}

// Helper function to update master file with email status
async function updateMasterFileEmailStatus(emailId, subject, isRead, hasReply) {
    try {
        const activeSessions = authProvider.getActiveSessions();
        
        for (const sessionId of activeSessions) {
            try {
                const graphClient = await authProvider.getGraphClient(sessionId);
                
                // Download master file
                const masterWorkbook = await downloadMasterFile(graphClient);
                if (!masterWorkbook) continue;
                
                // Find lead by email content or subject matching
                const leadsSheet = masterWorkbook.Sheets['Leads'];
                const leadsData = XLSX.utils.sheet_to_json(leadsSheet);
                
                // Try to match by subject line containing email address or company name
                const matchingLead = leadsData.find(lead => {
                    const emailContent = lead.Email_Content_Sent || '';
                    const leadEmail = lead.Email || '';
                    return emailContent.includes(subject) || subject.includes(leadEmail) || 
                           (lead['Company Name'] && subject.includes(lead['Company Name']));
                });
                
                if (matchingLead) {
                    console.log(`📧 Found matching lead: ${matchingLead.Email}`);
                    
                    const updates = {};
                    if (isRead && !matchingLead.Read_Date) {
                        updates.Status = 'Read';
                        updates.Read_Date = new Date().toISOString().split('T')[0];
                    }
                    if (hasReply) {
                        updates.Status = 'Replied';
                        updates.Reply_Date = new Date().toISOString().split('T')[0];
                    }
                    
                    if (Object.keys(updates).length > 0) {
                        await updateLeadInMasterFile(graphClient, matchingLead.Email, updates);
                    }
                }
                
            } catch (error) {
                console.log(`⚠️ Failed to update master file for session ${sessionId}:`, error.message);
            }
        }
        
    } catch (error) {
        console.error('❌ Master file email status update error:', error);
    }
}

// Helper function to update specific lead email status
async function updateLeadEmailStatus(graphClient, email, status, date) {
    try {
        const updates = {
            Status: status,
            'Last Updated': new Date().toISOString()
        };
        
        if (status === 'Read') {
            updates.Read_Date = date.split('T')[0];
        } else if (status === 'Replied') {
            updates.Reply_Date = date.split('T')[0];
        }
        
        await updateLeadInMasterFile(graphClient, email, updates);
        
    } catch (error) {
        console.error('❌ Lead email status update error:', error);
    }
}

// Helper function to update lead in master file
async function updateLeadInMasterFile(graphClient, email, updates) {
    try {
        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        if (!masterWorkbook) return;
        
        // Update lead
        const updatedWorkbook = excelProcessor.updateLeadInMaster(masterWorkbook, email, updates);
        
        // Save updated file
        const masterBuffer = excelProcessor.workbookToBuffer(updatedWorkbook);
        await advancedExcelUpload(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');
        
        console.log(`✅ Updated lead ${email} with status: ${updates.Status}`);
        
    } catch (error) {
        console.error('❌ Master file lead update error:', error);
    }
}

// Helper function to download master file
async function downloadMasterFile(graphClient) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
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

// Helper function to get campaign tracking stats
async function getCampaignTrackingStats(campaignId) {
    // This would typically query a database for campaign-specific tracking
    // For now, return placeholder stats
    return {
        totalSent: 0,
        totalRead: 0,
        totalReplied: 0,
        readRate: 0,
        replyRate: 0
    };
}

module.exports = router;