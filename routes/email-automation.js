const express = require('express');
const { requireGraphAuth } = require('../middleware/graphAuth');
const router = express.Router();

// Apply Graph authentication middleware to all routes
router.use(requireGraphAuth);

// Global email tracking storage (in production, use database)
global.emailTracking = global.emailTracking || new Map();

/**
 * Email Automation System with Read Receipt Tracking
 * Send emails through Microsoft Graph and track delivery/read status
 */

// Send email campaign
router.post('/send-campaign', async (req, res) => {
    try {
        const { 
            leads, 
            emailTemplate, 
            subject, 
            fromName = 'Lead Generation Team',
            trackReads = true,
            oneDriveFileId = null
        } = req.body;

        if (!leads || !Array.isArray(leads) || leads.length === 0) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Leads array is required and must not be empty'
            });
        }

        if (!emailTemplate || !subject) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Email template and subject are required'
            });
        }

        console.log(`📧 Starting email campaign for ${leads.length} leads...`);

        const campaignId = Date.now().toString() + Math.random().toString(36).substr(2, 9);
        const results = [];
        const errors = [];

        // Process leads in batches to respect Microsoft Graph rate limits
        const batchSize = 10;
        for (let i = 0; i < leads.length; i += batchSize) {
            const batch = leads.slice(i, i + batchSize);
            
            const batchPromises = batch.map(async (lead, index) => {
                try {
                    const globalIndex = i + index;
                    console.log(`📨 Sending email ${globalIndex + 1}/${leads.length} to ${lead.name} (${lead.email})`);

                    const emailResult = await sendPersonalizedEmail(
                        req.graphClient,
                        lead,
                        emailTemplate,
                        subject,
                        fromName,
                        trackReads,
                        campaignId
                    );

                    // Store tracking information
                    const trackingId = `${campaignId}-${lead.email}`;
                    global.emailTracking.set(trackingId, {
                        campaignId,
                        leadEmail: lead.email,
                        leadName: lead.name,
                        messageId: emailResult.messageId,
                        status: 'sent',
                        sentAt: new Date().toISOString(),
                        readAt: null,
                        repliedAt: null,
                        oneDriveFileId: oneDriveFileId,
                        trackingEnabled: trackReads
                    });

                    // Update OneDrive Excel if file ID provided
                    if (oneDriveFileId) {
                        try {
                            await updateExcelTracking(req.graphClient, oneDriveFileId, lead.email, {
                                sent: true,
                                status: 'Sent',
                                sentDate: new Date().toLocaleString()
                            });
                        } catch (excelError) {
                            console.warn(`⚠️ Failed to update Excel for ${lead.email}:`, excelError.message);
                        }
                    }

                    return {
                        ...lead,
                        email_sent: true,
                        email_status: 'sent',
                        sent_at: new Date().toISOString(),
                        tracking_id: trackingId,
                        message_id: emailResult.messageId
                    };

                } catch (error) {
                    console.error(`❌ Failed to send email to ${lead.email}:`, error.message);
                    errors.push({
                        email: lead.email,
                        name: lead.name,
                        error: error.message
                    });
                    
                    return {
                        ...lead,
                        email_sent: false,
                        email_status: 'failed',
                        error: error.message
                    };
                }
            });

            const batchResults = await Promise.all(batchPromises);
            results.push(...batchResults);

            // Add delay between batches to respect rate limits
            if (i + batchSize < leads.length) {
                await new Promise(resolve => setTimeout(resolve, 2000)); // 2 second delay
            }
        }

        console.log(`✅ Email campaign completed. Sent: ${results.filter(r => r.email_sent).length}, Failed: ${errors.length}`);

        res.json({
            success: true,
            campaignId: campaignId,
            totalLeads: leads.length,
            sent: results.filter(r => r.email_sent).length,
            failed: errors.length,
            results: results,
            errors: errors.length > 0 ? errors : undefined,
            trackingEnabled: trackReads,
            oneDriveFileId: oneDriveFileId,
            completedAt: new Date().toISOString()
        });

    } catch (error) {
        console.error('Email campaign error:', error);
        res.status(500).json({
            error: 'Email Campaign Error',
            message: 'Failed to send email campaign',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Get email tracking status
router.get('/tracking/:campaignId', async (req, res) => {
    try {
        const { campaignId } = req.params;

        global.emailTracking = global.emailTracking || new Map();
        
        const campaignEmails = Array.from(global.emailTracking.values())
            .filter(tracking => tracking.campaignId === campaignId);

        if (campaignEmails.length === 0) {
            return res.status(404).json({
                error: 'Campaign Not Found',
                message: 'No tracking data found for this campaign'
            });
        }

        const summary = {
            totalEmails: campaignEmails.length,
            sent: campaignEmails.filter(e => e.status === 'sent').length,
            read: campaignEmails.filter(e => e.readAt !== null).length,
            replied: campaignEmails.filter(e => e.repliedAt !== null).length,
            pending: campaignEmails.filter(e => e.status === 'sent' && !e.readAt).length
        };

        res.json({
            success: true,
            campaignId: campaignId,
            summary: summary,
            emails: campaignEmails.map(email => ({
                email: email.leadEmail,
                name: email.leadName,
                status: email.status,
                sentAt: email.sentAt,
                readAt: email.readAt,
                repliedAt: email.repliedAt
            }))
        });

    } catch (error) {
        console.error('Email tracking error:', error);
        res.status(500).json({
            error: 'Tracking Error',
            message: 'Failed to retrieve tracking data'
        });
    }
});

// Webhook endpoint for email read receipts and replies
router.post('/webhook/notifications', async (req, res) => {
    try {
        const notifications = req.body.value || [];
        
        console.log(`📬 Received ${notifications.length} webhook notifications`);

        for (const notification of notifications) {
            try {
                await processEmailNotification(req.graphClient, notification);
            } catch (notificationError) {
                console.error('Error processing notification:', notificationError);
            }
        }

        // Always return 200 OK for webhook
        res.status(200).json({ success: true });

    } catch (error) {
        console.error('Webhook processing error:', error);
        // Still return 200 OK to avoid webhook retry loops
        res.status(200).json({ success: true, error: error.message });
    }
});

// Webhook validation endpoint
router.get('/webhook/notifications', (req, res) => {
    const validationToken = req.query.validationToken;
    
    if (validationToken) {
        console.log('📬 Webhook validation received');
        res.setHeader('Content-Type', 'text/plain');
        res.status(200).send(validationToken);
    } else {
        res.status(400).json({ error: 'Missing validation token' });
    }
});

// Create webhook subscription for email notifications
router.post('/webhook/subscribe', async (req, res) => {
    try {
        const webhookUrl = process.env.RENDER_EXTERNAL_URL 
            ? `${process.env.RENDER_EXTERNAL_URL}/api/email/webhook/notifications`
            : `${req.protocol}://${req.get('host')}/api/email/webhook/notifications`;

        console.log(`📬 Creating webhook subscription for: ${webhookUrl}`);

        const subscription = await req.graphClient
            .api('/subscriptions')
            .post({
                changeType: 'created,updated',
                notificationUrl: webhookUrl,
                resource: '/me/messages',
                expirationDateTime: new Date(Date.now() + 3600000).toISOString(), // 1 hour from now
                clientState: 'LGA-EmailTracking'
            });

        console.log(`✅ Webhook subscription created: ${subscription.id}`);

        res.json({
            success: true,
            subscriptionId: subscription.id,
            expirationDateTime: subscription.expirationDateTime,
            notificationUrl: webhookUrl,
            message: 'Webhook subscription created successfully'
        });

    } catch (error) {
        console.error('Webhook subscription error:', error);
        res.status(500).json({
            error: 'Subscription Error',
            message: 'Failed to create webhook subscription',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Helper function to send personalized email
async function sendPersonalizedEmail(client, lead, template, subject, fromName, trackReads, campaignId) {
    // Personalize email content
    const personalizedContent = template
        .replace(/\{name\}/g, lead.name || 'there')
        .replace(/\{company\}/g, lead.organization_name || 'your company')
        .replace(/\{title\}/g, lead.title || '')
        .replace(/\{industry\}/g, lead.industry || '');

    const personalizedSubject = subject
        .replace(/\{name\}/g, lead.name || 'there')
        .replace(/\{company\}/g, lead.organization_name || 'your company');

    // Add tracking pixel if read tracking is enabled
    const trackingPixel = trackReads 
        ? `<img src="${process.env.RENDER_EXTERNAL_URL}/api/email/track-read?id=${campaignId}-${lead.email}" width="1" height="1" style="display:none;" />`
        : '';

    const emailMessage = {
        subject: personalizedSubject,
        body: {
            contentType: 'HTML',
            content: personalizedContent + trackingPixel
        },
        toRecipients: [
            {
                emailAddress: {
                    address: lead.email,
                    name: lead.name
                }
            }
        ],
        from: {
            emailAddress: {
                name: fromName
            }
        }
    };

    // Send email
    await client.api('/me/sendMail').post({
        message: emailMessage,
        saveToSentItems: true
    });

    return {
        messageId: `${campaignId}-${lead.email}`, // Simplified message ID
        success: true
    };
}

// Helper function to process webhook notifications
async function processEmailNotification(client, notification) {
    try {
        // Get the message details
        const message = await client.api(`/me/messages/${notification.resourceData.id}`).get();
        
        // Check if this is related to our tracking
        const messageId = message.id;
        const conversationId = message.conversationId;
        
        // Find related tracking record
        const trackingRecord = Array.from(global.emailTracking.values())
            .find(tracking => 
                message.subject.includes(tracking.leadEmail) || 
                message.toRecipients.some(recipient => recipient.emailAddress.address === tracking.leadEmail)
            );

        if (trackingRecord) {
            let updated = false;
            
            // Check if message was read (hasBeenRead property)
            if (message.isRead && !trackingRecord.readAt) {
                trackingRecord.readAt = new Date().toISOString();
                trackingRecord.status = 'read';
                updated = true;
                console.log(`📖 Email read by ${trackingRecord.leadEmail}`);
            }
            
            // Check if this is a reply
            if (message.sender && message.sender.emailAddress.address === trackingRecord.leadEmail) {
                trackingRecord.repliedAt = new Date().toISOString();
                trackingRecord.status = 'replied';
                updated = true;
                console.log(`↩️ Email reply from ${trackingRecord.leadEmail}`);
            }
            
            // Update Excel file if tracking record has OneDrive file ID
            if (updated && trackingRecord.oneDriveFileId) {
                try {
                    await updateExcelTracking(client, trackingRecord.oneDriveFileId, trackingRecord.leadEmail, {
                        sent: true,
                        status: trackingRecord.status,
                        sentDate: trackingRecord.sentAt ? new Date(trackingRecord.sentAt).toLocaleString() : '',
                        readDate: trackingRecord.readAt ? new Date(trackingRecord.readAt).toLocaleString() : '',
                        replyDate: trackingRecord.repliedAt ? new Date(trackingRecord.repliedAt).toLocaleString() : ''
                    });
                } catch (excelError) {
                    console.warn(`⚠️ Failed to update Excel tracking:`, excelError.message);
                }
            }
        }
        
    } catch (error) {
        console.error('Error processing email notification:', error);
    }
}

// Helper function to update Excel tracking (referenced from microsoft-graph.js)
async function updateExcelTracking(client, fileId, leadEmail, trackingData) {
    // This would make a call to the Excel update endpoint
    // For now, we'll implement a simplified version
    try {
        const response = await fetch(`${process.env.RENDER_EXTERNAL_URL}/api/microsoft-graph/onedrive/update-excel-tracking`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                fileId,
                leadEmail,
                trackingData
            })
        });
        
        if (!response.ok) {
            throw new Error(`Excel update failed: ${response.statusText}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error('Excel tracking update error:', error);
        throw error;
    }
}

// Pixel tracking endpoint for email read receipts
router.get('/track-read', (req, res) => {
    try {
        const { id } = req.query;
        
        if (id && global.emailTracking.has(id)) {
            const tracking = global.emailTracking.get(id);
            
            if (!tracking.readAt) {
                tracking.readAt = new Date().toISOString();
                tracking.status = 'read';
                console.log(`📖 Email read tracked: ${tracking.leadEmail}`);
                
                // Update Excel file asynchronously
                if (tracking.oneDriveFileId) {
                    updateExcelTracking(null, tracking.oneDriveFileId, tracking.leadEmail, {
                        sent: true,
                        status: 'read',
                        sentDate: tracking.sentAt ? new Date(tracking.sentAt).toLocaleString() : '',
                        readDate: new Date(tracking.readAt).toLocaleString()
                    }).catch(error => {
                        console.warn('Failed to update Excel on read tracking:', error);
                    });
                }
            }
        }
        
        // Return 1x1 transparent pixel
        const pixel = Buffer.from([
            0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00,
            0xff, 0xff, 0xff, 0x00, 0x00, 0x00, 0x21, 0xF9, 0x04, 0x01, 0x00, 0x00, 0x00,
            0x00, 0x2C, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02,
            0x44, 0x01, 0x00, 0x3B
        ]);
        
        res.setHeader('Content-Type', 'image/gif');
        res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
        res.send(pixel);
        
    } catch (error) {
        console.error('Read tracking error:', error);
        res.status(200).send(''); // Always return success to avoid breaking email display
    }
});

module.exports = router;