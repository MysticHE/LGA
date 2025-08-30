const express = require('express');
const XLSX = require('xlsx');
const axios = require('axios');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const { advancedExcelUpload } = require('./excel-upload-fix');
// const persistentStorage = require('../utils/persistentStorage'); // Removed - using simplified Excel lookup
const router = express.Router();

// Initialize processors
const excelProcessor = new ExcelProcessor();
const authProvider = getDelegatedAuthProvider();

// In-memory webhook subscription storage (for production, use database)
const webhookSubscriptions = new Map();

// In-memory email-to-session mapping (for production, use database)
const emailSessionMapping = new Map();

/**
 * Email Tracking and Webhook Integration
 * Handles email read/reply status updates from Microsoft Graph webhooks
 */

// Webhook validation endpoint with enhanced logging
router.get('/webhook/notifications', async (req, res) => {
    const { validationToken } = req.query;
    
    if (validationToken) {
        console.log('📧 Webhook validation requested from Microsoft Graph');
        console.log(`🔑 Validation token: ${validationToken.substring(0, 20)}...`);
        
        // Validate webhook URL accessibility
        const webhookUrl = process.env.RENDER_EXTERNAL_URL || process.env.WEBHOOK_BASE_URL;
        console.log(`🌐 Webhook URL configured: ${webhookUrl}`);
        
        return res.status(200).send(validationToken);
    }
    
    console.log('❌ Webhook validation failed - no token provided');
    res.status(400).json({ error: 'No validation token provided' });
});

// Webhook notification endpoint with enhanced processing
router.post('/webhook/notifications', async (req, res) => {
    try {
        console.log('📧 Webhook notification received from Microsoft Graph');
        console.log('📋 Notification payload:', JSON.stringify(req.body, null, 2));
        
        // Validate client state for security
        const notifications = req.body.value || [];
        
        if (notifications.length === 0) {
            console.log('⚠️ No notifications in webhook payload');
            return res.status(202).send('Accepted');
        }
        
        let processedCount = 0;
        let errorCount = 0;
        
        for (const notification of notifications) {
            try {
                // Validate client state matches our expected value
                const expectedClientState = process.env.WEBHOOK_CLIENT_STATE || 'lga-email-tracking';
                if (notification.clientState && !notification.clientState.includes(expectedClientState.split('-')[0])) {
                    console.log(`⚠️ Ignoring notification with unexpected client state: ${notification.clientState}`);
                    continue;
                }
                
                await processEmailNotification(notification);
                processedCount++;
            } catch (notificationError) {
                console.error('❌ Error processing individual notification:', notificationError);
                errorCount++;
            }
        }
        
        console.log(`✅ Webhook processing completed: ${processedCount} processed, ${errorCount} errors`);
        res.status(202).send('Accepted');
        
    } catch (error) {
        console.error('❌ Webhook processing error:', error);
        res.status(500).json({ error: 'Failed to process webhook' });
    }
});

// Create webhook subscription with automatic management
router.post('/webhook/subscribe', requireDelegatedAuth, async (req, res) => {
    try {
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Check if webhook URL is properly configured
        const webhookUrl = process.env.RENDER_EXTERNAL_URL || process.env.WEBHOOK_BASE_URL;
        if (!webhookUrl) {
            return res.status(400).json({
                success: false,
                message: 'RENDER_EXTERNAL_URL or WEBHOOK_BASE_URL environment variable is required for webhooks',
                required: 'Set RENDER_EXTERNAL_URL to your deployed app URL (e.g., https://your-app.onrender.com)'
            });
        }
        
        const notificationUrl = `${webhookUrl}/api/email/webhook/notifications`;
        
        // Subscribe to email read/reply events
        const subscription = {
            changeType: 'updated',
            notificationUrl: notificationUrl,
            resource: '/me/messages',
            expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(), // 24 hours
            clientState: process.env.WEBHOOK_CLIENT_STATE || 'lga-email-tracking'
            // Note: Removed includeResourceData to fix Graph API validation error
        };
        
        console.log(`📡 Creating webhook subscription: ${notificationUrl}`);
        
        const result = await graphClient.api('/subscriptions').post(subscription);
        
        console.log('✅ Webhook subscription created:', result.id);
        
        // Store subscription ID for renewal
        await storeWebhookSubscription(req.sessionId, result.id, result.expirationDateTime);
        
        res.json({
            success: true,
            subscriptionId: result.id,
            expirationDateTime: result.expirationDateTime,
            notificationUrl: notificationUrl,
            resource: subscription.resource
        });
        
    } catch (error) {
        console.error('❌ Webhook subscription error:', error);
        
        let errorMessage = 'Failed to create webhook subscription';
        if (error.code === 'InvalidRequest' && error.message.includes('notificationUrl')) {
            errorMessage = 'Webhook URL validation failed. Ensure RENDER_EXTERNAL_URL is accessible and uses HTTPS.';
        }
        
        res.status(500).json({
            success: false,
            message: errorMessage,
            error: error.message,
            code: error.code
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

// Test endpoint to manually trigger read status update for debugging
router.post('/test-read-update', requireDelegatedAuth, async (req, res) => {
    try {
        const { email, testType } = req.body;
        
        if (!email) {
            return res.status(400).json({
                success: false,
                message: 'Email address is required'
            });
        }
        
        console.log(`🧪 TEST: Manually updating tracking for ${email} (${testType})`);
        
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Download master file and update using the same pattern as email automation
        const masterWorkbook = await downloadMasterFile(graphClient);
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        let updates;
        if (testType === 'read') {
            updates = {
                Status: 'Read',
                Read_Date: new Date().toISOString().split('T')[0],
                'Last Updated': new Date().toISOString()
            };
        } else if (testType === 'reply') {
            updates = {
                Status: 'Replied',
                Reply_Date: new Date().toISOString().split('T')[0],
                'Last Updated': new Date().toISOString()
            };
        }

        const updatedWorkbook = excelProcessor.updateLeadInMaster(masterWorkbook, email, updates);
        
        if (!updatedWorkbook) {
            return res.status(404).json({
                success: false,
                message: `Email ${email} not found in Excel file`
            });
        }

        // Upload updated file back to OneDrive
        const excelBuffer = XLSX.write(updatedWorkbook, { type: 'buffer', bookType: 'xlsx' });
        await advancedExcelUpload(graphClient, excelBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');
        
        res.json({
            success: true,
            message: `Test ${testType} update completed for ${email}`,
            testType: testType,
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        console.error('❌ Test tracking update error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to update test tracking',
            error: error.message
        });
    }
});

// Diagnostic endpoint to check master file tracking data
router.get('/diagnostic/:email?', requireDelegatedAuth, async (req, res) => {
    try {
        const { email } = req.params;
        
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        const masterWorkbook = await downloadMasterFile(graphClient, false);
        
        if (!masterWorkbook) {
            return res.json({
                success: false,
                message: 'No master file found'
            });
        }
        
        // Use intelligent sheet detection
        const sheetInfo = excelProcessor.findLeadsSheet(masterWorkbook);
        if (!sheetInfo) {
            return res.json({
                success: false,
                message: 'No valid lead data sheet found',
                availableSheets: Object.keys(masterWorkbook.Sheets)
            });
        }
        
        const leadsSheet = sheetInfo.sheet;
        const leadsData = XLSX.utils.sheet_to_json(leadsSheet);
        
        console.log(`📊 DIAGNOSTIC: Using sheet "${sheetInfo.name}" with ${leadsData.length} leads`);
        
        // If specific email requested, return that lead's tracking data
        if (email) {
            console.log(`🔍 DIAGNOSTIC: Searching for "${email}" in ${leadsData.length} leads`);
            console.log(`🔍 DIAGNOSTIC: Available columns:`, Object.keys(leadsData[0] || {}));
            console.log(`🔍 DIAGNOSTIC: First 5 emails:`, leadsData.slice(0, 5).map(l => ({
                Email: l.Email,
                email: l.email,
                allEmailFields: Object.keys(l).filter(k => k.toLowerCase().includes('email')).map(k => ({ [k]: l[k] }))
            })));
            
            // Try multiple email field matching strategies
            const lead = leadsData.find(l => {
                const searchEmail = email.toLowerCase().trim();
                const leadEmails = [
                    l.Email,
                    l.email,
                    l['Email Address'], 
                    l['email_address'],
                    l.EmailAddress,
                    l['Contact Email']
                ].filter(Boolean).map(e => String(e).toLowerCase().trim());
                
                return leadEmails.includes(searchEmail);
            });
            
            if (lead) {
                res.json({
                    success: true,
                    email: email,
                    trackingData: {
                        Status: lead.Status,
                        Last_Email_Date: lead.Last_Email_Date,
                        Read_Date: lead.Read_Date,
                        Reply_Date: lead.Reply_Date,
                        Email_Count: lead.Email_Count,
                        'Last Updated': lead['Last Updated']
                    },
                    fullLead: lead
                });
            } else {
                res.json({
                    success: false,
                    message: `Lead with email ${email} not found`,
                    totalLeads: leadsData.length
                });
            }
        } else {
            // Return summary of all tracking data
            const trackingSummary = {
                totalLeads: leadsData.length,
                statusCounts: {},
                trackingStats: {
                    withReadDate: 0,
                    withReplyDate: 0,
                    emailsSent: 0
                },
                recentActivity: []
            };
            
            leadsData.forEach(lead => {
                const status = lead.Status || 'Unknown';
                trackingSummary.statusCounts[status] = (trackingSummary.statusCounts[status] || 0) + 1;
                
                if (lead.Read_Date) trackingSummary.trackingStats.withReadDate++;
                if (lead.Reply_Date) trackingSummary.trackingStats.withReplyDate++;
                if (lead.Last_Email_Date || lead['Email Sent'] === 'Yes') trackingSummary.trackingStats.emailsSent++;
                
                // Add recent activity (last 7 days)
                if (lead['Last Updated']) {
                    const lastUpdated = new Date(lead['Last Updated']);
                    const sevenDaysAgo = new Date();
                    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
                    
                    if (lastUpdated >= sevenDaysAgo) {
                        trackingSummary.recentActivity.push({
                            email: lead.Email,
                            status: lead.Status,
                            lastUpdated: lead['Last Updated'],
                            readDate: lead.Read_Date,
                            replyDate: lead.Reply_Date
                        });
                    }
                }
            });
            
            // Sort recent activity by date
            trackingSummary.recentActivity.sort((a, b) => 
                new Date(b.lastUpdated) - new Date(a.lastUpdated)
            );
            
            res.json({
                success: true,
                diagnostic: trackingSummary
            });
        }
        
    } catch (error) {
        console.error('❌ Diagnostic error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve diagnostic data',
            error: error.message
        });
    }
});

// Webhook health check endpoint
router.get('/webhook/health', async (req, res) => {
    try {
        const webhookUrl = process.env.RENDER_EXTERNAL_URL || process.env.WEBHOOK_BASE_URL;
        const clientState = process.env.WEBHOOK_CLIENT_STATE || 'lga-email-tracking';
        
        const health = {
            webhook: {
                configured: !!webhookUrl,
                url: webhookUrl ? `${webhookUrl}/api/email/webhook/notifications` : null,
                clientState: clientState,
                accessible: false
            },
            environment: {
                nodeEnv: process.env.NODE_ENV || 'development',
                renderUrl: process.env.RENDER_EXTERNAL_URL || 'Not configured',
                webhookClientState: process.env.WEBHOOK_CLIENT_STATE || 'Default'
            },
            tracking: {
                pixelEndpoint: `/api/email/track-read`,
                webhookEndpoint: `/api/email/webhook/notifications`,
                testEndpoint: `/api/email/test-read-update`,
                diagnosticEndpoint: `/api/email/diagnostic`
            }
        };
        
        // Test webhook URL accessibility if configured
        if (webhookUrl) {
            try {
                const testUrl = `${webhookUrl}/health`;
                const response = await axios.get(testUrl, { timeout: 5000 });
                health.webhook.accessible = response.status === 200;
            } catch (accessError) {
                health.webhook.accessible = false;
                health.webhook.accessError = accessError.message;
            }
        }
        
        res.json({
            success: true,
            health: health,
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        console.error('❌ Webhook health check error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to check webhook health',
            error: error.message
        });
    }
});

// Webhook subscription management endpoints

// Get active webhook subscriptions
router.get('/webhook/subscriptions', requireDelegatedAuth, async (req, res) => {
    try {
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Get all subscriptions
        const subscriptions = await graphClient.api('/subscriptions').get();
        
        // Filter for our email tracking subscriptions
        const emailSubscriptions = subscriptions.value.filter(sub => 
            sub.clientState && sub.clientState.includes('lga-email-tracking')
        );
        
        res.json({
            success: true,
            subscriptions: emailSubscriptions,
            total: emailSubscriptions.length,
            stored: Array.from(webhookSubscriptions.values())
        });
        
    } catch (error) {
        console.error('❌ Webhook subscription list error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve webhook subscriptions',
            error: error.message
        });
    }
});

// Renew webhook subscription
router.post('/webhook/renew/:subscriptionId', requireDelegatedAuth, async (req, res) => {
    try {
        const { subscriptionId } = req.params;
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Extend expiration by 24 hours
        const newExpiration = new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString();
        
        const result = await graphClient.api(`/subscriptions/${subscriptionId}`).patch({
            expirationDateTime: newExpiration
        });
        
        console.log(`✅ Webhook subscription renewed: ${subscriptionId}`);
        
        // Update stored subscription
        await storeWebhookSubscription(req.sessionId, subscriptionId, newExpiration);
        
        res.json({
            success: true,
            subscriptionId: subscriptionId,
            newExpirationDateTime: result.expirationDateTime
        });
        
    } catch (error) {
        console.error('❌ Webhook renewal error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to renew webhook subscription',
            error: error.message
        });
    }
});

// Delete webhook subscription
router.delete('/webhook/subscriptions/:subscriptionId', requireDelegatedAuth, async (req, res) => {
    try {
        const { subscriptionId } = req.params;
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        await graphClient.api(`/subscriptions/${subscriptionId}`).delete();
        
        console.log(`✅ Webhook subscription deleted: ${subscriptionId}`);
        
        // Remove from storage
        webhookSubscriptions.delete(`${req.sessionId}-${subscriptionId}`);
        
        res.json({
            success: true,
            message: 'Webhook subscription deleted'
        });
        
    } catch (error) {
        console.error('❌ Webhook deletion error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to delete webhook subscription',
            error: error.message
        });
    }
});

// Auto-create webhook subscription when user logs in
router.post('/webhook/auto-setup', requireDelegatedAuth, async (req, res) => {
    try {
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Check if webhook URL is configured
        const webhookUrl = process.env.RENDER_EXTERNAL_URL || process.env.WEBHOOK_BASE_URL;
        if (!webhookUrl) {
            return res.status(400).json({
                success: false,
                message: 'Webhook URL not configured',
                required: 'RENDER_EXTERNAL_URL environment variable'
            });
        }
        
        // Check if user already has active subscriptions
        const existing = await graphClient.api('/subscriptions').get();
        const emailTrackingSubscriptions = existing.value.filter(sub => 
            sub.clientState && sub.clientState.includes('lga-email-tracking') &&
            new Date(sub.expirationDateTime) > new Date()
        );
        
        if (emailTrackingSubscriptions.length > 0) {
            console.log(`📡 Found ${emailTrackingSubscriptions.length} existing webhook subscriptions`);
            return res.json({
                success: true,
                message: 'Webhook subscriptions already active',
                subscriptions: emailTrackingSubscriptions
            });
        }
        
        // Create new subscription - fix Graph API requirements
        const subscription = {
            changeType: 'updated',
            notificationUrl: `${webhookUrl}/api/email/webhook/notifications`,
            resource: '/me/messages',
            expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
            clientState: `lga-email-tracking-${req.sessionId}`
            // Note: Removed includeResourceData to fix "select clause" error
        };
        
        const result = await graphClient.api('/subscriptions').post(subscription);
        
        console.log('✅ Auto-setup webhook subscription created:', result.id);
        
        // Store subscription for management
        await storeWebhookSubscription(req.sessionId, result.id, result.expirationDateTime);
        
        res.json({
            success: true,
            message: 'Webhook subscription auto-setup completed',
            subscriptionId: result.id,
            expirationDateTime: result.expirationDateTime
        });
        
    } catch (error) {
        console.error('❌ Webhook auto-setup error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to auto-setup webhook subscriptions',
            error: error.message
        });
    }
});

// Comprehensive tracking system status
router.get('/system-status', async (req, res) => {
    try {
        const activeSessions = authProvider.getActiveSessions();
        // Simplified tracking - no persistent storage needed
        const emailMappings = {};
        const webhookSubs = {};
        
        const status = {
            activeSessions: {
                count: activeSessions.length,
                sessions: activeSessions.map(sessionId => {
                    const userInfo = authProvider.getUserInfo(sessionId);
                    return {
                        sessionId: sessionId,
                        user: userInfo?.username || 'Unknown',
                        name: userInfo?.name || 'Unknown'
                    };
                })
            },
            emailMappings: {
                count: Object.keys(emailMappings).length,
                recent: Object.entries(emailMappings)
                    .sort((a, b) => new Date(b[1].createdAt) - new Date(a[1].createdAt))
                    .slice(0, 10)
                    .map(([email, mapping]) => ({
                        email: email,
                        sessionId: mapping.sessionId,
                        createdAt: mapping.createdAt,
                        active: activeSessions.includes(mapping.sessionId)
                    }))
            },
            webhookSubscriptions: {
                count: Object.keys(webhookSubs).length,
                subscriptions: Object.values(webhookSubs)
            },
            tracking: {
                pixelEndpoint: '/api/email/track-read',
                webhookEndpoint: '/api/email/webhook/notifications',
                persistentStorage: false,
                sessionRecovery: false,
                method: 'Direct Excel lookup'
            }
        };
        
        res.json({
            success: true,
            status: status,
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        console.error('❌ System status error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to get system status',
            error: error.message
        });
    }
});

// Register email-session mapping - SIMPLIFIED VERSION (no persistent storage needed)
router.post('/register-email-session', async (req, res) => {
    try {
        const { email, sessionId } = req.body;
        
        if (!email || !sessionId) {
            return res.status(400).json({
                success: false,
                message: 'Email and sessionId are required'
            });
        }
        
        // With simplified tracking, no persistent storage needed
        console.log(`📝 Email sent: ${email} from session ${sessionId} (tracking via Excel lookup)`);
        
        res.json({
            success: true,
            message: 'Email registration completed (using simplified tracking)',
            email: email,
            sessionId: sessionId,
            trackingMethod: 'Direct Excel lookup'
        });
        
    } catch (error) {
        console.error('❌ Email registration error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to register email',
            error: error.message
        });
    }
});

// Helper function to store webhook subscription info
async function storeWebhookSubscription(sessionId, subscriptionId, expirationDateTime) {
    const key = `${sessionId}-${subscriptionId}`;
    webhookSubscriptions.set(key, {
        sessionId: sessionId,
        subscriptionId: subscriptionId,
        expirationDateTime: expirationDateTime,
        createdAt: new Date().toISOString()
    });
    console.log(`📝 Stored webhook subscription: ${subscriptionId}`);
}

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
        
        console.log(`📧 Tracking pixel hit - updating read status for email: ${email}`);
        
        // Get all active sessions to try finding the right Excel file
        const activeSessions = authProvider.getActiveSessions();
        
        if (activeSessions.length === 0) {
            console.log('⚠️ No active sessions found, cannot update tracking');
            return;
        }
        
        console.log(`🔍 Searching for email ${email} across ${activeSessions.length} active sessions...`);
        
        // Try each active session to find the one containing this email
        let updateSuccess = false;
        for (const sessionId of activeSessions) {
            try {
                console.log(`🔍 Trying session: ${sessionId} for email ${email}`);
                const graphClient = await authProvider.getGraphClient(sessionId);
                
                // Download the master file and check if this email exists
                const masterWorkbook = await downloadMasterFile(graphClient);
                if (!masterWorkbook) {
                    console.log(`⚠️ No master file for session ${sessionId}`);
                    continue;
                }
                
                // Use the same update pattern as email automation - let excelProcessor handle the lookup
                const updates = {
                    Status: 'Read',
                    Read_Date: new Date().toISOString().split('T')[0],
                    'Last Updated': new Date().toISOString()
                };
                
                const updatedWorkbook = excelProcessor.updateLeadInMaster(masterWorkbook, email, updates);
                
                if (updatedWorkbook) {
                    // Upload updated file back to OneDrive
                    const excelBuffer = XLSX.write(updatedWorkbook, { type: 'buffer', bookType: 'xlsx' });
                    await advancedExcelUpload(graphClient, excelBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');
                    
                    console.log(`✅ Successfully updated read status for ${email} in session ${sessionId}`);
                    updateSuccess = true;
                    break;
                } else {
                    console.log(`❌ Email ${email} not found in session ${sessionId}`);
                }
                
            } catch (sessionError) {
                console.error(`❌ Session ${sessionId} failed:`, sessionError.message);
                continue;
            }
        }
        
        if (!updateSuccess) {
            console.error(`❌ Failed to update ${email} in any of ${activeSessions.length} sessions`);
        }
        
    } catch (error) {
        console.error('❌ Read status update error:', error);
    }
}


// Helper function to update master file with email status
async function updateMasterFileEmailStatus(emailId, subject, isRead, hasReply) {
    try {
        const activeSessions = authProvider.getActiveSessions();
        
        if (activeSessions.length === 0) {
            console.log('⚠️ No active sessions for webhook processing');
            return;
        }
        
        // Use first valid session
        const sessionId = activeSessions[0];
        const graphClient = await authProvider.getGraphClient(sessionId);
        
        // Enhanced lead matching strategy
        const masterWorkbook = await downloadMasterFile(graphClient, false); // Fresh download
        if (!masterWorkbook) return;
        
        // Use intelligent sheet detection
        const sheetInfo = excelProcessor.findLeadsSheet(masterWorkbook);
        if (!sheetInfo) {
            console.error(`❌ No valid lead data sheet found for email tracking update`);
            return;
        }
        
        const leadsSheet = sheetInfo.sheet;
        const leadsData = XLSX.utils.sheet_to_json(leadsSheet);
        
        console.log(`📧 TRACKING: Using sheet "${sheetInfo.name}" with ${leadsData.length} leads for email ${email}`);
        
        // Improved matching logic - try multiple approaches
        let matchingLead = null;
        
        // 1. Try matching by email ID from Graph API (best match)
        if (emailId) {
            matchingLead = leadsData.find(lead => 
                lead['Graph_Email_ID'] === emailId || 
                lead['Message_ID'] === emailId
            );
        }
        
        // 2. Try matching by subject containing lead email
        if (!matchingLead && subject) {
            matchingLead = leadsData.find(lead => {
                const leadEmail = lead.Email || '';
                if (!leadEmail) return false;
                
                // Check if subject contains the lead's email domain or name
                const emailDomain = leadEmail.split('@')[1];
                const emailName = leadEmail.split('@')[0];
                
                return subject.toLowerCase().includes(leadEmail.toLowerCase()) ||
                       subject.toLowerCase().includes(emailDomain.toLowerCase()) ||
                       (lead.Name && subject.toLowerCase().includes(lead.Name.toLowerCase())) ||
                       (lead['Company Name'] && subject.toLowerCase().includes(lead['Company Name'].toLowerCase()));
            });
        }
        
        // 3. Try matching recent sent emails (last 7 days)
        if (!matchingLead) {
            const sevenDaysAgo = new Date();
            sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
            
            matchingLead = leadsData.find(lead => {
                if (!lead.Last_Email_Date) return false;
                const lastEmailDate = new Date(lead.Last_Email_Date);
                return lastEmailDate >= sevenDaysAgo && lead.Status === 'Sent';
            });
        }
        
        if (matchingLead) {
            console.log(`📧 Found matching lead for webhook: ${matchingLead.Email}`);
            
            const updates = {
                'Last Updated': new Date().toISOString()
            };
            
            if (isRead && !matchingLead.Read_Date) {
                updates.Status = 'Read';
                updates.Read_Date = new Date().toISOString().split('T')[0];
                console.log(`📖 Setting read date for ${matchingLead.Email}`);
            }
            
            if (hasReply) {
                updates.Status = 'Replied';
                updates.Reply_Date = new Date().toISOString().split('T')[0];
                console.log(`💬 Setting reply date for ${matchingLead.Email}`);
            }
            
            if (Object.keys(updates).length > 1) { // More than just 'Last Updated'
                await updateLeadInMasterFile(graphClient, matchingLead.Email, updates);
            }
        } else {
            console.log(`⚠️ No matching lead found for webhook notification. Subject: ${subject?.substring(0, 50)}...`);
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