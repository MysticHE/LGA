const express = require('express');
const XLSX = require('xlsx'); // Still needed for legacy diagnostic function - TODO: Remove when migrated to Graph API
const axios = require('axios');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const excelUpdateQueue = require('../utils/excelUpdateQueue');
const { updateLeadViaGraphAPI } = require('../utils/excelGraphAPI');
// const persistentStorage = require('../utils/persistentStorage'); // Removed - using simplified Excel lookup
const router = express.Router();

// Initialize processors
const excelProcessor = new ExcelProcessor();
const authProvider = getDelegatedAuthProvider();

/**
 * Email Tracking System
 * Handles email read status updates via tracking pixels and reply detection via cron jobs
 */


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

// Test endpoint to manually trigger reply detection for debugging
router.post('/test-reply-detection', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('🧪 TEST: Manually triggering reply detection...');
        
        const emailScheduler = require('../jobs/emailScheduler');
        
        // Trigger reply detection check
        await emailScheduler.checkInboxForReplies();
        
        res.json({
            success: true,
            message: 'Reply detection test completed',
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        console.error('❌ Test reply detection error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to test reply detection',
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
        
        // Use the new Graph API method for testing
        let updates;
        if (testType === 'read') {
            updates = {
                Status: 'Read',
                Read_Date: new Date().toISOString().split('T')[0],
                'Last Updated': require('../utils/dateFormatter').getCurrentFormattedDate()
            };
        } else if (testType === 'reply') {
            updates = {
                Status: 'Replied',
                Reply_Date: new Date().toISOString().split('T')[0],
                'Last Updated': require('../utils/dateFormatter').getCurrentFormattedDate()
            };
        }

        // Use direct Graph API update method
        const updateSuccess = await updateExcelViaGraphAPI(graphClient, email, updates);
        
        if (!updateSuccess) {
            return res.status(404).json({
                success: false,
                message: `Email ${email} not found in Excel file or update failed`
            });
        }
        
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
        // TODO: Migrate this diagnostic function to use Graph API directly like updateExcelViaGraphAPI
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
                if (lead.Last_Email_Date || lead.Status === 'Sent') trackingSummary.trackingStats.emailsSent++;
                
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


// Webhook subscription management endpoints


// Simplified tracking system status
router.get('/system-status', async (req, res) => {
    try {
        const activeSessions = authProvider.getActiveSessions();
        
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
            tracking: {
                pixelEndpoint: '/api/email/track-read',
                replyDetection: 'Cron job every 5 minutes',
                testEndpoint: '/api/email/test-reply-detection',
                diagnosticEndpoint: '/api/email/diagnostic',
                method: 'Direct Graph API updates'
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



// Fallback tracking storage for when no active sessions are available
const fs = require('fs').promises;
const path = require('path');

class TrackingFallbackManager {
    constructor() {
        this.fallbackDir = path.join(__dirname, '../tracking-fallback');
        this.ensureFallbackDirectory();
    }

    async ensureFallbackDirectory() {
        try {
            await fs.mkdir(this.fallbackDir, { recursive: true });
        } catch (error) {
            console.error('❌ Failed to create tracking fallback directory:', error.message);
        }
    }

    // Store tracking event when no sessions are available
    async storeTrackingEvent(email, eventType = 'read') {
        try {
            const event = {
                email: email,
                eventType: eventType,
                timestamp: new Date().toISOString(),
                date: new Date().toISOString().split('T')[0],
                processed: false
            };

            const filename = `tracking_${Date.now()}_${Math.random().toString(36).substr(2, 9)}.json`;
            const filepath = path.join(this.fallbackDir, filename);

            await fs.writeFile(filepath, JSON.stringify(event, null, 2));
            console.log(`📦 Tracking event stored for later processing: ${email} (${eventType})`);
            return true;
        } catch (error) {
            console.error('❌ Failed to store tracking event:', error.message);
            return false;
        }
    }

    // Process stored tracking events when sessions become available
    async processStoredEvents() {
        try {
            const files = await fs.readdir(this.fallbackDir);
            const trackingFiles = files.filter(file => file.startsWith('tracking_') && file.endsWith('.json'));

            if (trackingFiles.length === 0) {
                return { processed: 0, errors: 0 };
            }

            console.log(`📦 Processing ${trackingFiles.length} stored tracking events...`);

            let processed = 0;
            let errors = 0;

            for (const file of trackingFiles) {
                try {
                    const filepath = path.join(this.fallbackDir, file);
                    const data = await fs.readFile(filepath, 'utf8');
                    const event = JSON.parse(data);

                    if (!event.processed) {
                        // Try to process the event now
                        const success = await this.processTrackingEvent(event.email, event.eventType, event.date);
                        
                        if (success) {
                            // Mark as processed and delete file
                            await fs.unlink(filepath);
                            processed++;
                            console.log(`✅ Processed stored tracking event: ${event.email} (${event.eventType})`);
                        } else {
                            errors++;
                        }
                    }
                } catch (fileError) {
                    console.error(`❌ Error processing tracking file ${file}:`, fileError.message);
                    errors++;
                }
            }

            if (processed > 0) {
                console.log(`✅ Processed ${processed} stored tracking events, ${errors} errors`);
            }

            return { processed, errors };
        } catch (error) {
            console.error('❌ Error processing stored tracking events:', error.message);
            return { processed: 0, errors: 1 };
        }
    }

    // Try to process a single tracking event
    async processTrackingEvent(email, eventType, date) {
        try {
            const activeSessions = authProvider.getActiveSessions();
            if (activeSessions.length === 0) {
                return false; // Still no sessions available
            }

            for (const sessionId of activeSessions) {
                try {
                    const graphClient = await authProvider.getGraphClient(sessionId);
                    
                    const updates = {
                        Status: eventType === 'read' ? 'Read' : 'Clicked',
                        Read_Date: date,
                        'Last Updated': require('../utils/dateFormatter').getCurrentFormattedDate()
                    };

                    // Queue Excel update
                    const updateSuccess = await excelUpdateQueue.queueUpdate(
                        email,
                        () => updateLeadViaGraphAPI(graphClient, email, updates),
                        { 
                            type: 'fallback-tracking', 
                            email: email, 
                            source: 'tracking-pixel-fallback',
                            eventType: eventType
                        }
                    );
                    
                    if (updateSuccess) {
                        return true; // Successfully processed
                    }
                } catch (sessionError) {
                    continue; // Try next session
                }
            }

            return false; // Failed to process
        } catch (error) {
            console.error(`❌ Error processing tracking event for ${email}:`, error.message);
            return false;
        }
    }

    // Clean up old stored events (older than 7 days)
    async cleanupOldEvents() {
        try {
            const maxAge = 7 * 24 * 60 * 60 * 1000; // 7 days
            const now = new Date();
            const files = await fs.readdir(this.fallbackDir);

            let cleanedCount = 0;
            for (const file of files) {
                if (file.startsWith('tracking_') && file.endsWith('.json')) {
                    const filepath = path.join(this.fallbackDir, file);
                    const stats = await fs.stat(filepath);

                    if (now - stats.mtime > maxAge) {
                        await fs.unlink(filepath);
                        cleanedCount++;
                    }
                }
            }

            if (cleanedCount > 0) {
                console.log(`🧹 Cleaned up ${cleanedCount} old tracking fallback files`);
            }
        } catch (error) {
            console.error('❌ Error cleaning up old tracking events:', error.message);
        }
    }
}

const trackingFallback = new TrackingFallbackManager();

// Helper function to update email read status using direct Graph API Excel updates
async function updateEmailReadStatus(trackingId) {
    try {
        // Parse tracking ID - format: email-timestamp
        // Use lastIndexOf to handle emails with hyphens (e.g., user@domain-name.com)
        const lastDashIndex = trackingId.lastIndexOf('-');
        const email = lastDashIndex !== -1 ? trackingId.substring(0, lastDashIndex) : trackingId;
        const timestamp = lastDashIndex !== -1 ? trackingId.substring(lastDashIndex + 1) : '';
        
        if (!email) {
            console.log('❌ Invalid tracking ID format');
            return;
        }
        
        console.log(`📧 Tracking pixel hit - updating read status for email: ${email}`);
        
        // Get active sessions
        const activeSessions = authProvider.getActiveSessions();
        if (activeSessions.length === 0) {
            console.log('⚠️ No active sessions found, storing tracking event for later processing');
            
            // Store tracking event for later processing
            const stored = await trackingFallback.storeTrackingEvent(email, 'read');
            if (stored) {
                console.log(`📦 Tracking event stored for email: ${email}`);
            }
            return;
        }
        
        console.log(`🔍 Updating tracking for email: ${email} across ${activeSessions.length} sessions...`);
        
        // Try to process any stored events first (if sessions just became available)
        trackingFallback.processStoredEvents().catch(error => {
            console.error('❌ Error processing stored tracking events:', error.message);
        });
        
        // Try each active session
        let updateSuccess = false;
        for (const sessionId of activeSessions) {
            try {
                const graphClient = await authProvider.getGraphClient(sessionId);
                
                // Queue Excel update to prevent race conditions
                updateSuccess = await excelUpdateQueue.queueUpdate(
                    email, // Use email as file identifier
                    () => updateLeadViaGraphAPI(graphClient, email, {
                        Status: 'Read',
                        Read_Date: new Date().toISOString().split('T')[0],
                        'Last Updated': require('../utils/dateFormatter').getCurrentFormattedDate()
                    }),
                    { type: 'read-tracking', email: email, source: 'tracking-pixel' }
                );
                
                if (updateSuccess) {
                    console.log(`✅ Direct Graph API update successful for ${email} in session ${sessionId}`);
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
            // Store as fallback for retry later
            await trackingFallback.storeTrackingEvent(email, 'read');
        }
        
    } catch (error) {
        console.error('❌ Read status update error:', error);
    }
}

// Test endpoint: Manually process stored tracking events
router.get('/test-process-stored-events', async (req, res) => {
    try {
        console.log('🧪 TEST: Manually processing stored tracking events...');
        
        const results = await trackingFallback.processStoredEvents();
        
        res.json({
            success: true,
            message: 'Stored tracking events processing test completed',
            results: {
                processed: results.processed,
                errors: results.errors,
                timestamp: new Date().toISOString()
            }
        });
        
    } catch (error) {
        console.error('❌ Test stored events processing error:', error);
        res.status(500).json({
            success: false,
            error: 'Test failed',
            message: error.message
        });
    }
});

// Export both the router and the TrackingFallbackManager class
module.exports = router;
module.exports.TrackingFallbackManager = TrackingFallbackManager;
