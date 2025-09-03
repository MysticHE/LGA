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
                'Last Updated': new Date().toISOString()
            };
        } else if (testType === 'reply') {
            updates = {
                Status: 'Replied',
                Reply_Date: new Date().toISOString().split('T')[0],
                'Last Updated': new Date().toISOString()
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



// Helper function to update email read status using direct Graph API Excel updates
async function updateEmailReadStatus(trackingId) {
    try {
        // Parse tracking ID - format: email-timestamp
        const [email, timestamp] = trackingId.split('-');
        
        if (!email) {
            console.log('❌ Invalid tracking ID format');
            return;
        }
        
        console.log(`📧 Tracking pixel hit - updating read status for email: ${email}`);
        
        // Get active sessions
        const activeSessions = authProvider.getActiveSessions();
        if (activeSessions.length === 0) {
            console.log('⚠️ No active sessions found, cannot update tracking');
            return;
        }
        
        console.log(`🔍 Updating tracking for email: ${email} across ${activeSessions.length} sessions...`);
        
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
                        'Last Updated': new Date().toISOString()
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
        }
        
    } catch (error) {
        console.error('❌ Read status update error:', error);
    }
}

module.exports = router;
