const express = require('express');
const { requireDelegatedAuth } = require('../middleware/delegatedGraphAuth');
const BounceDetector = require('../utils/bounceDetector');
const { getExcelColumnLetter, getLeadsViaGraphAPI, updateLeadViaGraphAPI } = require('../utils/excelGraphAPI');
const router = express.Router();

// Initialize bounce detector
const bounceDetector = new BounceDetector();

/**
 * Email Bounce Detection and Reporting Endpoints
 * Provides manual bounce checking and statistics
 */

// Manual bounce detection trigger
router.post('/check-bounces', requireDelegatedAuth, async (req, res) => {
    try {
        const { hoursBack = 24 } = req.body;
        
        console.log(`🔍 Manual bounce detection triggered - checking last ${hoursBack} hours`);
        
        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Check inbox for bounces
        const bounces = await bounceDetector.checkInboxForBounces(graphClient, hoursBack);
        
        if (bounces.length === 0) {
            return res.json({
                success: true,
                message: `No bounces found in the last ${hoursBack} hours`,
                bounces: [],
                processed: 0,
                statistics: {
                    hoursChecked: hoursBack,
                    bouncesFound: 0,
                    bouncesProcessed: 0
                }
            });
        }
        
        // Process detected bounces
        const results = await bounceDetector.processBounces(
            bounces,
            async (email, updates) => {
                // Use the Graph API update function from email automation
                return await updateLeadViaGraphAPI(graphClient, email, updates);
            }
        );
        
        console.log(`✅ Manual bounce check completed: ${results.bounced} emails marked as bounced`);
        
        res.json({
            success: true,
            message: `Bounce detection completed: ${results.bounced} emails marked as bounced`,
            bounces: bounces.map(bounce => ({
                email: bounce.originalRecipient,
                bounceReason: bounce.bounceReason,
                bounceDate: bounce.bounceDate,
                messageId: bounce.messageId
            })),
            processed: results.processed,
            statistics: {
                hoursChecked: hoursBack,
                bouncesFound: bounces.length,
                bouncesProcessed: results.bounced,
                errors: results.errors.length
            },
            errors: results.errors
        });
        
    } catch (error) {
        console.error('❌ Manual bounce detection error:', error);
        res.status(500).json({
            success: false,
            message: 'Bounce detection failed',
            error: error.message
        });
    }
});

// Get bounce statistics from master list
router.get('/statistics', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('📊 Retrieving bounce statistics...');
        
        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Get all leads data
        const leadsData = await getLeadsViaGraphAPI(graphClient);
        
        if (!leadsData) {
            return res.json({
                success: true,
                message: 'No master file found',
                statistics: {
                    totalLeads: 0,
                    bouncedEmails: 0,
                    validEmails: 0,
                    bounceRate: 0
                }
            });
        }
        
        // Calculate bounce statistics
        const statistics = bounceDetector.getBounceStatistics(leadsData);
        
        // Get additional insights
        const insights = {
            recentBounces: leadsData.filter(lead => 
                lead['Email Bounce'] === 'Yes' && 
                lead['Last Updated'] && 
                new Date(lead['Last Updated']) > new Date(Date.now() - 7 * 24 * 60 * 60 * 1000) // Last 7 days
            ).length,
            statusBreakdown: getStatusBreakdown(leadsData),
            recommendations: getBounceRecommendations(statistics)
        };
        
        res.json({
            success: true,
            statistics: statistics,
            insights: insights,
            lastUpdated: new Date().toISOString()
        });
        
    } catch (error) {
        console.error('❌ Bounce statistics error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve bounce statistics',
            error: error.message
        });
    }
});

// Get list of bounced emails
router.get('/bounced-emails', requireDelegatedAuth, async (req, res) => {
    try {
        const { limit = 50, offset = 0 } = req.query;
        
        console.log(`📋 Retrieving bounced emails (limit: ${limit}, offset: ${offset})`);
        
        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Get all leads data
        const leadsData = await getLeadsViaGraphAPI(graphClient);
        
        if (!leadsData) {
            return res.json({
                success: true,
                data: [],
                total: 0,
                message: 'No master file found'
            });
        }
        
        // Filter bounced emails
        const bouncedEmails = leadsData
            .filter(lead => lead['Email Bounce'] === 'Yes')
            .map(lead => ({
                name: lead.Name,
                email: lead.Email,
                company: lead['Company Name'],
                status: lead.Status,
                lastEmailDate: lead.Last_Email_Date,
                emailCount: lead.Email_Count || 0,
                lastUpdated: lead['Last Updated']
            }))
            .sort((a, b) => new Date(b.lastUpdated) - new Date(a.lastUpdated)); // Sort by most recent
        
        // Apply pagination
        const total = bouncedEmails.length;
        const paginatedData = bouncedEmails.slice(parseInt(offset), parseInt(offset) + parseInt(limit));
        
        res.json({
            success: true,
            data: paginatedData,
            total: total,
            limit: parseInt(limit),
            offset: parseInt(offset),
            hasMore: (parseInt(offset) + parseInt(limit)) < total
        });
        
    } catch (error) {
        console.error('❌ Bounced emails retrieval error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve bounced emails',
            error: error.message
        });
    }
});

// Mark email as not bounced (manual correction)
router.post('/mark-valid/:email', requireDelegatedAuth, async (req, res) => {
    try {
        const { email } = req.params;
        
        console.log(`✅ Manually marking ${email} as valid (not bounced)`);
        
        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);
        
        // Update lead status
        const updates = {
            'Email Bounce': 'No',
            'Status': 'Sent', // Reset to sent status
            'Last Updated': require('../utils/dateFormatter').getCurrentFormattedDate()
        };
        
        const success = await updateLeadViaGraphAPI(graphClient, email, updates);
        
        if (!success) {
            return res.status(404).json({
                success: false,
                message: 'Email not found in master list'
            });
        }
        
        console.log(`✅ ${email} marked as valid`);
        
        res.json({
            success: true,
            message: `${email} has been marked as valid (not bounced)`,
            email: email,
            updatedFields: Object.keys(updates)
        });
        
    } catch (error) {
        console.error('❌ Mark valid error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to mark email as valid',
            error: error.message
        });
    }
});

// Test bounce detection patterns
router.post('/test-patterns', (req, res) => {
    try {
        const { subject = '', sender = '', body = '' } = req.body;
        
        console.log('🧪 Testing bounce detection patterns...');
        
        // Create mock message
        const mockMessage = {
            subject: subject,
            from: { emailAddress: { address: sender } },
            body: { content: body },
            receivedDateTime: new Date().toISOString()
        };
        
        // Test bounce detection
        const bounceInfo = bounceDetector.detectBounce(mockMessage);
        
        res.json({
            success: true,
            message: 'Bounce pattern test completed',
            input: {
                subject: subject,
                sender: sender,
                bodyLength: body.length
            },
            result: {
                isBounce: bounceInfo !== null,
                bounceInfo: bounceInfo
            }
        });
        
    } catch (error) {
        console.error('❌ Pattern test error:', error);
        res.status(500).json({
            success: false,
            message: 'Pattern test failed',
            error: error.message
        });
    }
});

// Helper functions

// Get status breakdown for insights
function getStatusBreakdown(leadsData) {
    const breakdown = {};
    leadsData.forEach(lead => {
        const status = lead.Status || 'Unknown';
        breakdown[status] = (breakdown[status] || 0) + 1;
    });
    return breakdown;
}

// Get bounce recommendations based on statistics
function getBounceRecommendations(statistics) {
    const recommendations = [];
    
    if (statistics.bounceRate > 10) {
        recommendations.push('High bounce rate detected (>10%). Consider reviewing email list quality.');
    }
    
    if (statistics.bounceRate > 5) {
        recommendations.push('Bounce rate above 5%. Monitor sender reputation and consider list cleaning.');
    }
    
    if (statistics.bouncedEmails > 0) {
        recommendations.push('Remove bounced emails from future campaigns to maintain sender reputation.');
    }
    
    if (statistics.bounceRate <= 2) {
        recommendations.push('Excellent bounce rate! Your email list quality is good.');
    }
    
    return recommendations;
}



module.exports = router;