const express = require('express');
const EmailDelayUtils = require('../utils/emailDelayUtils');
const router = express.Router();

// Initialize delay utilities
const emailDelayUtils = new EmailDelayUtils();

/**
 * Email Delay Testing Endpoints
 * For testing and demonstrating the delay functionality
 */

// Test random delay functionality
router.post('/test-delay', async (req, res) => {
    try {
        const { min = 30, max = 120 } = req.body;
        
        console.log(`🧪 Testing random delay between ${min}-${max} seconds...`);
        const startTime = Date.now();
        
        const delayMs = await emailDelayUtils.randomDelay(min, max);
        
        const endTime = Date.now();
        const actualDelay = Math.round((endTime - startTime) / 1000);
        
        res.json({
            success: true,
            message: `Delay test completed`,
            delayRequested: `${min}-${max} seconds`,
            actualDelayMs: delayMs,
            actualDelaySeconds: actualDelay,
            actualDelayFormatted: emailDelayUtils.formatDelayTime(delayMs)
        });
        
    } catch (error) {
        console.error('❌ Delay test error:', error);
        res.status(500).json({
            success: false,
            message: 'Delay test failed',
            error: error.message
        });
    }
});

// Get delay statistics and configuration
router.get('/delay-stats', (req, res) => {
    try {
        const stats = emailDelayUtils.getDelayStats();
        
        res.json({
            success: true,
            delayConfiguration: stats,
            features: {
                randomDelay: 'Basic random delay between min-max range',
                progressiveDelay: 'Increases delay over time during bulk sending',
                smartDelay: 'Adjusts delay based on time of day and volume',
                batchDelay: 'Extended delays between email batches'
            },
            usage: {
                bulkCampaigns: 'Progressive delays for natural sending patterns',
                scheduledEmails: 'Smart delays based on context',
                testing: 'Fixed or random delays for development'
            }
        });
        
    } catch (error) {
        console.error('❌ Delay stats error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to get delay stats',
            error: error.message
        });
    }
});

// Estimate bulk sending time
router.post('/estimate-time', (req, res) => {
    try {
        const { emailCount, avgDelay = 75 } = req.body;
        
        if (!emailCount || emailCount < 1) {
            return res.status(400).json({
                success: false,
                message: 'Valid email count is required'
            });
        }
        
        const estimation = emailDelayUtils.estimateBulkSendingTime(emailCount, avgDelay);
        
        res.json({
            success: true,
            emailCount: emailCount,
            averageDelay: `${avgDelay} seconds`,
            estimation: estimation,
            breakdown: {
                processingTime: `${emailCount * 5} seconds (5s per email)`,
                delayTime: `${(emailCount - 1) * avgDelay} seconds (delays between emails)`,
                totalTime: estimation.formatted
            }
        });
        
    } catch (error) {
        console.error('❌ Time estimation error:', error);
        res.status(500).json({
            success: false,
            message: 'Time estimation failed',
            error: error.message
        });
    }
});

// Test progressive delay pattern
router.post('/test-progressive', async (req, res) => {
    try {
        const { emailCount = 5 } = req.body;
        
        console.log(`🧪 Testing progressive delay pattern for ${emailCount} emails...`);
        
        const results = [];
        const startTime = Date.now();
        
        for (let i = 0; i < emailCount; i++) {
            const emailStartTime = Date.now();
            
            // Simulate email sending (skip first email delay)
            if (i > 0) {
                const delayMs = await emailDelayUtils.progressiveDelay(i, emailCount);
                results.push({
                    emailIndex: i,
                    delayMs: delayMs,
                    delayFormatted: emailDelayUtils.formatDelayTime(delayMs),
                    timestamp: new Date().toLocaleTimeString()
                });
            } else {
                results.push({
                    emailIndex: i,
                    delayMs: 0,
                    delayFormatted: 'No delay (first email)',
                    timestamp: new Date().toLocaleTimeString()
                });
            }
            
            console.log(`📧 Simulated email ${i + 1}/${emailCount} sent`);
        }
        
        const totalTime = Date.now() - startTime;
        
        res.json({
            success: true,
            message: `Progressive delay test completed for ${emailCount} emails`,
            totalTimeMs: totalTime,
            totalTimeFormatted: emailDelayUtils.formatDelayTime(totalTime),
            emailDelays: results,
            pattern: 'Progressive delays increase over time to avoid detection patterns'
        });
        
    } catch (error) {
        console.error('❌ Progressive delay test error:', error);
        res.status(500).json({
            success: false,
            message: 'Progressive delay test failed',
            error: error.message
        });
    }
});

// Test smart delay based on current conditions
router.post('/test-smart', async (req, res) => {
    try {
        const { emailsSentToday = 0 } = req.body;
        
        console.log(`🧪 Testing smart delay with ${emailsSentToday} emails sent today...`);
        
        const currentTime = new Date();
        const startTime = Date.now();
        
        const delayMs = await emailDelayUtils.smartDelay(emailsSentToday, currentTime);
        
        const endTime = Date.now();
        const actualDelay = Math.round((endTime - startTime) / 1000);
        
        res.json({
            success: true,
            message: 'Smart delay test completed',
            context: {
                currentTime: currentTime.toLocaleTimeString(),
                currentHour: currentTime.getHours(),
                emailsSentToday: emailsSentToday,
                timeOfDayCategory: getTimeCategory(currentTime.getHours())
            },
            result: {
                delayMs: delayMs,
                delaySeconds: actualDelay,
                delayFormatted: emailDelayUtils.formatDelayTime(delayMs)
            },
            explanation: 'Smart delay adjusts based on time of day and email volume'
        });
        
    } catch (error) {
        console.error('❌ Smart delay test error:', error);
        res.status(500).json({
            success: false,
            message: 'Smart delay test failed',
            error: error.message
        });
    }
});

// Helper function to categorize time of day
function getTimeCategory(hour) {
    if (hour >= 9 && hour <= 17) {
        return 'Business Hours (faster delays)';
    } else if (hour >= 18 && hour <= 21) {
        return 'Evening (moderate delays)';
    } else {
        return 'Night/Early Morning (slower delays)';
    }
}

module.exports = router;