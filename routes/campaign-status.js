const express = require('express');
const { requireDelegatedAuth } = require('../middleware/delegatedGraphAuth');
const CampaignLockManager = require('../utils/campaignLockManager');
const ProcessSingleton = require('../utils/processSingleton');
const router = express.Router();

// Initialize managers
const campaignLockManager = new CampaignLockManager();
const processSingleton = new ProcessSingleton('lga-server');

/**
 * Campaign Status and Control API
 * Provides monitoring and control endpoints for campaign management
 */

// Get campaign status for current session
router.get('/status/:sessionId', requireDelegatedAuth, async (req, res) => {
    try {
        const { sessionId } = req.params;
        
        const isLocked = campaignLockManager.isLocked(sessionId);
        const activeLocks = campaignLockManager.getActiveLocks();
        const sessionLock = activeLocks.find(lock => lock.sessionId === sessionId);
        
        res.json({
            success: true,
            sessionId,
            campaignRunning: isLocked,
            lockInfo: sessionLock || null,
            activeCampaigns: activeLocks.length,
            systemInfo: {
                processId: process.pid,
                uptime: Math.round(process.uptime()),
                memoryUsage: process.memoryUsage(),
                platform: process.platform
            }
        });
    } catch (error) {
        console.error('❌ Status check error:', error.message);
        res.status(500).json({
            success: false,
            message: 'Failed to get campaign status',
            error: error.message
        });
    }
});

// Get all active campaigns across all sessions
router.get('/active', requireDelegatedAuth, async (req, res) => {
    try {
        const activeLocks = campaignLockManager.getActiveLocks();
        const runningInstanceInfo = processSingleton.getRunningInstanceInfo();
        
        res.json({
            success: true,
            activeCampaigns: activeLocks,
            totalActive: activeLocks.length,
            serverInfo: runningInstanceInfo,
            timestamp: new Date().toISOString()
        });
    } catch (error) {
        console.error('❌ Active campaigns check error:', error.message);
        res.status(500).json({
            success: false,
            message: 'Failed to get active campaigns',
            error: error.message
        });
    }
});

// Emergency stop campaign for session
router.post('/stop/:sessionId', requireDelegatedAuth, async (req, res) => {
    try {
        const { sessionId } = req.params;
        const { force = false } = req.body;
        
        // Check if campaign is actually running
        const isLocked = campaignLockManager.isLocked(sessionId);
        
        if (!isLocked) {
            return res.json({
                success: true,
                message: 'No active campaign found for this session',
                sessionId,
                action: 'none_required'
            });
        }
        
        // Get lock info before removing
        const activeLocks = campaignLockManager.getActiveLocks();
        const sessionLock = activeLocks.find(lock => lock.sessionId === sessionId);
        
        if (force || (sessionLock && sessionLock.pid === process.pid)) {
            campaignLockManager.releaseLock(sessionId);
            
            console.log(`🛑 Campaign stopped for session ${sessionId} ${force ? '(FORCED)' : '(NORMAL)'}`);
            
            res.json({
                success: true,
                message: `Campaign stopped for session ${sessionId}`,
                sessionId,
                action: 'stopped',
                wasForced: force,
                previousLockInfo: sessionLock
            });
        } else {
            res.status(403).json({
                success: false,
                message: 'Cannot stop campaign owned by different process',
                sessionId,
                lockInfo: sessionLock,
                hint: 'Use force=true parameter to override (use with caution)'
            });
        }
    } catch (error) {
        console.error('❌ Campaign stop error:', error.message);
        res.status(500).json({
            success: false,
            message: 'Failed to stop campaign',
            error: error.message
        });
    }
});

// Emergency cleanup - remove all campaign locks (admin only)
router.post('/cleanup/all', requireDelegatedAuth, async (req, res) => {
    try {
        const { confirm = false } = req.body;
        
        if (!confirm) {
            return res.status(400).json({
                success: false,
                message: 'Cleanup requires confirmation',
                hint: 'Send POST with {"confirm": true} to proceed',
                warning: 'This will forcibly stop ALL active campaigns'
            });
        }
        
        const activeLocks = campaignLockManager.getActiveLocks();
        campaignLockManager.cleanupAllLocks();
        
        console.log(`🧹 EMERGENCY CLEANUP: Removed ${activeLocks.length} campaign locks`);
        
        res.json({
            success: true,
            message: 'All campaign locks cleaned up',
            locksRemoved: activeLocks.length,
            previousLocks: activeLocks,
            warning: 'All active campaigns have been forcibly stopped'
        });
    } catch (error) {
        console.error('❌ Cleanup error:', error.message);
        res.status(500).json({
            success: false,
            message: 'Failed to cleanup locks',
            error: error.message
        });
    }
});

// Health check endpoint
router.get('/health', async (req, res) => {
    try {
        const activeLocks = campaignLockManager.getActiveLocks();
        const runningInstanceInfo = processSingleton.getRunningInstanceInfo();
        
        res.json({
            status: 'healthy',
            timestamp: new Date().toISOString(),
            server: {
                pid: process.pid,
                uptime: Math.round(process.uptime()),
                memory: process.memoryUsage(),
                version: process.version
            },
            campaigns: {
                active: activeLocks.length,
                locks: activeLocks
            },
            singleton: runningInstanceInfo
        });
    } catch (error) {
        res.status(500).json({
            status: 'error',
            message: error.message
        });
    }
});

module.exports = router;