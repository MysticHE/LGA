const fs = require('fs');
const path = require('path');

/**
 * Campaign Lock Manager
 * Prevents multiple campaign instances from running simultaneously
 * Uses file-based locking to survive server restarts
 */
class CampaignLockManager {
    constructor() {
        this.lockDir = path.join(__dirname, '../locks');
        this.ensureLockDirectory();
    }

    ensureLockDirectory() {
        if (!fs.existsSync(this.lockDir)) {
            fs.mkdirSync(this.lockDir, { recursive: true });
        }
    }

    /**
     * Acquire a campaign lock
     * @param {string} sessionId - Session identifier
     * @param {string} campaignType - Type of campaign (manual, scheduled, etc.)
     * @returns {boolean} - True if lock acquired successfully
     */
    acquireLock(sessionId, campaignType = 'manual') {
        const lockFile = path.join(this.lockDir, `campaign_${sessionId}.lock`);
        
        try {
            // Check if lock file exists
            if (fs.existsSync(lockFile)) {
                const lockData = JSON.parse(fs.readFileSync(lockFile, 'utf8'));
                const lockAge = Date.now() - lockData.timestamp;
                
                // If lock is older than 30 minutes, consider it stale and remove
                if (lockAge > 30 * 60 * 1000) {
                    console.log(`🔓 Removing stale campaign lock (${Math.round(lockAge / 1000 / 60)}m old)`);
                    fs.unlinkSync(lockFile);
                } else {
                    console.log(`🔒 Campaign already running for session ${sessionId} (${campaignType})`);
                    console.log(`📍 Lock details: PID ${lockData.pid}, started ${new Date(lockData.timestamp).toLocaleTimeString()}`);
                    return false;
                }
            }

            // Create new lock
            const lockData = {
                sessionId,
                campaignType,
                pid: process.pid,
                timestamp: Date.now(),
                startTime: new Date().toISOString()
            };

            fs.writeFileSync(lockFile, JSON.stringify(lockData, null, 2));
            console.log(`🔐 Campaign lock acquired for session ${sessionId} (PID: ${process.pid})`);
            return true;

        } catch (error) {
            console.error('❌ Failed to acquire campaign lock:', error.message);
            return false;
        }
    }

    /**
     * Release a campaign lock
     * @param {string} sessionId - Session identifier
     */
    releaseLock(sessionId) {
        const lockFile = path.join(this.lockDir, `campaign_${sessionId}.lock`);
        
        try {
            if (fs.existsSync(lockFile)) {
                const lockData = JSON.parse(fs.readFileSync(lockFile, 'utf8'));
                
                // Verify this process owns the lock
                if (lockData.pid === process.pid) {
                    fs.unlinkSync(lockFile);
                    console.log(`🔓 Campaign lock released for session ${sessionId}`);
                } else {
                    console.log(`⚠️ Cannot release lock owned by different process (PID: ${lockData.pid})`);
                }
            }
        } catch (error) {
            console.error('❌ Failed to release campaign lock:', error.message);
        }
    }

    /**
     * Check if a campaign is currently locked
     * @param {string} sessionId - Session identifier
     * @returns {boolean} - True if campaign is locked
     */
    isLocked(sessionId) {
        const lockFile = path.join(this.lockDir, `campaign_${sessionId}.lock`);
        
        try {
            if (!fs.existsSync(lockFile)) {
                return false;
            }

            const lockData = JSON.parse(fs.readFileSync(lockFile, 'utf8'));
            const lockAge = Date.now() - lockData.timestamp;
            
            // Consider locks older than 30 minutes as stale
            if (lockAge > 30 * 60 * 1000) {
                fs.unlinkSync(lockFile);
                return false;
            }

            return true;
        } catch (error) {
            console.error('❌ Failed to check campaign lock:', error.message);
            return false;
        }
    }

    /**
     * Get active campaign locks
     * @returns {Array} - Array of active lock information
     */
    getActiveLocks() {
        const locks = [];
        
        try {
            if (!fs.existsSync(this.lockDir)) {
                return locks;
            }

            const lockFiles = fs.readdirSync(this.lockDir)
                .filter(file => file.endsWith('.lock'));

            for (const file of lockFiles) {
                try {
                    const lockData = JSON.parse(
                        fs.readFileSync(path.join(this.lockDir, file), 'utf8')
                    );
                    
                    const lockAge = Date.now() - lockData.timestamp;
                    
                    // Skip stale locks
                    if (lockAge <= 30 * 60 * 1000) {
                        locks.push({
                            ...lockData,
                            ageMinutes: Math.round(lockAge / 1000 / 60),
                            file
                        });
                    } else {
                        // Clean up stale lock
                        fs.unlinkSync(path.join(this.lockDir, file));
                    }
                } catch (parseError) {
                    console.error(`❌ Failed to parse lock file ${file}:`, parseError.message);
                }
            }
        } catch (error) {
            console.error('❌ Failed to get active locks:', error.message);
        }

        return locks;
    }

    /**
     * Clean up all locks (emergency cleanup)
     * Use with caution - only for debugging or emergency situations
     */
    cleanupAllLocks() {
        try {
            if (!fs.existsSync(this.lockDir)) {
                return;
            }

            const lockFiles = fs.readdirSync(this.lockDir)
                .filter(file => file.endsWith('.lock'));

            for (const file of lockFiles) {
                fs.unlinkSync(path.join(this.lockDir, file));
            }

            console.log(`🧹 Cleaned up ${lockFiles.length} campaign locks`);
        } catch (error) {
            console.error('❌ Failed to cleanup locks:', error.message);
        }
    }

    /**
     * Setup process exit handlers to cleanup locks
     */
    setupExitHandlers() {
        const cleanup = () => {
            // Find and release any locks owned by this process
            try {
                const locks = this.getActiveLocks();
                const myLocks = locks.filter(lock => lock.pid === process.pid);
                
                for (const lock of myLocks) {
                    const sessionId = lock.sessionId;
                    this.releaseLock(sessionId);
                }
            } catch (error) {
                console.error('❌ Error during lock cleanup:', error.message);
            }
        };

        // Handle various exit signals
        process.on('exit', cleanup);
        process.on('SIGINT', () => {
            console.log('\n📝 Cleaning up campaign locks...');
            cleanup();
            process.exit(0);
        });
        process.on('SIGTERM', cleanup);
        process.on('uncaughtException', (error) => {
            console.error('❌ Uncaught exception:', error.message);
            cleanup();
            process.exit(1);
        });
    }
}

module.exports = CampaignLockManager;