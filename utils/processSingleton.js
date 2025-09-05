const fs = require('fs');
const path = require('path');

/**
 * Process Singleton Manager
 * Ensures only one instance of the application runs at a time
 * Prevents duplicate server instances that could cause campaign conflicts
 */
class ProcessSingleton {
    constructor(name = 'lga-server') {
        this.name = name;
        this.lockFile = path.join(__dirname, '../locks', `${name}.pid`);
        this.ensureLockDirectory();
    }

    ensureLockDirectory() {
        const lockDir = path.dirname(this.lockFile);
        if (!fs.existsSync(lockDir)) {
            fs.mkdirSync(lockDir, { recursive: true });
        }
    }

    /**
     * Check if another instance is already running
     * @returns {boolean} - True if another instance is running
     */
    isAnotherInstanceRunning() {
        try {
            if (!fs.existsSync(this.lockFile)) {
                return false;
            }

            const lockData = JSON.parse(fs.readFileSync(this.lockFile, 'utf8'));
            
            // Check if the process is still running
            if (this.isProcessRunning(lockData.pid)) {
                console.log(`🔒 Another instance is already running (PID: ${lockData.pid})`);
                console.log(`📍 Started: ${new Date(lockData.startTime).toLocaleString()}`);
                console.log(`🖥️  Port: ${lockData.port || 'unknown'}`);
                return true;
            } else {
                // Process is dead, clean up stale lock
                console.log(`🧹 Cleaning up stale lock file (PID ${lockData.pid} no longer running)`);
                fs.unlinkSync(this.lockFile);
                return false;
            }
        } catch (error) {
            console.error('❌ Error checking for running instance:', error.message);
            // If we can't determine, assume no other instance (fail open)
            return false;
        }
    }

    /**
     * Check if a process is still running
     * @param {number} pid - Process ID to check
     * @returns {boolean} - True if process is running
     */
    isProcessRunning(pid) {
        try {
            // On Windows, sending signal 0 checks if process exists without killing it
            process.kill(pid, 0);
            return true;
        } catch (error) {
            // ESRCH means process doesn't exist
            return error.code !== 'ESRCH';
        }
    }

    /**
     * Create a lock file for this instance
     * @param {number} port - Port number the server is running on
     */
    createLock(port = 3000) {
        try {
            const lockData = {
                pid: process.pid,
                port,
                startTime: new Date().toISOString(),
                version: process.version,
                platform: process.platform,
                cwd: process.cwd()
            };

            fs.writeFileSync(this.lockFile, JSON.stringify(lockData, null, 2));
            console.log(`🔐 Process lock created (PID: ${process.pid}, Port: ${port})`);
        } catch (error) {
            console.error('❌ Failed to create process lock:', error.message);
        }
    }

    /**
     * Remove the lock file
     */
    removeLock() {
        try {
            if (fs.existsSync(this.lockFile)) {
                const lockData = JSON.parse(fs.readFileSync(this.lockFile, 'utf8'));
                
                // Only remove lock if we own it
                if (lockData.pid === process.pid) {
                    fs.unlinkSync(this.lockFile);
                    console.log(`🔓 Process lock removed (PID: ${process.pid})`);
                } else {
                    console.log(`⚠️ Cannot remove lock owned by different process (PID: ${lockData.pid})`);
                }
            }
        } catch (error) {
            console.error('❌ Failed to remove process lock:', error.message);
        }
    }

    /**
     * Setup exit handlers to clean up lock
     */
    setupExitHandlers() {
        const cleanup = () => {
            this.removeLock();
        };

        // Handle various exit signals
        process.on('exit', cleanup);
        process.on('SIGINT', () => {
            console.log('\\n📝 Shutting down server...');
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

    /**
     * Force start by removing any existing locks
     * USE WITH CAUTION - only for development/debugging
     */
    forceStart() {
        if (fs.existsSync(this.lockFile)) {
            console.log('🚨 FORCE START: Removing existing lock file');
            fs.unlinkSync(this.lockFile);
        }
    }

    /**
     * Get information about the running instance
     * @returns {Object|null} - Lock information or null if no instance running
     */
    getRunningInstanceInfo() {
        try {
            if (!fs.existsSync(this.lockFile)) {
                return null;
            }

            const lockData = JSON.parse(fs.readFileSync(this.lockFile, 'utf8'));
            
            if (this.isProcessRunning(lockData.pid)) {
                return {
                    ...lockData,
                    uptime: Date.now() - new Date(lockData.startTime).getTime(),
                    isRunning: true
                };
            } else {
                return {
                    ...lockData,
                    isRunning: false,
                    staleLock: true
                };
            }
        } catch (error) {
            console.error('❌ Error getting instance info:', error.message);
            return null;
        }
    }
}

module.exports = ProcessSingleton;