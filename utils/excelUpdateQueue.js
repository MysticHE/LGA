/**
 * Excel Update Queue Manager
 * Prevents race conditions when multiple processes try to update Excel simultaneously
 */

class ExcelUpdateQueue {
    constructor() {
        this.queues = new Map(); // fileId -> array of pending updates
        this.processing = new Map(); // fileId -> boolean
        this.retryConfig = {
            maxRetries: 3,
            baseDelay: 1000,
            maxDelay: 5000
        };
    }

    /**
     * Add update to queue for specific Excel file
     * @param {string} fileId - Excel file identifier (usually email address)
     * @param {Function} updateFunction - Async function to execute
     * @param {object} context - Context data for logging
     * @returns {Promise} Promise that resolves when update completes
     */
    async queueUpdate(fileId, updateFunction, context = {}) {
        return new Promise((resolve, reject) => {
            const updateItem = {
                id: `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
                updateFunction,
                context,
                resolve,
                reject,
                createdAt: new Date(),
                retries: 0
            };

            if (!this.queues.has(fileId)) {
                this.queues.set(fileId, []);
            }

            this.queues.get(fileId).push(updateItem);
            
            console.log(`📋 Queued Excel update: ${context.type || 'update'} for ${context.email || fileId} (queue length: ${this.queues.get(fileId).length})`);
            
            // Start processing if not already running
            this.processQueue(fileId);
        });
    }

    /**
     * Process queued updates for specific file
     * @param {string} fileId - Excel file identifier
     */
    async processQueue(fileId) {
        if (this.processing.get(fileId)) {
            return; // Already processing this file
        }

        const queue = this.queues.get(fileId);
        if (!queue || queue.length === 0) {
            return;
        }

        this.processing.set(fileId, true);

        while (queue.length > 0) {
            const updateItem = queue.shift();
            
            try {
                console.log(`⚡ Processing Excel update: ${updateItem.context.type || 'update'} for ${updateItem.context.email || fileId}`);
                
                const result = await this.executeWithRetry(updateItem);
                updateItem.resolve(result);
                
                console.log(`✅ Excel update completed: ${updateItem.context.type || 'update'} for ${updateItem.context.email || fileId}`);
                
                // Small delay between updates to prevent API throttling
                if (queue.length > 0) {
                    await this.sleep(500); // 500ms between updates
                }
                
            } catch (error) {
                console.error(`❌ Excel update failed: ${updateItem.context.type || 'update'} for ${updateItem.context.email || fileId}:`, error.message);
                updateItem.reject(error);
            }
        }

        this.processing.set(fileId, false);
        console.log(`📋 Excel update queue completed for ${fileId}`);
    }

    /**
     * Execute update with retry logic
     * @param {object} updateItem - Update item with function and context
     * @returns {Promise} Promise that resolves with update result
     */
    async executeWithRetry(updateItem) {
        let lastError;

        for (let attempt = 0; attempt <= this.retryConfig.maxRetries; attempt++) {
            try {
                if (attempt > 0) {
                    const delayMs = Math.min(
                        this.retryConfig.baseDelay * Math.pow(2, attempt - 1),
                        this.retryConfig.maxDelay
                    );
                    console.log(`🔄 Retry attempt ${attempt} for ${updateItem.context.email || 'unknown'} after ${delayMs}ms...`);
                    await this.sleep(delayMs);
                }

                return await updateItem.updateFunction();
                
            } catch (error) {
                lastError = error;
                console.log(`⚠️ Excel update attempt ${attempt + 1} failed: ${error.message}`);
                
                // Don't retry on authentication errors
                if (error.message.includes('401') || error.message.includes('unauthorized')) {
                    throw error;
                }
            }
        }

        throw lastError;
    }

    /**
     * Sleep utility
     * @param {number} ms - Milliseconds to wait
     */
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Get queue status for monitoring
     * @returns {object} Status of all queues
     */
    getQueueStatus() {
        const status = {};
        for (const [fileId, queue] of this.queues.entries()) {
            status[fileId] = {
                pending: queue.length,
                processing: this.processing.get(fileId) || false,
                lastUpdate: queue.length > 0 ? queue[queue.length - 1].createdAt : null
            };
        }
        return status;
    }
}

// Global instance
const excelUpdateQueue = new ExcelUpdateQueue();

module.exports = excelUpdateQueue;