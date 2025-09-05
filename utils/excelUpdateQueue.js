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
     * Add update to queue for specific Excel file with deduplication
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
                priority: context.priority || 'normal',
                timestamp: Date.now(),
                resolve,
                reject,
                createdAt: new Date(),
                retries: 0
            };

            if (!this.queues.has(fileId)) {
                this.queues.set(fileId, []);
            }

            const queue = this.queues.get(fileId);
            
            // DEDUPLICATION: Check for duplicate campaign updates
            if (context.type === 'campaign-complete' || context.type === 'campaign-send') {
                const duplicateIndex = queue.findIndex(item => 
                    item.context.email === context.email && 
                    (item.context.type === 'campaign-complete' || item.context.type === 'campaign-send')
                );
                
                if (duplicateIndex !== -1) {
                    console.log(`🔄 Deduplicating campaign update for ${context.email} - using existing queue item`);
                    // Resolve this promise when the existing item completes
                    const existingItem = queue[duplicateIndex];
                    const originalResolve = existingItem.resolve;
                    const originalReject = existingItem.reject;
                    
                    existingItem.resolve = (result) => {
                        originalResolve(result);
                        resolve(result); // Also resolve this duplicate request
                    };
                    existingItem.reject = (error) => {
                        originalReject(error);
                        reject(error); // Also reject this duplicate request
                    };
                    
                    return; // Don't add duplicate to queue
                }
            }

            // Insert based on priority (high priority first)
            if (updateItem.priority === 'high') {
                const insertIndex = queue.findIndex(item => item.priority !== 'high');
                if (insertIndex === -1) {
                    queue.push(updateItem);
                } else {
                    queue.splice(insertIndex, 0, updateItem);
                }
            } else {
                queue.push(updateItem);
            }
            
            console.log(`📋 Queued Excel update: ${context.type || 'update'} for ${context.email || fileId} (priority: ${updateItem.priority}, queue length: ${queue.length})`);
            
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
                const updateType = updateItem.context.type || 'update';
                const emailAddress = updateItem.context.email || fileId;
                const priority = updateItem.priority === 'high' ? '🔥' : '📊';
                
                console.log(`${priority} Processing Excel update: ${updateType} for ${emailAddress} (priority: ${updateItem.priority})`);
                
                const startTime = Date.now();
                const result = await this.executeWithRetry(updateItem);
                const duration = Date.now() - startTime;
                
                updateItem.resolve(result);
                
                console.log(`✅ Excel update completed: ${updateType} for ${emailAddress} (${duration}ms, priority: ${updateItem.priority})`);
                
                // Enhanced delay between updates to prevent Graph API rate limiting
                if (queue.length > 0) {
                    // Adaptive delay based on queue length and recent failures
                    const baseDelay = 1000; // 1 second minimum
                    const queuePenalty = Math.min(queue.length * 200, 2000); // Up to 2 seconds for large queues
                    const adaptiveDelay = baseDelay + queuePenalty;
                    
                    console.log(`⏳ Waiting ${adaptiveDelay}ms before next update (queue: ${queue.length})`);
                    await this.sleep(adaptiveDelay);
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
                
                // Don't retry on authentication errors or specific Graph API issues
                if (error.message.includes('401') || error.message.includes('unauthorized') ||
                    error.message.includes('Authentication expired') || 
                    error.message.includes('Invalid authentication token')) {
                    console.log('🚨 Authentication error - stopping retries for this update');
                    throw error;
                }
                
                // Reduce retry attempts for rate limiting errors to prevent amplification
                if (error.message.includes('We\'re sorry. We ran into a problem completing your request') ||
                    error.message.includes('TooManyRequests') || 
                    error.message.includes('429')) {
                    console.log('🚨 Rate limiting detected - using reduced retry strategy');
                    if (attempt >= 2) { // Only retry twice for rate limiting
                        throw error;
                    }
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