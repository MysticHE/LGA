/**
 * Email Delay Utilities
 * Provides random delay functionality to improve email deliverability
 * and avoid being flagged as spam by email providers
 */

class EmailDelayUtils {
    constructor() {
        // Default delay range: 30-120 seconds
        this.minDelay = 30000; // 30 seconds in milliseconds
        this.maxDelay = 120000; // 120 seconds in milliseconds
    }

    /**
     * Generate random delay between min and max values
     * @param {number} min - Minimum delay in milliseconds (default: 30000)
     * @param {number} max - Maximum delay in milliseconds (default: 120000)
     * @returns {number} Random delay in milliseconds
     */
    getRandomDelay(min = this.minDelay, max = this.maxDelay) {
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }

    /**
     * Sleep/wait for specified milliseconds
     * @param {number} ms - Milliseconds to wait
     * @returns {Promise} Promise that resolves after delay
     */
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Wait for random delay between 30-120 seconds
     * @param {number} min - Minimum delay in seconds (default: 30)
     * @param {number} max - Maximum delay in seconds (default: 120)
     * @returns {Promise} Promise that resolves after random delay
     */
    async randomDelay(min = 30, max = 120) {
        const minMs = min * 1000;
        const maxMs = max * 1000;
        const delayMs = this.getRandomDelay(minMs, maxMs);
        
        console.log(`⏳ Waiting ${Math.round(delayMs / 1000)}s before next email...`);
        await this.sleep(delayMs);
        
        return delayMs;
    }

    /**
     * Format delay time for human-readable display
     * @param {number} ms - Milliseconds
     * @returns {string} Formatted time string
     */
    formatDelayTime(ms) {
        const seconds = Math.round(ms / 1000);
        if (seconds < 60) {
            return `${seconds}s`;
        } else {
            const minutes = Math.floor(seconds / 60);
            const remainingSeconds = seconds % 60;
            return `${minutes}m ${remainingSeconds}s`;
        }
    }

    /**
     * Calculate estimated completion time for bulk email sending
     * @param {number} emailCount - Number of emails to send
     * @param {number} avgDelay - Average delay in seconds (default: 75)
     * @returns {object} Estimation object with total time and completion time
     */
    estimateBulkSendingTime(emailCount, avgDelay = 75) {
        const totalDelaySeconds = (emailCount - 1) * avgDelay; // No delay after last email
        const estimatedProcessingSeconds = emailCount * 5; // ~5 seconds per email processing
        const totalSeconds = totalDelaySeconds + estimatedProcessingSeconds;
        
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        
        const completionTime = new Date(Date.now() + (totalSeconds * 1000));
        
        return {
            totalSeconds,
            formatted: hours > 0 ? `${hours}h ${minutes}m ${seconds}s` : 
                      minutes > 0 ? `${minutes}m ${seconds}s` : `${seconds}s`,
            completionTime: completionTime.toLocaleTimeString(),
            estimatedFinish: completionTime
        };
    }

    /**
     * Progressive delay - starts with shorter delays and increases over time
     * Helps avoid detection patterns while maintaining efficiency
     * @param {number} emailIndex - Current email index (0-based)
     * @param {number} totalEmails - Total number of emails
     * @returns {Promise} Promise that resolves after progressive delay
     */
    async progressiveDelay(emailIndex, totalEmails) {
        // Skip delay for first email
        if (emailIndex === 0) {
            return 0;
        }

        // Calculate progressive multiplier (1.0 to 1.5)
        const progressRatio = emailIndex / totalEmails;
        const multiplier = 1.0 + (progressRatio * 0.5);
        
        // Base delays with progressive increase
        const baseMin = 30;
        const baseMax = 120;
        const adjustedMin = Math.round(baseMin * multiplier);
        const adjustedMax = Math.round(baseMax * multiplier);
        
        return await this.randomDelay(adjustedMin, adjustedMax);
    }

    /**
     * Batch delay - adds extra delay between batches of emails
     * @param {number} batchIndex - Current batch index
     * @param {number} batchSize - Size of each batch
     * @returns {Promise} Promise that resolves after batch delay
     */
    async batchDelay(batchIndex, batchSize) {
        if (batchIndex === 0) {
            return 0; // No delay for first batch
        }
        
        // Add 2-5 minutes between batches
        const batchDelayMin = 120; // 2 minutes
        const batchDelayMax = 300; // 5 minutes
        
        console.log(`⏸️ Batch ${batchIndex} completed. Adding extended delay between batches...`);
        return await this.randomDelay(batchDelayMin, batchDelayMax);
    }

    /**
     * Smart delay based on time of day and email volume
     * @param {number} emailsSentToday - Number of emails sent today
     * @param {Date} currentTime - Current time (optional, defaults to now)
     * @returns {Promise} Promise that resolves after smart delay
     */
    async smartDelay(emailsSentToday = 0, currentTime = new Date()) {
        const hour = currentTime.getHours();
        
        // Base delay multipliers based on time of day
        let timeMultiplier = 1.0;
        if (hour >= 9 && hour <= 17) {
            timeMultiplier = 0.8; // Faster during business hours
        } else if (hour >= 18 && hour <= 21) {
            timeMultiplier = 1.2; // Slower during evening
        } else {
            timeMultiplier = 1.5; // Much slower during night/early morning
        }
        
        // Volume multiplier - slow down if many emails sent today
        let volumeMultiplier = 1.0;
        if (emailsSentToday > 100) {
            volumeMultiplier = 1.5;
        } else if (emailsSentToday > 50) {
            volumeMultiplier = 1.2;
        }
        
        // Calculate adjusted delays
        const baseMin = 30;
        const baseMax = 120;
        const adjustedMin = Math.round(baseMin * timeMultiplier * volumeMultiplier);
        const adjustedMax = Math.round(baseMax * timeMultiplier * volumeMultiplier);
        
        console.log(`🧠 Smart delay: Time=${hour}h, Volume=${emailsSentToday}, Multipliers=${timeMultiplier}x${volumeMultiplier}`);
        return await this.randomDelay(adjustedMin, adjustedMax);
    }

    /**
     * Get delay statistics for monitoring
     * @returns {object} Delay configuration and statistics
     */
    getDelayStats() {
        return {
            minDelay: this.minDelay / 1000, // in seconds
            maxDelay: this.maxDelay / 1000, // in seconds
            averageDelay: (this.minDelay + this.maxDelay) / 2000, // in seconds
            range: `${this.minDelay / 1000}-${this.maxDelay / 1000} seconds`
        };
    }
}

module.exports = EmailDelayUtils;