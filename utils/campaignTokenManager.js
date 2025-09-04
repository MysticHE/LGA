/**
 * Campaign Token Manager
 * Prevents token expiration during long-running email campaigns
 * Proactively refreshes tokens when campaigns exceed token lifetime
 */

class CampaignTokenManager {
    constructor() {
        this.tokenLifetime = 60 * 60 * 1000; // 1 hour in ms
        this.refreshBuffer = 10 * 60 * 1000; // 10 minutes buffer
        this.campaignStartTimes = new Map(); // Track when campaigns started
    }

    /**
     * Start campaign token tracking
     */
    startCampaignTracking(sessionId, estimatedDurationMs) {
        const startTime = new Date();
        this.campaignStartTimes.set(sessionId, {
            startTime: startTime,
            estimatedDuration: estimatedDurationMs,
            lastTokenRefresh: startTime,
            tokenRefreshCount: 0
        });

        console.log(`🕐 Campaign token tracking started for session ${sessionId}`);
        console.log(`📊 Campaign duration: ${Math.round(estimatedDurationMs / 60000)} minutes`);
        
        // If campaign will exceed token lifetime, warn about token management
        if (estimatedDurationMs > (this.tokenLifetime - this.refreshBuffer)) {
            console.log(`⚠️ Long campaign detected - will refresh tokens during execution`);
        }
    }

    /**
     * Check if token refresh is needed during campaign
     * Returns true if token should be refreshed before next operation
     */
    needsTokenRefresh(sessionId) {
        const campaign = this.campaignStartTimes.get(sessionId);
        if (!campaign) {
            return false; // No active campaign
        }

        const now = new Date();
        const timeSinceLastRefresh = now.getTime() - campaign.lastTokenRefresh.getTime();
        
        // Refresh token if it's been more than 50 minutes since last refresh
        const refreshThreshold = this.tokenLifetime - this.refreshBuffer;
        
        if (timeSinceLastRefresh > refreshThreshold) {
            console.log(`🔄 Token refresh needed for session ${sessionId}`);
            console.log(`⏰ Time since last refresh: ${Math.round(timeSinceLastRefresh / 60000)} minutes`);
            return true;
        }

        return false;
    }

    /**
     * Record successful token refresh during campaign
     */
    recordTokenRefresh(sessionId) {
        const campaign = this.campaignStartTimes.get(sessionId);
        if (campaign) {
            campaign.lastTokenRefresh = new Date();
            campaign.tokenRefreshCount++;
            console.log(`✅ Token refresh recorded for session ${sessionId} (count: ${campaign.tokenRefreshCount})`);
        }
    }

    /**
     * Proactively refresh token if needed during campaign
     */
    async ensureValidToken(authProvider, sessionId) {
        if (!this.needsTokenRefresh(sessionId)) {
            return true; // Token is still valid
        }

        try {
            console.log(`🔄 Proactively refreshing token during campaign...`);
            
            // Force token refresh by requesting new access token
            await authProvider.getAccessToken(sessionId);
            
            // Record the refresh
            this.recordTokenRefresh(sessionId);
            
            console.log(`✅ Campaign token refresh successful`);
            return true;
            
        } catch (error) {
            console.error(`❌ Campaign token refresh failed:`, error.message);
            return false;
        }
    }

    /**
     * End campaign tracking and cleanup
     */
    endCampaignTracking(sessionId) {
        const campaign = this.campaignStartTimes.get(sessionId);
        if (campaign) {
            const duration = new Date().getTime() - campaign.startTime.getTime();
            console.log(`🏁 Campaign completed for session ${sessionId}`);
            console.log(`📊 Total duration: ${Math.round(duration / 60000)} minutes`);
            console.log(`🔄 Token refreshes performed: ${campaign.tokenRefreshCount}`);
            
            this.campaignStartTimes.delete(sessionId);
        }
    }

    /**
     * Get campaign statistics
     */
    getCampaignStats(sessionId) {
        const campaign = this.campaignStartTimes.get(sessionId);
        if (!campaign) {
            return null;
        }

        const now = new Date();
        const elapsed = now.getTime() - campaign.startTime.getTime();
        const timeSinceLastRefresh = now.getTime() - campaign.lastTokenRefresh.getTime();
        
        return {
            startTime: campaign.startTime,
            elapsed: elapsed,
            elapsedMinutes: Math.round(elapsed / 60000),
            estimatedDurationMinutes: Math.round(campaign.estimatedDuration / 60000),
            tokenRefreshCount: campaign.tokenRefreshCount,
            timeSinceLastRefreshMinutes: Math.round(timeSinceLastRefresh / 60000),
            needsRefresh: this.needsTokenRefresh(sessionId)
        };
    }

    /**
     * Check token every N emails (batch refresh check)
     */
    shouldCheckToken(emailIndex, checkInterval = 10) {
        return emailIndex > 0 && emailIndex % checkInterval === 0;
    }
}

module.exports = CampaignTokenManager;