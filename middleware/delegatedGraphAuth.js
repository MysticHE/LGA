const { Client } = require('@microsoft/microsoft-graph-client');
const { AuthenticationProvider } = require('@microsoft/microsoft-graph-client');
const msal = require('@azure/msal-node');
const persistentStorage = require('../utils/persistentStorage');

/**
 * Microsoft Graph Delegated Authentication Provider
 * Uses MSAL for user authentication with OAuth2 flow
 */

class DelegatedGraphAuth {
    constructor() {
        this.tenantId = process.env.AZURE_TENANT_ID;
        this.clientId = process.env.AZURE_CLIENT_ID;
        this.clientSecret = process.env.AZURE_CLIENT_SECRET;
        this.redirectUri = process.env.RENDER_EXTERNAL_URL 
            ? `${process.env.RENDER_EXTERNAL_URL}/auth/callback`
            : 'http://localhost:3000/auth/callback';
        
        this.msalConfig = {
            auth: {
                clientId: this.clientId,
                clientSecret: this.clientSecret,
                authority: `https://login.microsoftonline.com/${this.tenantId}`
            },
            system: {
                loggerOptions: {
                    loggerCallback: (level, message, containsPii) => {
                        if (process.env.NODE_ENV === 'development') {
                            console.log(message);
                        }
                    },
                    piiLoggingEnabled: false,
                    logLevel: process.env.NODE_ENV === 'development' ? msal.LogLevel.Verbose : msal.LogLevel.Warning,
                }
            }
        };
        
        // Only initialize if credentials are available
        if (this.clientId && this.clientSecret && this.tenantId) {
            this.msalInstance = new msal.ConfidentialClientApplication(this.msalConfig);
            this.userTokens = new Map(); // In-memory cache
            
            // Load existing sessions on startup
            this.loadPersistedSessions();
            
            this.validateConfig();
            console.log('✅ Delegated Microsoft Graph authentication initialized with persistent storage');
        } else {
            console.log('⚠️ Azure credentials not provided - Microsoft Graph features disabled');
            this.msalInstance = null;
            this.userTokens = new Map();
        }
    }

    validateConfig() {
        const required = ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET'];
        const missing = required.filter(key => !process.env[key]);
        
        if (missing.length > 0) {
            console.log('⚠️ Azure Configuration Warning:');
            console.log(`   Missing environment variables: ${missing.join(', ')}`);
            console.log('   Email automation features will be disabled');
        }
    }

    // Get authorization URL for user login
    getAuthUrl(sessionId) {
        if (!this.msalInstance) {
            throw new Error('Azure authentication not available - missing credentials');
        }
        const authUrlParameters = {
            scopes: [
                'https://graph.microsoft.com/User.Read',
                'https://graph.microsoft.com/Files.ReadWrite.All',
                'https://graph.microsoft.com/Mail.Send',
                'https://graph.microsoft.com/Mail.ReadWrite',
                'offline_access'
            ],
            redirectUri: this.redirectUri,
            state: sessionId // Use session ID to track the request
        };

        return this.msalInstance.getAuthCodeUrl(authUrlParameters);
    }

    // Handle the callback from Microsoft login
    async handleCallback(code, sessionId) {
        try {
            const tokenRequest = {
                code: code,
                scopes: [
                    'https://graph.microsoft.com/User.Read',
                    'https://graph.microsoft.com/Files.ReadWrite.All',
                    'https://graph.microsoft.com/Mail.Send',
                    'https://graph.microsoft.com/Mail.ReadWrite',
                    'offline_access'
                ],
                redirectUri: this.redirectUri,
            };

            const response = await this.msalInstance.acquireTokenByCode(tokenRequest);
            
            // Store tokens for this session
            this.userTokens.set(sessionId, {
                accessToken: response.accessToken,
                refreshToken: response.refreshToken,
                expiresOn: response.expiresOn,
                account: response.account,
                scopes: response.scopes,
                createdAt: new Date().toISOString()
            });

            console.log(`✅ User authenticated: ${response.account.username}`);
            
            // Save to persistent storage immediately (async, don't wait)
            setImmediate(async () => {
                try {
                    await persistentStorage.saveSessions(this.userTokens);
                    await persistentStorage.saveUserContext(
                        sessionId, 
                        response.account.username,
                        '/LGA-Email-Automation'
                    );
                    console.log(`💾 Session persisted: ${sessionId}`);
                } catch (persistError) {
                    console.error('❌ Failed to persist session:', persistError);
                }
            });
            return {
                success: true,
                user: response.account.username,
                sessionId: sessionId
            };

        } catch (error) {
            console.error('Authentication callback error:', error.message);
            return {
                success: false,
                error: error.message
            };
        }
    }

    // Get valid access token for a session
    async getAccessToken(sessionId) {
        const tokenData = this.userTokens.get(sessionId);
        
        if (!tokenData) {
            throw new Error('User not authenticated');
        }

        // Handle sessions that need refresh (restored from persistent storage)
        if (tokenData.needsRefresh || !tokenData.accessToken) {
            console.log(`🔄 Session ${sessionId} needs token refresh (restored from storage)`);
            return await this.refreshSessionToken(sessionId);
        }

        // Check if token is still valid (with 5 minute buffer)
        const now = new Date();
        const expiresOn = new Date(tokenData.expiresOn);
        const timeUntilExpiry = expiresOn.getTime() - now.getTime();
        
        if (timeUntilExpiry > 5 * 60 * 1000) { // More than 5 minutes left
            return tokenData.accessToken;
        }

        // Token is expired or will expire soon, refresh it
        console.log(`🔄 Token expiring soon for session ${sessionId}, refreshing...`);
        return await this.refreshSessionToken(sessionId);
    }

    /**
     * Refresh token for a specific session
     */
    async refreshSessionToken(sessionId) {
        const tokenData = this.userTokens.get(sessionId);
        
        if (!tokenData || !tokenData.refreshToken) {
            console.error(`❌ No refresh token available for session ${sessionId}`);
            this.userTokens.delete(sessionId);
            throw new Error('Authentication expired, please login again');
        }

        try {
            const refreshTokenRequest = {
                refreshToken: tokenData.refreshToken,
                scopes: tokenData.scopes || [
                    'https://graph.microsoft.com/User.Read',
                    'https://graph.microsoft.com/Files.ReadWrite.All',
                    'https://graph.microsoft.com/Mail.Send',
                    'https://graph.microsoft.com/Mail.ReadWrite'
                ],
            };

            const response = await this.msalInstance.acquireTokenByRefreshToken(refreshTokenRequest);
            
            // Update stored tokens with refreshed data
            const updatedTokenData = {
                accessToken: response.accessToken,
                refreshToken: response.refreshToken || tokenData.refreshToken,
                expiresOn: response.expiresOn,
                account: response.account,
                scopes: response.scopes,
                createdAt: tokenData.createdAt,
                needsRefresh: false, // Clear the refresh flag
                hasStoredRefreshToken: true
            };

            this.userTokens.set(sessionId, updatedTokenData);

            console.log(`✅ Token refreshed for user: ${response.account.username}`);
            
            // Save updated session to persistent storage (async, don't wait)
            setImmediate(async () => {
                try {
                    await persistentStorage.saveSessions(this.userTokens);
                    console.log(`💾 Updated session saved to storage: ${sessionId}`);
                } catch (saveError) {
                    console.error('❌ Failed to save updated session:', saveError);
                }
            });

            return response.accessToken;

        } catch (error) {
            console.error(`❌ Token refresh failed for session ${sessionId}:`, error.message);
            // Remove invalid tokens
            this.userTokens.delete(sessionId);
            throw new Error('Authentication expired, please login again');
        }
    }

    // Create Microsoft Graph client for authenticated user
    async getGraphClient(sessionId) {
        const accessToken = await this.getAccessToken(sessionId);
        
        const authProvider = {
            getAccessToken: async () => {
                return accessToken;
            }
        };

        return Client.initWithMiddleware({
            authProvider: authProvider,
            debugLogging: process.env.NODE_ENV === 'development'
        });
    }

    // Check if user is authenticated
    isAuthenticated(sessionId) {
        return this.userTokens.has(sessionId);
    }

    // Get user info
    getUserInfo(sessionId) {
        const tokenData = this.userTokens.get(sessionId);
        return tokenData ? tokenData.account : null;
    }

    // Get all active sessions
    getActiveSessions() {
        return Array.from(this.userTokens.keys());
    }

    // Load persisted sessions on startup with complete token data
    async loadPersistedSessions() {
        try {
            const sessions = await persistentStorage.loadSessions();
            
            for (const [sessionId, sessionData] of Object.entries(sessions)) {
                // Only load sessions that haven't expired
                const expiresOn = new Date(sessionData.expiresOn);
                if (expiresOn > new Date()) {
                    // Create complete session entry with decrypted refresh tokens
                    const restoredSession = {
                        account: sessionData.account,
                        expiresOn: sessionData.expiresOn,
                        needsRefresh: sessionData.needsRefresh || true,
                        createdAt: sessionData.createdAt,
                        scopes: sessionData.scopes,
                        hasStoredRefreshToken: sessionData.hasStoredRefreshToken
                    };

                    // Include refresh token if available (already decrypted by persistentStorage)
                    if (sessionData.refreshToken) {
                        restoredSession.refreshToken = sessionData.refreshToken;
                        console.log(`✅ Restored session with refresh token: ${sessionId}`);
                    } else if (sessionData.hasStoredRefreshToken) {
                        console.warn(`⚠️ Session ${sessionId} should have refresh token but decryption failed`);
                    }

                    this.userTokens.set(sessionId, restoredSession);
                }
            }
            
            console.log(`🔄 Restored ${this.userTokens.size} sessions from persistent storage with background capability`);
            
            // Start background token refresh for restored sessions
            if (this.userTokens.size > 0) {
                console.log('🚀 Starting background token refresh for restored sessions...');
                setImmediate(() => this.refreshExpiringSessions());
            }
        } catch (error) {
            console.error('❌ Failed to load persisted sessions:', error.message);
        }
    }

    /**
     * Proactively refresh tokens that will expire soon
     * This keeps sessions active for background operations
     */
    async refreshExpiringSessions() {
        try {
            const now = new Date();
            const refreshThreshold = 15 * 60 * 1000; // Refresh if expires within 15 minutes
            let refreshedCount = 0;
            const refreshPromises = [];

            console.log(`🔄 Checking ${this.userTokens.size} sessions for token refresh...`);

            for (const [sessionId, tokenData] of this.userTokens.entries()) {
                const expiresOn = new Date(tokenData.expiresOn);
                const timeUntilExpiry = expiresOn.getTime() - now.getTime();
                
                // Skip if session needs refresh or expires soon
                if (tokenData.needsRefresh || timeUntilExpiry < refreshThreshold) {
                    if (tokenData.refreshToken) {
                        console.log(`🔄 Proactively refreshing session: ${sessionId} (expires in ${Math.round(timeUntilExpiry / 60000)} minutes)`);
                        refreshPromises.push(
                            this.refreshSessionToken(sessionId)
                                .then(() => {
                                    refreshedCount++;
                                    console.log(`✅ Background refresh successful: ${sessionId}`);
                                })
                                .catch(error => {
                                    console.error(`❌ Background refresh failed for ${sessionId}:`, error.message);
                                })
                        );
                    } else {
                        console.warn(`⚠️ Session ${sessionId} needs refresh but no refresh token available`);
                    }
                }
            }

            if (refreshPromises.length > 0) {
                await Promise.allSettled(refreshPromises);
                console.log(`🎉 Background token refresh completed: ${refreshedCount}/${refreshPromises.length} sessions refreshed`);
            } else {
                console.log(`✅ All tokens are fresh, no refresh needed`);
            }

        } catch (error) {
            console.error('❌ Background token refresh error:', error.message);
        }
    }

    // Clean up expired sessions and save to persistent storage
    async cleanupExpiredSessions() {
        const now = new Date();
        let cleanedCount = 0;
        
        for (const [sessionId, tokenData] of this.userTokens.entries()) {
            const expiresOn = new Date(tokenData.expiresOn);
            // Remove sessions that expired more than 1 hour ago (beyond refresh token validity)
            if (expiresOn.getTime() < (now.getTime() - 60 * 60 * 1000)) {
                this.userTokens.delete(sessionId);
                cleanedCount++;
                console.log(`🧹 Cleaned up expired session: ${sessionId}`);
            }
        }
        
        // Save updated sessions to persistent storage
        await persistentStorage.saveSessions(this.userTokens);
        
        return cleanedCount;
    }

    // Logout user
    logout(sessionId) {
        this.userTokens.delete(sessionId);
        console.log(`✅ User logged out: ${sessionId}`);
    }

    // Test connection for authenticated user
    async testConnection(sessionId) {
        try {
            const client = await this.getGraphClient(sessionId);
            const userInfo = await client.api('/me').get();
            
            console.log(`✅ Graph connection test successful for: ${userInfo.displayName}`);
            return {
                success: true,
                user: userInfo.displayName,
                email: userInfo.mail || userInfo.userPrincipalName
            };
        } catch (error) {
            console.error('Graph connection test failed:', error.message);
            return {
                success: false,
                error: error.message
            };
        }
    }
}

// Create singleton instance
let delegatedAuthInstance = null;

function getDelegatedAuthProvider() {
    if (!delegatedAuthInstance) {
        delegatedAuthInstance = new DelegatedGraphAuth();
    }
    return delegatedAuthInstance;
}

// Middleware function to check authentication
function requireDelegatedAuth(req, res, next) {
    try {
        // Check if Azure credentials are available
        const requiredAzureVars = ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET'];
        const missingAzureVars = requiredAzureVars.filter(envVar => !process.env[envVar]);
        
        if (missingAzureVars.length > 0) {
            console.error(`❌ Email automation blocked - Missing Azure credentials: ${missingAzureVars.join(', ')}`);
            return res.status(503).json({
                error: 'Service Unavailable',
                message: 'Email automation is currently disabled due to missing Azure configuration',
                missingCredentials: missingAzureVars,
                troubleshooting: {
                    step1: 'Check your Render environment variables',
                    step2: 'Ensure AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET are set',
                    step3: 'Restart the service after adding credentials'
                }
            });
        }
        
        const authProvider = getDelegatedAuthProvider();
        const sessionId = req.session?.id || req.headers['x-session-id'];
        
        if (!sessionId) {
            console.warn('⚠️  No session ID provided in request');
            return res.status(401).json({
                error: 'Authentication Required',
                message: 'Please authenticate with Microsoft 365 - No session ID provided',
                authUrl: `/auth/login?redirect=${encodeURIComponent(req.originalUrl)}`,
                troubleshooting: {
                    issue: 'No session ID found',
                    expectedHeader: 'X-Session-Id',
                    receivedHeaders: Object.keys(req.headers).filter(h => h.toLowerCase().includes('session'))
                }
            });
        }
        
        if (!authProvider.isAuthenticated(sessionId)) {
            console.warn(`⚠️  Session ${sessionId} not authenticated`);
            return res.status(401).json({
                error: 'Authentication Required',
                message: 'Please authenticate with Microsoft 365 - Session not found or expired',
                authUrl: `/auth/login?redirect=${encodeURIComponent(req.originalUrl)}`,
                troubleshooting: {
                    issue: 'Session not authenticated',
                    sessionId: sessionId,
                    activeSessions: authProvider.getActiveSessions().length
                }
            });
        }

        req.delegatedAuth = authProvider;
        req.sessionId = sessionId;
        next();
    } catch (error) {
        console.error('❌ Delegated auth middleware error:', error.message);
        res.status(500).json({
            error: 'Authentication Error',
            message: 'Failed to validate authentication',
            details: process.env.NODE_ENV === 'development' ? error.message : 'Internal server error',
            troubleshooting: {
                error: error.message,
                timestamp: new Date().toISOString(),
                endpoint: req.originalUrl
            }
        });
    }
}

module.exports = {
    DelegatedGraphAuth,
    getDelegatedAuthProvider,
    requireDelegatedAuth
};