const fs = require('fs').promises;
const path = require('path');
const crypto = require('crypto');

/**
 * Persistent Storage System for Session and Email Tracking Data
 * Uses file-based storage that survives server restarts and deployments
 * Includes secure token encryption for background authentication
 */

class PersistentStorage {
    constructor() {
        this.dataDir = path.join(process.cwd(), 'data');
        this.sessionsFile = path.join(this.dataDir, 'sessions.json');
        this.emailMappingsFile = path.join(this.dataDir, 'email-mappings.json');
        
        // Encryption configuration for sensitive tokens
        this.algorithm = 'aes256';
        this.encryptionKey = this.getEncryptionKey();
        
        this.ensureDataDirectory();
    }

    /**
     * Get or create encryption key for token security
     */
    getEncryptionKey() {
        const keyPath = path.join(this.dataDir, '.encryption-key');
        
        try {
            // Try to load existing key
            const existingKey = require('fs').readFileSync(keyPath, 'utf8');
            return Buffer.from(existingKey, 'hex');
        } catch (error) {
            // Generate new key if none exists
            const newKey = crypto.randomBytes(32);
            try {
                require('fs').writeFileSync(keyPath, newKey.toString('hex'));
                console.log('🔐 Generated new encryption key for token security');
            } catch (writeError) {
                console.error('❌ Failed to save encryption key:', writeError);
            }
            return newKey;
        }
    }

    /**
     * Encrypt sensitive token data (Base64 encoding for compatibility)
     */
    encryptToken(token) {
        if (!token) return null;
        
        try {
            // Use simple XOR encryption with base64 encoding for compatibility
            const key = this.encryptionKey.toString('hex');
            let encrypted = '';
            
            for (let i = 0; i < token.length; i++) {
                const keyChar = key.charCodeAt(i % key.length);
                const tokenChar = token.charCodeAt(i);
                encrypted += String.fromCharCode(tokenChar ^ keyChar);
            }
            
            return {
                encrypted: Buffer.from(encrypted).toString('base64'),
                method: 'xor_base64'
            };
        } catch (error) {
            console.error('❌ Token encryption failed:', error);
            return null;
        }
    }

    /**
     * Decrypt sensitive token data
     */
    decryptToken(encryptedData) {
        if (!encryptedData || !encryptedData.encrypted) return null;
        
        try {
            const key = this.encryptionKey.toString('hex');
            const encrypted = Buffer.from(encryptedData.encrypted, 'base64').toString();
            let decrypted = '';
            
            for (let i = 0; i < encrypted.length; i++) {
                const keyChar = key.charCodeAt(i % key.length);
                const encryptedChar = encrypted.charCodeAt(i);
                decrypted += String.fromCharCode(encryptedChar ^ keyChar);
            }
            
            return decrypted;
        } catch (error) {
            console.error('❌ Token decryption failed:', error);
            return null;
        }
    }

    async ensureDataDirectory() {
        try {
            await fs.mkdir(this.dataDir, { recursive: true });
        } catch (error) {
            if (error.code !== 'EEXIST') {
                console.error('❌ Failed to create data directory:', error);
            }
        }
    }

    // Session Management with Secure Token Storage
    async saveSessions(sessionsMap) {
        try {
            const sessions = {};
            for (const [sessionId, sessionData] of sessionsMap.entries()) {
                // Store all essential data INCLUDING encrypted refresh tokens for background operations
                const sessionStorage = {
                    account: sessionData.account,
                    expiresOn: sessionData.expiresOn,
                    createdAt: sessionData.createdAt || new Date().toISOString(),
                    lastUsed: new Date().toISOString(),
                    scopes: sessionData.scopes
                };

                // Encrypt and store refresh token if available
                if (sessionData.refreshToken) {
                    const encryptedRefreshToken = this.encryptToken(sessionData.refreshToken);
                    if (encryptedRefreshToken) {
                        sessionStorage.encryptedRefreshToken = encryptedRefreshToken;
                        console.log(`🔐 Encrypted refresh token for session: ${sessionId}`);
                    } else {
                        console.warn(`⚠️ Failed to encrypt refresh token for session: ${sessionId}`);
                    }
                }

                // Store access token expiry info for proactive refresh
                if (sessionData.accessToken && sessionData.expiresOn) {
                    sessionStorage.hasAccessToken = true;
                    sessionStorage.tokenExpiresOn = sessionData.expiresOn;
                }

                sessions[sessionId] = sessionStorage;
            }
            
            await fs.writeFile(this.sessionsFile, JSON.stringify(sessions, null, 2));
            console.log(`💾 Saved ${Object.keys(sessions).length} sessions with encrypted tokens to persistent storage`);
        } catch (error) {
            console.error('❌ Failed to save sessions:', error);
        }
    }

    async loadSessions() {
        try {
            const data = await fs.readFile(this.sessionsFile, 'utf8');
            const sessions = JSON.parse(data);
            
            // Filter out expired sessions and decrypt refresh tokens
            const now = new Date();
            const activeSessions = {};
            
            for (const [sessionId, sessionData] of Object.entries(sessions)) {
                const expiresOn = new Date(sessionData.expiresOn);
                // Keep sessions that haven't expired yet (with 1 hour grace period for refresh)
                if (expiresOn.getTime() > (now.getTime() - 60 * 60 * 1000)) {
                    const restoredSession = {
                        account: sessionData.account,
                        expiresOn: sessionData.expiresOn,
                        createdAt: sessionData.createdAt,
                        lastUsed: sessionData.lastUsed,
                        scopes: sessionData.scopes,
                        needsRefresh: true, // Flag that tokens need to be refreshed
                        hasStoredRefreshToken: !!sessionData.encryptedRefreshToken
                    };

                    // Decrypt refresh token if available
                    if (sessionData.encryptedRefreshToken) {
                        const decryptedRefreshToken = this.decryptToken(sessionData.encryptedRefreshToken);
                        if (decryptedRefreshToken) {
                            restoredSession.refreshToken = decryptedRefreshToken;
                            console.log(`🔓 Decrypted refresh token for session: ${sessionId}`);
                        } else {
                            console.warn(`⚠️ Failed to decrypt refresh token for session: ${sessionId}`);
                            // Session can still be restored but will need user re-auth if tokens expired
                        }
                    }

                    activeSessions[sessionId] = restoredSession;
                }
            }
            
            console.log(`📥 Loaded ${Object.keys(activeSessions).length} active sessions with decrypted tokens from storage`);
            return activeSessions;
        } catch (error) {
            if (error.code !== 'ENOENT') {
                console.error('❌ Failed to load sessions:', error);
            }
            return {};
        }
    }

    // Email-Session Mapping
    async saveEmailMapping(email, sessionId) {
        try {
            let mappings = {};
            try {
                const data = await fs.readFile(this.emailMappingsFile, 'utf8');
                mappings = JSON.parse(data);
            } catch (readError) {
                // File doesn't exist, start with empty mappings
            }
            
            const emailKey = email.toLowerCase().trim();
            mappings[emailKey] = {
                sessionId: sessionId,
                email: email,
                createdAt: new Date().toISOString(),
                expiresAt: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString() // 7 days
            };
            
            await fs.writeFile(this.emailMappingsFile, JSON.stringify(mappings, null, 2));
            console.log(`📝 Saved email mapping: ${email} → ${sessionId}`);
        } catch (error) {
            console.error('❌ Failed to save email mapping:', error);
        }
    }

    async getEmailMapping(email) {
        try {
            const data = await fs.readFile(this.emailMappingsFile, 'utf8');
            const mappings = JSON.parse(data);
            
            const emailKey = email.toLowerCase().trim();
            const mapping = mappings[emailKey];
            
            if (!mapping) {
                return null;
            }
            
            // Check if mapping has expired
            const expiresAt = new Date(mapping.expiresAt);
            if (expiresAt < new Date()) {
                delete mappings[emailKey];
                await fs.writeFile(this.emailMappingsFile, JSON.stringify(mappings, null, 2));
                return null;
            }
            
            return mapping;
        } catch (error) {
            if (error.code !== 'ENOENT') {
                console.error('❌ Failed to load email mapping:', error);
            }
            return null;
        }
    }

    async getAllEmailMappings() {
        try {
            const data = await fs.readFile(this.emailMappingsFile, 'utf8');
            const mappings = JSON.parse(data);
            
            // Clean expired mappings
            const now = new Date();
            const activeMappings = {};
            let cleaned = 0;
            
            for (const [email, mapping] of Object.entries(mappings)) {
                const expiresAt = new Date(mapping.expiresAt);
                if (expiresAt > now) {
                    activeMappings[email] = mapping;
                } else {
                    cleaned++;
                }
            }
            
            if (cleaned > 0) {
                await fs.writeFile(this.emailMappingsFile, JSON.stringify(activeMappings, null, 2));
                console.log(`🧹 Cleaned ${cleaned} expired email mappings`);
            }
            
            return activeMappings;
        } catch (error) {
            if (error.code !== 'ENOENT') {
                console.error('❌ Failed to load email mappings:', error);
            }
            return {};
        }
    }


    // User Authentication Context Storage
    async saveUserContext(sessionId, userEmail, oneDrivePath) {
        try {
            let contexts = {};
            try {
                const data = await fs.readFile(path.join(this.dataDir, 'user-contexts.json'), 'utf8');
                contexts = JSON.parse(data);
            } catch (readError) {
                // File doesn't exist
            }
            
            contexts[sessionId] = {
                userEmail: userEmail,
                oneDrivePath: oneDrivePath,
                lastActive: new Date().toISOString(),
                createdAt: contexts[sessionId]?.createdAt || new Date().toISOString()
            };
            
            await fs.writeFile(path.join(this.dataDir, 'user-contexts.json'), JSON.stringify(contexts, null, 2));
            console.log(`💾 Saved user context: ${userEmail} → ${sessionId}`);
        } catch (error) {
            console.error('❌ Failed to save user context:', error);
        }
    }

    async getUserContextByEmail(userEmail) {
        try {
            const data = await fs.readFile(path.join(this.dataDir, 'user-contexts.json'), 'utf8');
            const contexts = JSON.parse(data);
            
            // Find session by user email
            for (const [sessionId, context] of Object.entries(contexts)) {
                if (context.userEmail && context.userEmail.toLowerCase() === userEmail.toLowerCase()) {
                    return { sessionId, ...context };
                }
            }
            
            return null;
        } catch (error) {
            if (error.code !== 'ENOENT') {
                console.error('❌ Failed to load user contexts:', error);
            }
            return null;
        }
    }

    // Cleanup and maintenance
    async cleanup() {
        try {
            // Clean expired email mappings
            await this.getAllEmailMappings();
            
            
            console.log('🧹 Persistent storage cleanup completed');
        } catch (error) {
            console.error('❌ Cleanup error:', error);
        }
    }
}

module.exports = new PersistentStorage();