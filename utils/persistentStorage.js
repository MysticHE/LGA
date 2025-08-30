const fs = require('fs').promises;
const path = require('path');

/**
 * Persistent Storage System for Session and Email Tracking Data
 * Uses file-based storage that survives server restarts and deployments
 */

class PersistentStorage {
    constructor() {
        this.dataDir = path.join(process.cwd(), 'data');
        this.sessionsFile = path.join(this.dataDir, 'sessions.json');
        this.emailMappingsFile = path.join(this.dataDir, 'email-mappings.json');
        
        this.ensureDataDirectory();
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

    // Session Management
    async saveSessions(sessionsMap) {
        try {
            const sessions = {};
            for (const [sessionId, sessionData] of sessionsMap.entries()) {
                // Only store essential data, not sensitive tokens
                sessions[sessionId] = {
                    account: sessionData.account,
                    expiresOn: sessionData.expiresOn,
                    createdAt: sessionData.createdAt || new Date().toISOString(),
                    lastUsed: new Date().toISOString()
                };
            }
            
            await fs.writeFile(this.sessionsFile, JSON.stringify(sessions, null, 2));
            console.log(`💾 Saved ${Object.keys(sessions).length} sessions to persistent storage`);
        } catch (error) {
            console.error('❌ Failed to save sessions:', error);
        }
    }

    async loadSessions() {
        try {
            const data = await fs.readFile(this.sessionsFile, 'utf8');
            const sessions = JSON.parse(data);
            
            // Filter out expired sessions
            const now = new Date();
            const activeSessions = {};
            
            for (const [sessionId, sessionData] of Object.entries(sessions)) {
                const expiresOn = new Date(sessionData.expiresOn);
                // Keep sessions that haven't expired yet (with 1 hour grace period for refresh)
                if (expiresOn.getTime() > (now.getTime() - 60 * 60 * 1000)) {
                    activeSessions[sessionId] = sessionData;
                }
            }
            
            console.log(`📥 Loaded ${Object.keys(activeSessions).length} active sessions from storage`);
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