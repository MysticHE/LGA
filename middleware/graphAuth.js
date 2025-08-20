const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

/**
 * Microsoft Graph Authentication Middleware
 * Configured for Render deployment with environment variables
 */

class GraphAuthProvider {
    constructor() {
        // Render environment variables
        this.tenantId = process.env.AZURE_TENANT_ID;
        this.clientId = process.env.AZURE_CLIENT_ID;
        this.clientSecret = process.env.AZURE_CLIENT_SECRET;
        this.redirectUri = process.env.AZURE_REDIRECT_URI;
        this.serviceAccountEmail = process.env.AZURE_SERVICE_ACCOUNT_EMAIL;
        
        this.client = null;
        this.authProvider = null;
        
        this.validateConfig();
        this.initializeClient();
    }

    validateConfig() {
        const required = ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET'];
        const missing = required.filter(key => !process.env[key]);
        
        if (missing.length > 0) {
            throw new Error(`Missing required environment variables for Microsoft Graph: ${missing.join(', ')}`);
        }

        if (!this.serviceAccountEmail) {
            console.log('⚠️ AZURE_SERVICE_ACCOUNT_EMAIL not set - will use organization admin for operations');
        }
        
        console.log('✅ Microsoft Graph configuration validated');
    }

    initializeClient() {
        try {
            // Create Azure Identity credential
            const credential = new ClientSecretCredential(
                this.tenantId,
                this.clientId,
                this.clientSecret
            );

            // Create authentication provider
            this.authProvider = new TokenCredentialAuthenticationProvider(credential, {
                scopes: [
                    'https://graph.microsoft.com/.default'
                ]
            });

            // Initialize Microsoft Graph client
            this.client = Client.initWithMiddleware({
                authProvider: this.authProvider,
                debugLogging: process.env.NODE_ENV === 'development'
            });

            console.log('✅ Microsoft Graph client initialized successfully');
        } catch (error) {
            console.error('❌ Failed to initialize Microsoft Graph client:', error.message);
            throw error;
        }
    }

    getClient() {
        if (!this.client) {
            throw new Error('Microsoft Graph client not initialized');
        }
        return this.client;
    }

    getAuthProvider() {
        return this.authProvider;
    }

    // Test the connection
    async testConnection() {
        try {
            const client = this.getClient();
            // Use application-level endpoint instead of /me
            const response = await client.api('/users').top(1).get();
            console.log('✅ Microsoft Graph connection test successful');
            return { success: true, user: 'Application Access', users: response.value.length };
        } catch (error) {
            console.error('❌ Microsoft Graph connection test failed:', error.message);
            return { success: false, error: error.message };
        }
    }

    // Get service account user ID for operations
    async getServiceAccountUserId() {
        if (!this.serviceAccountEmail) {
            // If no specific service account, get the first admin user
            try {
                const users = await this.client.api('/users').filter("assignedLicenses/any(x:x/skuId ne null)").top(1).get();
                if (users.value.length > 0) {
                    return users.value[0].id;
                }
                throw new Error('No licensed users found in organization');
            } catch (error) {
                console.error('Failed to get service account user:', error.message);
                throw error;
            }
        }
        
        try {
            const user = await this.client.api(`/users/${this.serviceAccountEmail}`).get();
            return user.id;
        } catch (error) {
            console.error(`Failed to get user ID for ${this.serviceAccountEmail}:`, error.message);
            throw error;
        }
    }

    // Get application-only access token for webhook notifications
    async getAppAccessToken() {
        try {
            const credential = new ClientSecretCredential(
                this.tenantId,
                this.clientId,
                this.clientSecret
            );
            
            const token = await credential.getToken(['https://graph.microsoft.com/.default']);
            return token.token;
        } catch (error) {
            console.error('❌ Failed to get app access token:', error.message);
            throw error;
        }
    }
}

// Create singleton instance
let graphAuthInstance = null;

function getGraphAuthProvider() {
    if (!graphAuthInstance) {
        graphAuthInstance = new GraphAuthProvider();
    }
    return graphAuthInstance;
}

// Middleware function for routes
function requireGraphAuth(req, res, next) {
    try {
        const graphAuth = getGraphAuthProvider();
        req.graphClient = graphAuth.getClient();
        req.graphAuth = graphAuth;
        next();
    } catch (error) {
        console.error('Graph auth middleware error:', error);
        res.status(500).json({
            error: 'Microsoft Graph Authentication Error',
            message: 'Failed to authenticate with Microsoft Graph',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
}

module.exports = {
    GraphAuthProvider,
    getGraphAuthProvider,
    requireGraphAuth
};