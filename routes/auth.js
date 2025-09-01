const express = require('express');
const { getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const router = express.Router();

/**
 * Microsoft 365 Authentication Routes
 * Handles OAuth2 delegated authentication flow
 */

// Generate session ID helper
function generateSessionId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// Initiate Microsoft 365 login
router.get('/login', async (req, res) => {
    try {
        const authProvider = getDelegatedAuthProvider();
        const sessionId = generateSessionId();
        const redirectUrl = req.query.redirect || '/';
        
        // Store redirect URL in session (in production, use proper session store)
        if (!global.authSessions) {
            global.authSessions = new Map();
        }
        global.authSessions.set(sessionId, { redirectUrl });
        
        const authUrl = await authProvider.getAuthUrl(sessionId);
        
        console.log(`🔑 Initiating Microsoft 365 login for session: ${sessionId}`);
        
        res.json({
            success: true,
            authUrl: authUrl,
            sessionId: sessionId,
            message: 'Redirect user to authUrl for Microsoft 365 authentication'
        });

    } catch (error) {
        console.error('Login initiation error:', error);
        res.status(500).json({
            success: false,
            error: 'Login Error',
            message: 'Failed to initiate Microsoft 365 login',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Handle OAuth callback from Microsoft
router.get('/callback', async (req, res) => {
    try {
        const { code, state: sessionId, error, error_description } = req.query;
        
        if (error) {
            console.error('OAuth error:', error, error_description);
            return res.status(400).send(`
                <html>
                <body>
                    <h2>Authentication Failed</h2>
                    <p>Error: ${error}</p>
                    <p>Description: ${error_description || 'Unknown error'}</p>
                    <a href="/">Return to Application</a>
                </body>
                </html>
            `);
        }

        if (!code || !sessionId) {
            return res.status(400).send(`
                <html>
                <body>
                    <h2>Authentication Failed</h2>
                    <p>Missing authorization code or session ID</p>
                    <a href="/">Return to Application</a>
                </body>
                </html>
            `);
        }

        const authProvider = getDelegatedAuthProvider();
        const result = await authProvider.handleCallback(code, sessionId);
        
        if (result.success) {
            // Get redirect URL from session
            const sessionData = global.authSessions?.get(sessionId);
            const baseRedirectUrl = sessionData?.redirectUrl || '/';
            
            // Add sessionId as URL parameter
            const redirectUrl = baseRedirectUrl.includes('?') 
                ? `${baseRedirectUrl}&sessionId=${sessionId}`
                : `${baseRedirectUrl}?sessionId=${sessionId}`;
            
            console.log(`✅ Authentication successful for session: ${sessionId}`);
            
            // Session is now persistent and will enable 24/7 background operation
            console.log(`🔄 Session ${sessionId} configured for continuous background operation`);
            
            res.send(`
                <html>
                <head>
                    <title>Authentication Successful</title>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <style>
                        body { font-family: Arial, sans-serif; text-align: center; padding: 50px; background: #f5f5f5; }
                        .container { background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); max-width: 500px; margin: 0 auto; }
                        .success { color: #28a745; font-size: 24px; margin-bottom: 20px; }
                        .info { color: #666; margin-bottom: 30px; }
                        .btn { background: #007bff; color: white; padding: 12px 24px; border: none; border-radius: 5px; text-decoration: none; display: inline-block; }
                        .btn:hover { background: #0056b3; }
                    </style>
                </head>
                <body>
                    <div class="container">
                        <div class="success">✅ Authentication Successful!</div>
                        <div class="info">
                            <p>Welcome, <strong>${result.user}</strong></p>
                            <p>You have successfully authenticated with Microsoft 365.</p>
                            <p>Session ID: <code>${sessionId}</code></p>
                        </div>
                        <a href="${redirectUrl}" class="btn">Continue to Application</a>
                    </div>
                    <script>
                        // Auto-close if opened in popup
                        if (window.opener) {
                            window.opener.postMessage({
                                type: 'auth_success',
                                sessionId: '${sessionId}',
                                user: '${result.user}'
                            }, '*');
                            window.close();
                        }
                        
                        // Auto-redirect after 3 seconds
                        setTimeout(() => {
                            window.location.href = '${redirectUrl}';
                        }, 3000);
                    </script>
                </body>
                </html>
            `);
        } else {
            res.status(400).send(`
                <html>
                <body>
                    <h2>Authentication Failed</h2>
                    <p>Error: ${result.error}</p>
                    <a href="/auth/login">Try Again</a>
                </body>
                </html>
            `);
        }

    } catch (error) {
        console.error('Authentication callback error:', error);
        res.status(500).send(`
            <html>
            <body>
                <h2>Authentication Error</h2>
                <p>An unexpected error occurred during authentication.</p>
                <p>Please try again or contact support.</p>
                <a href="/">Return to Application</a>
            </body>
            </html>
        `);
    }
});

// Check authentication status
router.get('/status', async (req, res) => {
    try {
        const sessionId = req.headers['x-session-id'] || req.query.sessionId;
        
        if (!sessionId) {
            return res.json({
                authenticated: false,
                message: 'No session ID provided'
            });
        }

        const authProvider = getDelegatedAuthProvider();
        const isAuth = authProvider.isAuthenticated(sessionId);
        
        if (isAuth) {
            const userInfo = authProvider.getUserInfo(sessionId);
            res.json({
                authenticated: true,
                user: userInfo.username,
                name: userInfo.name,
                sessionId: sessionId
            });
        } else {
            res.json({
                authenticated: false,
                message: 'Session not authenticated'
            });
        }

    } catch (error) {
        console.error('Auth status error:', error);
        res.status(500).json({
            authenticated: false,
            error: error.message
        });
    }
});

// Logout user
router.post('/logout', (req, res) => {
    try {
        const sessionId = req.headers['x-session-id'] || req.body.sessionId;
        
        if (sessionId) {
            const authProvider = getDelegatedAuthProvider();
            authProvider.logout(sessionId);
        }

        res.json({
            success: true,
            message: 'Logged out successfully'
        });

    } catch (error) {
        console.error('Logout error:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// Test Microsoft Graph connection for authenticated user
router.get('/test-graph', async (req, res) => {
    try {
        const sessionId = req.headers['x-session-id'] || req.query.sessionId;
        
        if (!sessionId) {
            return res.status(401).json({
                success: false,
                error: 'Session ID required'
            });
        }

        const authProvider = getDelegatedAuthProvider();
        const result = await authProvider.testConnection(sessionId);
        
        if (result.success) {
            res.json({
                success: true,
                message: 'Microsoft Graph connection successful',
                user: result.user,
                email: result.email,
                authType: 'Delegated'
            });
        } else {
            res.status(401).json({
                success: false,
                error: result.error,
                message: 'Microsoft Graph connection failed'
            });
        }

    } catch (error) {
        console.error('Graph test error:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

module.exports = router;