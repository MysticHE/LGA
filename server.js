const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
require('dotenv').config();

const apolloRoutes = require('./routes/apollo');
const leadRoutes = require('./routes/leads');
const microsoftGraphRoutes = require('./routes/microsoft-graph');
const emailAutomationRoutes = require('./routes/email-automation');
const emailTemplatesRoutes = require('./routes/email-templates');
const emailTrackingRoutes = require('./routes/email-tracking');
const emailSchedulerRoutes = require('./routes/email-scheduler');
const emailDelayTestRoutes = require('./routes/email-delay-test');
const emailBounceRoutes = require('./routes/email-bounce');
const authRoutes = require('./routes/auth');
const campaignStatusRoutes = require('./routes/campaign-status');
const ProcessSingleton = require('./utils/processSingleton');

// 🔒 PROCESS SINGLETON: Prevent multiple server instances
const singleton = new ProcessSingleton('lga-server');

// Check if another instance is already running
if (singleton.isAnotherInstanceRunning()) {
    console.error('❌ Another instance of the server is already running!');
    console.error('📝 This prevents campaign conflicts and duplicate email sends.');
    console.error('🚫 Exiting to avoid conflicts.');
    
    const runningInfo = singleton.getRunningInstanceInfo();
    if (runningInfo) {
        console.error(`📍 Running instance: PID ${runningInfo.pid}, Port ${runningInfo.port}`);
        console.error(`⏰ Started: ${new Date(runningInfo.startTime).toLocaleString()}`);
    }
    
    console.error('\n💡 To start a new instance:');
    console.error('   1. Stop the existing server (Ctrl+C)');
    console.error('   2. Wait a moment for cleanup');
    console.error('   3. Restart with npm start');
    
    process.exit(1);
}

// Create lock for this instance
singleton.setupExitHandlers();

// Check Azure configuration before initializing email services
const requiredAzureVars = ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET'];
const missingAzureVars = requiredAzureVars.filter(envVar => !process.env[envVar]);

if (missingAzureVars.length > 0) {
    console.warn('⚠️  AZURE CREDENTIALS MISSING - Email automation disabled');
    console.warn(`📝 Missing: ${missingAzureVars.join(', ')}`);
    console.warn('📋 Please check your Render environment variables for Azure credentials');
} else {
    // Initialize email scheduler only if Azure credentials are available
    try {
        const emailScheduler = require('./jobs/emailScheduler');
        console.log('📅 Starting email scheduler...');
        emailScheduler.start();
        console.log('✅ Email automation services initialized');
    } catch (error) {
        console.error('❌ Failed to initialize email services:', error.message);
        console.warn('📝 Email automation will be disabled');
    }
}

const app = express();
const PORT = process.env.PORT || 3000;

// Security middleware
app.use(helmet({
    contentSecurityPolicy: {
        directives: {
            defaultSrc: ["'self'"],
            styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
            scriptSrc: ["'self'", "'unsafe-inline'", "https://unpkg.com"],
            scriptSrcAttr: ["'unsafe-inline'"], // Allow inline event handlers temporarily
            imgSrc: ["'self'", "data:", "https:"],
            connectSrc: ["'self'", "https://api.apify.com", "https://api.openai.com", "https://graph.microsoft.com", "https://login.microsoftonline.com"]
        }
    }
}));

// CORS configuration
app.use(cors({
    origin: process.env.NODE_ENV === 'production' 
        ? ['https://yourdomain.com'] 
        : ['http://localhost:3000', 'http://127.0.0.1:3000'],
    credentials: true
}));

// Body parsing middleware
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Serve static files (before rate limiting)
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname)); // Serve files from root directory (for styles, etc.)

// No rate limiting - removed for simplified system

// API Routes
app.use('/api/apollo', apolloRoutes);
app.use('/api/leads', leadRoutes);
app.use('/api/microsoft-graph', microsoftGraphRoutes);
app.use('/api/email-automation', emailAutomationRoutes);
app.use('/api/email-automation/templates', emailTemplatesRoutes);
app.use('/api/email', emailTrackingRoutes);
app.use('/api/email-scheduler', emailSchedulerRoutes);  // Fixed: Use correct path
app.use('/api/email-delay', emailDelayTestRoutes);
app.use('/api/email-bounce', emailBounceRoutes);
app.use('/api/campaign-status', campaignStatusRoutes);
app.use('/auth', authRoutes);

// Serve the main application
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'lead-generator.html'));
});

// Serve email automation page
app.get('/email-automation', (req, res) => {
    res.sendFile(path.join(__dirname, 'email-automation.html'));
});

// Serve prompt editor page
app.get('/prompt-editor', (req, res) => {
    res.sendFile(path.join(__dirname, 'prompt-editor.html'));
});

// Favicon route to prevent 404 errors
app.get('/favicon.ico', (req, res) => {
    res.status(204).send();
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.status(200).json({
        status: 'OK',
        timestamp: new Date().toISOString(),
        version: '1.0.0'
    });
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Error:', err);
    
    
    if (err.name === 'ValidationError') {
        return res.status(400).json({
            error: 'Validation Error',
            message: err.message
        });
    }
    
    res.status(500).json({
        error: 'Internal Server Error',
        message: process.env.NODE_ENV === 'development' ? err.message : 'Something went wrong'
    });
});

// 404 handler
app.use('*', (req, res) => {
    res.status(404).json({
        error: 'Not Found',
        message: 'The requested resource was not found'
    });
});

// Start server
app.listen(PORT, async () => {
    // Create singleton lock after successful startup
    singleton.createLock(PORT);
    
    console.log(`🚀 Lead Generation Server running on port ${PORT}`);
    console.log(`📊 Environment: ${process.env.NODE_ENV || 'development'}`);
    console.log(`🌐 Access: http://localhost:${PORT}`);
    console.log(`🔐 Process singleton protection: ENABLED`);
    console.log(`✅ Server ready - v1.1.0 (Duplicate Prevention Edition)`);
    
    // Initialize persistent storage and session recovery
    try {
        const persistentStorage = require('./utils/persistentStorage');
        await persistentStorage.cleanup();
        console.log('💾 Persistent storage initialized and cleaned');
    } catch (storageError) {
        console.error('❌ Failed to initialize persistent storage:', storageError);
    }
    
    // Check for required environment variables
    const requiredEnvVars = ['APIFY_API_TOKEN', 'OPENAI_API_KEY'];
    const optionalEnvVars = ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET'];
    const missing = requiredEnvVars.filter(envVar => !process.env[envVar]);
    const missingOptional = optionalEnvVars.filter(envVar => !process.env[envVar]);
    
    if (missing.length > 0) {
        console.warn(`⚠️  Missing required environment variables: ${missing.join(', ')}`);
        console.warn('📝 Please check your .env file');
    }
    
    if (missingOptional.length > 0) {
        console.warn(`ℹ️  Optional Microsoft Graph variables not set: ${missingOptional.join(', ')}`);
        console.warn('📝 OneDrive and Email automation features will be disabled');
    } else {
        console.log('🔗 Microsoft Graph integration enabled');
    }
});

// Graceful shutdown
process.on('SIGTERM', () => {
    console.log('📴 SIGTERM received. Shutting down gracefully...');
    process.exit(0);
});

process.on('SIGINT', () => {
    console.log('📴 SIGINT received. Shutting down gracefully...');
    process.exit(0);
});