const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
require('dotenv').config();

const apolloRoutes = require('./routes/apollo');
const leadRoutes = require('./routes/leads');
const microsoftGraphRoutes = require('./routes/microsoft-graph');
const emailAutomationRoutes = require('./routes/email-automation');
const authRoutes = require('./routes/auth');
const { rateLimiter } = require('./middleware/rateLimiter');

const app = express();
const PORT = process.env.PORT || 3000;

// Security middleware
app.use(helmet({
    contentSecurityPolicy: {
        directives: {
            defaultSrc: ["'self'"],
            styleSrc: ["'self'", "'unsafe-inline'"],
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

// Rate limiting
app.use(rateLimiter);

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));

// API Routes
app.use('/api/apollo', apolloRoutes);
app.use('/api/leads', leadRoutes);
app.use('/api/microsoft-graph', microsoftGraphRoutes);
app.use('/api/email', emailAutomationRoutes);
app.use('/auth', authRoutes);

// Serve the main application
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'lead-generator.html'));
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
    
    if (err.type === 'rate_limit') {
        return res.status(429).json({
            error: 'Rate limit exceeded',
            message: 'Too many requests. Please try again later.'
        });
    }
    
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
app.listen(PORT, () => {
    console.log(`🚀 Lead Generation Server running on port ${PORT}`);
    console.log(`📊 Environment: ${process.env.NODE_ENV || 'development'}`);
    console.log(`🌐 Access: http://localhost:${PORT}`);
    console.log(`✅ Server ready - v1.1.0`);
    
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