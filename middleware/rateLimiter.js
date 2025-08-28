const { RateLimiterMemory } = require('rate-limiter-flexible');

// Create rate limiter instance
const rateLimiter = new RateLimiterMemory({
    keyPrefix: 'lead_gen',
    points: parseInt(process.env.MAX_REQUESTS_PER_MINUTE) || 10, // Number of requests
    duration: 60, // Per 60 seconds
    blockDuration: 60, // Block for 60 seconds if limit exceeded
});

// Rate limiting middleware
const rateLimiterMiddleware = async (req, res, next) => {
    try {
        // Exempt job status polling endpoints from rate limiting
        if (req.path.includes('/job-status/') || req.path.includes('/job-result/')) {
            return next();
        }
        
        // Exempt authentication endpoints from rate limiting
        if (req.path.startsWith('/auth/')) {
            return next();
        }
        
        // Exempt Microsoft Graph test endpoint from rate limiting
        if (req.path === '/api/microsoft-graph/test') {
            return next();
        }
        
        // Exempt email automation status/stats endpoints from rate limiting
        if (req.path.includes('/api/email-automation/master-list/stats') || 
            req.path.includes('/api/email-automation/templates')) {
            return next();
        }
        
        // Exempt static files from rate limiting
        const staticFileExtensions = ['.css', '.js', '.png', '.jpg', '.jpeg', '.gif', '.ico', '.svg', '.woff', '.woff2', '.ttf', '.eot'];
        const isStaticFile = staticFileExtensions.some(ext => req.path.toLowerCase().endsWith(ext));
        if (isStaticFile || req.path === '/favicon.ico') {
            return next();
        }
        
        // Use IP address as the key
        const key = req.ip || req.connection.remoteAddress;
        
        await rateLimiter.consume(key);
        next();
    } catch (rejRes) {
        // Rate limit exceeded
        const secs = Math.round(rejRes.msBeforeNext / 1000) || 1;
        
        res.set('Retry-After', String(secs));
        res.status(429).json({
            error: 'Rate limit exceeded',
            message: `Too many requests. Try again in ${secs} seconds.`,
            retryAfter: secs
        });
    }
};

module.exports = {
    rateLimiter: rateLimiterMiddleware,
    rateLimiterInstance: rateLimiter
};