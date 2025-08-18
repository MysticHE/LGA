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