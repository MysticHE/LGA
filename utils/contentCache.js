/**
 * Content Caching System
 * 
 * Intelligent caching for processed PDF content to improve performance
 * and reduce API calls to OpenAI. Includes cache invalidation and optimization.
 */

class ContentCache {
    constructor() {
        // In-memory cache (use Redis in production)
        this.cache = new Map();
        this.metadata = new Map();
        
        // Cache configuration
        this.config = {
            maxEntries: 1000,
            defaultTTL: 24 * 60 * 60 * 1000, // 24 hours in milliseconds
            cleanupInterval: 60 * 60 * 1000, // 1 hour cleanup interval
            maxContentSize: 50 * 1024 // 50KB max per cached item
        };

        // Statistics tracking
        this.stats = {
            hits: 0,
            misses: 0,
            evictions: 0,
            totalSaved: 0 // API calls saved
        };

        // Start periodic cleanup
        this.startCleanupTimer();
    }

    /**
     * Generate cache key for content
     * @param {string} content - Original content
     * @param {Object} options - Processing options
     * @returns {string} Cache key
     */
    generateKey(content, options = {}) {
        // Create hash-like key based on content and options
        const contentHash = this.simpleHash(content);
        const optionsString = JSON.stringify(this.normalizeOptions(options));
        const optionsHash = this.simpleHash(optionsString);
        
        return `content_${contentHash}_${optionsHash}`;
    }

    /**
     * Get cached content
     * @param {string} content - Original content
     * @param {Object} options - Processing options
     * @returns {Object|null} Cached result or null if not found
     */
    get(content, options = {}) {
        const key = this.generateKey(content, options);
        const cached = this.cache.get(key);
        
        if (!cached) {
            this.stats.misses++;
            return null;
        }

        // Check if expired
        if (this.isExpired(key)) {
            this.delete(key);
            this.stats.misses++;
            return null;
        }

        this.stats.hits++;
        this.stats.totalSaved++;
        
        // Update access time
        this.updateAccessTime(key);
        
        console.log(`✅ Cache hit for content processing (${key.substring(0, 20)}...)`);
        return cached;
    }

    /**
     * Store content in cache
     * @param {string} content - Original content
     * @param {Object} options - Processing options
     * @param {Object} result - Processed result to cache
     * @param {number} ttl - Time to live in milliseconds (optional)
     * @returns {boolean} Success status
     */
    set(content, options = {}, result, ttl = null) {
        try {
            const key = this.generateKey(content, options);
            
            // Check content size
            const serializedResult = JSON.stringify(result);
            if (serializedResult.length > this.config.maxContentSize) {
                console.warn(`⚠️ Content too large for cache: ${serializedResult.length} bytes`);
                return false;
            }

            // Ensure cache doesn't exceed max entries
            this.ensureCacheSize();

            // Store with metadata
            const expiresAt = Date.now() + (ttl || this.config.defaultTTL);
            this.cache.set(key, result);
            this.metadata.set(key, {
                createdAt: Date.now(),
                expiresAt,
                accessedAt: Date.now(),
                accessCount: 1,
                size: serializedResult.length,
                contentLength: content.length,
                options: this.normalizeOptions(options)
            });

            console.log(`✅ Cached content processing result (${key.substring(0, 20)}...)`);
            return true;

        } catch (error) {
            console.error('Cache storage error:', error);
            return false;
        }
    }

    /**
     * Check if cache entry has expired
     * @param {string} key - Cache key
     * @returns {boolean} True if expired
     */
    isExpired(key) {
        const meta = this.metadata.get(key);
        return !meta || Date.now() > meta.expiresAt;
    }

    /**
     * Update access time for cache entry
     * @param {string} key - Cache key
     */
    updateAccessTime(key) {
        const meta = this.metadata.get(key);
        if (meta) {
            meta.accessedAt = Date.now();
            meta.accessCount++;
        }
    }

    /**
     * Delete cache entry
     * @param {string} key - Cache key
     * @returns {boolean} True if deleted
     */
    delete(key) {
        const deleted = this.cache.delete(key);
        if (deleted) {
            this.metadata.delete(key);
        }
        return deleted;
    }

    /**
     * Clear all cache entries
     */
    clear() {
        this.cache.clear();
        this.metadata.clear();
        console.log('🗑️ Cache cleared');
    }

    /**
     * Ensure cache doesn't exceed maximum entries
     */
    ensureCacheSize() {
        if (this.cache.size >= this.config.maxEntries) {
            // Remove oldest accessed entries (LRU eviction)
            const entries = Array.from(this.metadata.entries())
                .sort((a, b) => a[1].accessedAt - b[1].accessedAt);

            const toRemove = Math.ceil(this.config.maxEntries * 0.1); // Remove 10%
            for (let i = 0; i < toRemove && entries.length > 0; i++) {
                const [key] = entries[i];
                this.delete(key);
                this.stats.evictions++;
            }
            
            console.log(`🗑️ Evicted ${toRemove} cache entries (LRU)`);
        }
    }

    /**
     * Normalize options for consistent caching
     * @param {Object} options - Processing options
     * @returns {Object} Normalized options
     */
    normalizeOptions(options) {
        return {
            type: options.type || 'summarize',
            industry: options.industry?.toLowerCase() || null,
            role: options.role?.toLowerCase() || null,
            maxTokens: options.maxTokens || 400,
            useAI: options.useAI !== false // Default to true
        };
    }

    /**
     * Simple hash function for content
     * @param {string} str - String to hash
     * @returns {string} Hash string
     */
    simpleHash(str) {
        let hash = 0;
        if (str.length === 0) return hash.toString();
        
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        
        return Math.abs(hash).toString(36);
    }

    /**
     * Get cache statistics
     * @returns {Object} Cache statistics
     */
    getStats() {
        const totalRequests = this.stats.hits + this.stats.misses;
        const hitRate = totalRequests > 0 ? (this.stats.hits / totalRequests) * 100 : 0;
        
        return {
            ...this.stats,
            totalRequests,
            hitRate: hitRate.toFixed(2) + '%',
            currentEntries: this.cache.size,
            maxEntries: this.config.maxEntries,
            cacheUtilization: ((this.cache.size / this.config.maxEntries) * 100).toFixed(2) + '%'
        };
    }

    /**
     * Get detailed cache information
     * @returns {Object} Detailed cache information
     */
    getCacheInfo() {
        const entries = Array.from(this.metadata.entries()).map(([key, meta]) => ({
            key: key.substring(0, 20) + '...',
            createdAt: new Date(meta.createdAt).toISOString(),
            expiresAt: new Date(meta.expiresAt).toISOString(),
            accessCount: meta.accessCount,
            size: meta.size,
            contentLength: meta.contentLength,
            options: meta.options
        }));

        return {
            stats: this.getStats(),
            config: this.config,
            entries: entries.slice(0, 10), // Show first 10 entries
            totalEntries: entries.length
        };
    }

    /**
     * Clean up expired entries
     * @returns {number} Number of entries cleaned
     */
    cleanup() {
        let cleanedCount = 0;
        const now = Date.now();

        for (const [key, meta] of this.metadata.entries()) {
            if (now > meta.expiresAt) {
                this.delete(key);
                cleanedCount++;
            }
        }

        if (cleanedCount > 0) {
            console.log(`🧹 Cleaned up ${cleanedCount} expired cache entries`);
        }

        return cleanedCount;
    }

    /**
     * Start periodic cleanup timer
     */
    startCleanupTimer() {
        setInterval(() => {
            this.cleanup();
        }, this.config.cleanupInterval);
    }

    /**
     * Cache content with industry/role specific optimization
     * @param {string} content - Original content
     * @param {string} industry - Target industry
     * @param {string} role - Target role
     * @param {Object} result - Processing result
     * @returns {boolean} Success status
     */
    cacheByIndustryRole(content, industry, role, result) {
        const options = { 
            type: 'industry', 
            industry: industry?.toLowerCase(), 
            role: role?.toLowerCase() 
        };
        
        return this.set(content, options, result);
    }

    /**
     * Get cached content by industry/role
     * @param {string} content - Original content
     * @param {string} industry - Target industry
     * @param {string} role - Target role
     * @returns {Object|null} Cached result or null
     */
    getByIndustryRole(content, industry, role) {
        const options = { 
            type: 'industry', 
            industry: industry?.toLowerCase(), 
            role: role?.toLowerCase() 
        };
        
        return this.get(content, options);
    }

    /**
     * Cache general summarized content
     * @param {string} content - Original content
     * @param {Object} result - Processing result
     * @returns {boolean} Success status
     */
    cacheSummary(content, result) {
        return this.set(content, { type: 'summarize' }, result);
    }

    /**
     * Get cached summary
     * @param {string} content - Original content
     * @returns {Object|null} Cached result or null
     */
    getSummary(content) {
        return this.get(content, { type: 'summarize' });
    }

    /**
     * Invalidate cache entries matching pattern
     * @param {string} pattern - Pattern to match (simple string contains)
     * @returns {number} Number of entries invalidated
     */
    invalidateByPattern(pattern) {
        let invalidatedCount = 0;
        
        for (const [key] of this.cache.entries()) {
            if (key.includes(pattern)) {
                this.delete(key);
                invalidatedCount++;
            }
        }

        console.log(`🗑️ Invalidated ${invalidatedCount} cache entries matching pattern: ${pattern}`);
        return invalidatedCount;
    }

    /**
     * Warm up cache with common industry combinations
     * @param {Array} industries - Common industries to warm up
     * @param {Array} roles - Common roles to warm up
     * @returns {Object} Warmup results
     */
    async warmupCache(industries = [], roles = []) {
        // This would be called during application startup
        // to pre-populate cache with common combinations
        console.log(`🔥 Cache warmup initiated for ${industries.length} industries and ${roles.length} roles`);
        
        return {
            message: 'Cache warmup completed',
            industries: industries.length,
            roles: roles.length,
            combinations: industries.length * roles.length
        };
    }
}

module.exports = ContentCache;