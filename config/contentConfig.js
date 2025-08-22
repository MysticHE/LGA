/**
 * Content Processing Configuration
 * 
 * Centralized configuration for PDF content processing, analysis, and caching.
 * Allows for easy tuning of processing parameters and behavior.
 */

const contentConfig = {
    // Processing Engine Settings
    processing: {
        // Maximum content length for processing (characters)
        maxContentLength: 50000,
        
        // Target output length for optimized content (characters)
        targetOutputLength: 2500,
        
        // Minimum segment length to consider for processing
        minSegmentLength: 30,
        
        // Maximum number of segments to process per file
        maxSegmentsPerFile: 20,
        
        // Content quality thresholds
        qualityThresholds: {
            minScore: 3,        // Minimum score to include segment
            highQuality: 7,     // Score threshold for high-quality content
            excellentQuality: 9 // Score threshold for excellent content
        }
    },

    // AI Analysis Settings
    ai: {
        // Default model for content analysis
        model: 'gpt-4o-mini',
        
        // Token limits for different analysis types
        tokenLimits: {
            summarize: 400,
            industry: 350,
            email: 300,
            fallback: 200
        },
        
        // Temperature settings for different analysis types
        temperature: {
            summarize: 0.3,     // More focused for summaries
            industry: 0.4,      // Balanced for industry-specific content
            email: 0.5,         // More creative for email content
            default: 0.3
        },
        
        // Retry settings for API calls
        retry: {
            maxRetries: 3,
            baseDelay: 1000,    // Base delay in milliseconds
            maxDelay: 5000      // Maximum delay in milliseconds
        },
        
        // Rate limiting settings
        rateLimiting: {
            requestsPerMinute: 50,
            batchSize: 3,
            delayBetweenBatches: 1000
        }
    },

    // Cache Configuration
    cache: {
        // Cache settings
        enabled: true,
        maxEntries: 1000,
        defaultTTL: 24 * 60 * 60 * 1000,    // 24 hours
        cleanupInterval: 60 * 60 * 1000,     // 1 hour
        maxContentSize: 50 * 1024,           // 50KB
        
        // Cache strategies
        strategies: {
            aggressive: {
                ttl: 7 * 24 * 60 * 60 * 1000,  // 7 days
                maxEntries: 2000
            },
            conservative: {
                ttl: 6 * 60 * 60 * 1000,       // 6 hours
                maxEntries: 500
            },
            development: {
                ttl: 30 * 60 * 1000,           // 30 minutes
                maxEntries: 100
            }
        },
        
        // Current strategy (can be overridden by environment)
        currentStrategy: process.env.NODE_ENV === 'production' ? 'aggressive' : 'development'
    },

    // Industry-specific settings
    industries: {
        // Default industry mappings for insurance products
        mappings: {
            'technology': {
                priority: ['cyber', 'professional', 'employment', 'directors'],
                keywords: ['data breach', 'software', 'IP', 'privacy', 'GDPR'],
                riskFactors: ['cyber attacks', 'data breaches', 'IP theft', 'employment disputes']
            },
            'manufacturing': {
                priority: ['property', 'liability', 'workers', 'commercial'],
                keywords: ['equipment', 'machinery', 'workplace safety', 'supply chain'],
                riskFactors: ['equipment damage', 'workplace injuries', 'product liability', 'supply disruption']
            },
            'healthcare': {
                priority: ['professional', 'cyber', 'employment', 'liability'],
                keywords: ['HIPAA', 'patient data', 'malpractice', 'compliance'],
                riskFactors: ['malpractice claims', 'data breaches', 'regulatory violations', 'employment issues']
            },
            'finance': {
                priority: ['cyber', 'professional', 'directors', 'employment'],
                keywords: ['financial data', 'compliance', 'fiduciary', 'regulation'],
                riskFactors: ['cyber attacks', 'regulatory fines', 'professional errors', 'director liability']
            },
            'retail': {
                priority: ['property', 'liability', 'commercial', 'employment'],
                keywords: ['customer safety', 'property damage', 'inventory', 'theft'],
                riskFactors: ['customer injuries', 'property damage', 'theft', 'employment disputes']
            },
            'construction': {
                priority: ['liability', 'workers', 'property', 'commercial'],
                keywords: ['workplace safety', 'equipment', 'project liability', 'subcontractors'],
                riskFactors: ['workplace injuries', 'property damage', 'project delays', 'liability claims']
            },
            'professional': {
                priority: ['professional', 'cyber', 'employment', 'directors'],
                keywords: ['professional errors', 'client data', 'compliance', 'expertise'],
                riskFactors: ['professional mistakes', 'data breaches', 'client disputes', 'regulatory issues']
            }
        },
        
        // Default fallback for unknown industries
        default: {
            priority: ['liability', 'property', 'professional', 'cyber'],
            keywords: ['business protection', 'risk management', 'compliance', 'coverage'],
            riskFactors: ['business risks', 'liability claims', 'property damage', 'cyber threats']
        }
    },

    // Content patterns and regex
    patterns: {
        // Content to remove or de-prioritize
        removePatterns: [
            /^page \d+.*$/gmi,
            /^\d+\s*$/gm,
            /^.*confidential.*$/gmi,
            /^.*proprietary.*$/gmi,
            /^.*copyright.*$/gmi,
            /^.*all rights reserved.*$/gmi,
            /^.*disclaimer.*$/gmi,
            /^.*terms and conditions.*$/gmi
        ],
        
        // Insurance product detection
        insuranceProducts: /(?:liability|property|cyber|professional|directors|officers|employment|commercial|general|auto|workers|compensation|umbrella|excess)\s+(?:insurance|coverage|policy|protection)/gi,
        
        // Business value terms
        businessValue: /(?:protect|reduce|mitigate|coverage|benefits|savings|efficiency|compliance|peace of mind|financial protection)/gi,
        
        // Section headers
        sectionHeaders: /^([A-Z][A-Z\s&]{2,}):?\s*$/gm,
        
        // Bullet points
        bulletPoints: /^[\s]*[•·‣⁃▪▫‣]\s*/gm,
        
        // Numbered lists
        numberedLists: /^[\s]*\d+\.\s*/gm
    },

    // Optimization settings
    optimization: {
        // Content compilation settings
        compilation: {
            maxSources: 5,              // Maximum number of source files to include
            maxSegmentsPerSource: 4,    // Maximum segments per source file
            prioritizeHighScore: true,   // Prioritize high-scoring segments
            preserveStructure: true,     // Maintain bullet points and formatting
            includeSourceAttribution: true // Include source file names
        },
        
        // Truncation settings
        truncation: {
            intelligentTruncation: true, // Truncate at sentence boundaries
            preserveKeyTerms: true,      // Keep important insurance terms
            minSentenceLength: 10,       // Minimum sentence length to preserve
            ellipsisIndicator: '...'     // Indicator for truncated content
        },
        
        // Quality scoring weights
        scoring: {
            sectionType: 0.3,           // Weight for section type (products, benefits, etc.)
            industryRelevance: 0.25,    // Weight for industry-specific content
            insuranceKeywords: 0.2,     // Weight for insurance keyword density
            contentQuality: 0.15,       // Weight for general content quality
            businessValue: 0.1          // Weight for business value terms
        }
    },

    // Feature flags
    features: {
        enableAISummarization: false,    // Disabled - AI Analyzer removed to preserve specific product details
        enableIndustryOptimization: true, // Enable industry-specific optimization
        enableCaching: true,             // Enable content caching
        enableBatchProcessing: true,     // Enable batch processing for multiple files
        enableFallbackProcessing: true,  // Enable fallback when AI processing fails
        enableContentValidation: true,   // Enable content validation before processing
        enablePerformanceMetrics: true   // Enable performance monitoring
    },

    // Logging and monitoring
    logging: {
        enabled: true,
        level: process.env.NODE_ENV === 'production' ? 'info' : 'debug',
        includeTimestamps: true,
        includePerformanceMetrics: true,
        
        // What to log
        logProcessingSteps: true,
        logCacheOperations: false,      // Can be verbose
        logAIApiCalls: true,
        logErrors: true
    },

    // Performance monitoring
    performance: {
        enableMetrics: true,
        thresholds: {
            processingTime: 5000,       // Max processing time in ms
            apiResponseTime: 3000,      // Max API response time in ms
            cacheHitRate: 0.7,          // Minimum acceptable cache hit rate
            compressionRatio: 0.6       // Minimum compression ratio
        },
        
        // Alerts
        alerts: {
            slowProcessing: true,
            lowCacheHitRate: true,
            apiErrors: true,
            highTokenUsage: true
        }
    }
};

/**
 * Get configuration for specific environment
 * @param {string} env - Environment name
 * @returns {Object} Environment-specific configuration
 */
function getEnvironmentConfig(env = process.env.NODE_ENV) {
    const envOverrides = {
        development: {
            cache: {
                ...contentConfig.cache,
                currentStrategy: 'development'
            },
            logging: {
                ...contentConfig.logging,
                level: 'debug',
                logCacheOperations: true
            },
            ai: {
                ...contentConfig.ai,
                rateLimiting: {
                    ...contentConfig.ai.rateLimiting,
                    requestsPerMinute: 20  // Lower rate limit for development
                }
            }
        },
        
        production: {
            cache: {
                ...contentConfig.cache,
                currentStrategy: 'aggressive'
            },
            logging: {
                ...contentConfig.logging,
                level: 'info',
                logCacheOperations: false
            },
            performance: {
                ...contentConfig.performance,
                alerts: {
                    ...contentConfig.performance.alerts,
                    // All alerts enabled in production
                }
            }
        },
        
        test: {
            cache: {
                ...contentConfig.cache,
                enabled: false  // Disable cache for testing
            },
            features: {
                ...contentConfig.features,
                enablePerformanceMetrics: false
            },
            logging: {
                ...contentConfig.logging,
                level: 'error'  // Minimal logging for tests
            }
        }
    };

    return {
        ...contentConfig,
        ...envOverrides[env] || {}
    };
}

/**
 * Validate configuration settings
 * @param {Object} config - Configuration to validate
 * @returns {Object} Validation result
 */
function validateConfig(config = contentConfig) {
    const errors = [];
    const warnings = [];

    // Validate required settings
    if (!config.ai?.model) {
        errors.push('AI model not specified');
    }

    if (config.processing?.maxContentLength < 1000) {
        warnings.push('Maximum content length is very low');
    }

    if (config.cache?.maxEntries < 100 && config.cache?.enabled) {
        warnings.push('Cache max entries is very low');
    }

    // Validate token limits
    Object.entries(config.ai?.tokenLimits || {}).forEach(([type, limit]) => {
        if (limit < 50 || limit > 1000) {
            warnings.push(`Token limit for ${type} may be inappropriate: ${limit}`);
        }
    });

    return {
        isValid: errors.length === 0,
        errors,
        warnings,
        configVersion: '1.0.0'
    };
}

module.exports = {
    contentConfig,
    getEnvironmentConfig,
    validateConfig
};