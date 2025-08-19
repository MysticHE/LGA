const express = require('express');
const axios = require('axios');
const router = express.Router();

// Apollo URL generation endpoint
router.post('/generate-url', async (req, res) => {
    try {
        const { jobTitles, companySizes } = req.body;

        // Validation
        if (!jobTitles || !Array.isArray(jobTitles) || jobTitles.length === 0) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Job titles are required and must be a non-empty array'
            });
        }

        if (!companySizes || !Array.isArray(companySizes) || companySizes.length === 0) {
            return res.status(400).json({
                error: 'Validation Error', 
                message: 'Company sizes are required and must be a non-empty array'
            });
        }

        // Generate Apollo URL (same logic as frontend)
        const baseUrl = "https://app.apollo.io/#/people?page=1";
        
        const defaultFilters = [
            "contactEmailStatusV2[]=verified",
            "existFields[]=person_title_normalized",
            "existFields[]=organization_domain", 
            "personLocations[]=Singapore",
            "personLocations[]=Singapore%2C%20Singapore",
            "sortAscending=true",
            "sortByField=sanitized_organization_name_unanalyzed"
        ];

        const titleFilters = jobTitles.map(title => 
            `personTitles[]=${encodeURIComponent(title)}`
        );

        const sizeFilters = companySizes.map(size => {
            const normalized = size.replace("-", ",");
            return `organizationNumEmployeesRanges[]=${encodeURIComponent(normalized)}`;
        });

        const allFilters = [...defaultFilters, ...titleFilters, ...sizeFilters];
        const apolloUrl = `${baseUrl}&${allFilters.join("&")}`;

        res.json({
            success: true,
            apolloUrl,
            filters: {
                jobTitles,
                companySizes,
                location: 'Singapore'
            }
        });

    } catch (error) {
        console.error('Apollo URL generation error:', error);
        res.status(500).json({
            error: 'Server Error',
            message: 'Failed to generate Apollo URL'
        });
    }
});

// Apollo lead scraping endpoint with streaming support
router.post('/scrape-leads-stream', async (req, res) => {
    try {
        const { apolloUrl, maxRecords = 500, stream = false } = req.body;

        // Validation
        if (!apolloUrl) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Apollo URL is required'
            });
        }

        if (!apolloUrl.includes('apollo.io')) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Invalid Apollo URL'
            });
        }

        // Check if Apify API token is configured
        if (!process.env.APIFY_API_TOKEN) {
            return res.status(500).json({
                error: 'Configuration Error',
                message: 'Apify API token not configured'
            });
        }

        // Handle unlimited vs limited records
        const maxLimit = parseInt(process.env.MAX_LEADS_PER_REQUEST) || 2000;
        let recordLimit;
        
        if (maxRecords === 0) {
            recordLimit = maxLimit;
        } else {
            recordLimit = Math.min(parseInt(maxRecords) || 500, maxLimit);
        }

        console.log(`🔍 Starting Apollo scrape for ${recordLimit} records (streaming: ${stream})...`);

        // If streaming requested, set up SSE
        if (stream) {
            res.writeHead(200, {
                'Content-Type': 'text/event-stream',
                'Cache-Control': 'no-cache',
                'Connection': 'keep-alive',
                'Access-Control-Allow-Origin': '*'
            });
        }

        function sendEvent(type, data) {
            if (stream) {
                res.write(`event: ${type}\ndata: ${JSON.stringify(data)}\n\n`);
            }
        }

        // First, scrape ALL data from Apify (this works fine)
        console.log(`⏱️ Scraping all ${recordLimit} records from Apollo...`);
        if (stream) sendEvent('progress', { message: 'Scraping from Apollo...', stage: 'scraping' });

        let apifyResponse;
        let retryCount = 0;
        const maxRetries = 2;

        while (retryCount <= maxRetries) {
            try {
                console.log(`🎯 Attempt ${retryCount + 1}/${maxRetries + 1} - Calling Apify scraper...`);
                
                apifyResponse = await axios.post(
                    'https://api.apify.com/v2/acts/code_crafter~apollo-io-scraper/run-sync-get-dataset-items',
                    {
                        cleanOutput: true,
                        totalRecords: recordLimit,
                        url: apolloUrl
                    },
                    {
                        headers: {
                            'Accept': 'application/json',
                            'Authorization': `Bearer ${process.env.APIFY_API_TOKEN}`,
                            'Connection': 'keep-alive',
                            'User-Agent': 'LGA-Lead-Generator/1.0'
                        },
                        timeout: 0,
                        maxRedirects: 5,
                        validateStatus: function (status) {
                            return status < 500;
                        }
                    }
                ).catch(error => {
                    if (error.config && error.config.headers && error.config.headers.Authorization) {
                        error.config.headers.Authorization = 'Bearer [REDACTED]';
                    }
                    throw error;
                });
                
                console.log('✅ Apify scraper completed successfully');
                break;
                
            } catch (error) {
                retryCount++;
                console.error(`❌ Attempt ${retryCount}/${maxRetries + 1} failed:`, error.code || error.message);
                
                if (retryCount > maxRetries) {
                    if (stream) {
                        sendEvent('error', { error: 'Scraping failed', message: error.message });
                        return res.end();
                    }
                    throw error;
                }
                
                await new Promise(resolve => setTimeout(resolve, 5000));
            }
        }

        const rawData = apifyResponse.data || [];
        console.log(`✅ Successfully scraped ${rawData.length} leads`);

        // Process duplicates
        const uniqueLeads = [];
        const seen = new Set();
        
        rawData.forEach(lead => {
            const email = (lead.email || '').toLowerCase().trim();
            const linkedin = (lead.linkedin_url || '').toLowerCase().trim();
            const name = (lead.name || '').toLowerCase().trim();
            const company = (lead.organization_name || '').toLowerCase().trim();
            
            let identifier;
            if (email && email !== '') {
                identifier = email;
            } else if (linkedin && linkedin !== '') {
                identifier = linkedin;
            } else {
                identifier = `${name}|${company}`;
            }
            
            if (!seen.has(identifier)) {
                seen.add(identifier);
                uniqueLeads.push(lead);
            }
        });

        // Transform all leads
        const transformedLeads = uniqueLeads.map(lead => ({
            name: lead.name || '',
            title: lead.title || '',
            organization_name: lead.organization_name || '',
            organization_website_url: lead.organization_website_url || '',
            estimated_num_employees: lead.estimated_num_employees || '',
            email: lead.email || '',
            email_verified: lead.email ? 'Y' : 'N',
            linkedin_url: lead.linkedin_url || '',
            industry: lead.industry || '',
            country: lead.country || 'Singapore',
            conversion_status: 'Pending'
        }));

        if (stream) {
            // Stream the data in chunks of 100
            const CHUNK_SIZE = 100;
            const totalChunks = Math.ceil(transformedLeads.length / CHUNK_SIZE);
            
            sendEvent('scraped_complete', { 
                totalRecords: transformedLeads.length,
                willStreamInChunks: totalChunks 
            });

            for (let i = 0; i < transformedLeads.length; i += CHUNK_SIZE) {
                const chunk = transformedLeads.slice(i, i + CHUNK_SIZE);
                const chunkNumber = Math.floor(i / CHUNK_SIZE) + 1;

                sendEvent('chunk', {
                    chunkNumber,
                    totalChunks,
                    leads: chunk,
                    processed: Math.min(i + CHUNK_SIZE, transformedLeads.length),
                    total: transformedLeads.length
                });

                // Small delay between chunks
                await new Promise(resolve => setTimeout(resolve, 100));
            }

            sendEvent('complete', {
                success: true,
                totalProcessed: transformedLeads.length
            });
            
            res.end();
        } else {
            // Regular response (for backward compatibility)
            res.json({
                success: true,
                count: transformedLeads.length,
                leads: transformedLeads,
                metadata: {
                    apolloUrl,
                    scrapedAt: new Date().toISOString(),
                    maxRecords: recordLimit,
                    finalCount: transformedLeads.length
                }
            });
        }

    } catch (error) {
        console.error('Apollo scraping error:', error);
        
        if (req.body.stream) {
            res.write(`event: error\ndata: ${JSON.stringify({
                error: 'Scraping Error',
                message: 'Failed to scrape leads from Apollo'
            })}\n\n`);
            res.end();
        } else {
            res.status(500).json({
                error: 'Scraping Error',
                message: 'Failed to scrape leads from Apollo',
                details: process.env.NODE_ENV === 'development' ? error.message : undefined
            });
        }
    }
});

// Apollo lead scraping endpoint
router.post('/scrape-leads', async (req, res) => {
    try {
        const { apolloUrl, maxRecords = 500 } = req.body;

        // Validation
        if (!apolloUrl) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Apollo URL is required'
            });
        }

        if (!apolloUrl.includes('apollo.io')) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Invalid Apollo URL'
            });
        }

        // Check if Apify API token is configured
        if (!process.env.APIFY_API_TOKEN) {
            return res.status(500).json({
                error: 'Configuration Error',
                message: 'Apify API token not configured'
            });
        }

        // Handle unlimited vs limited records
        const maxLimit = parseInt(process.env.MAX_LEADS_PER_REQUEST) || 2000;
        let recordLimit;
        
        if (maxRecords === 0) {
            recordLimit = maxLimit; // Use system max when unlimited requested
        } else {
            recordLimit = Math.min(parseInt(maxRecords) || 500, maxLimit);
        }

        console.log(`🔍 Starting Apollo scrape for ${recordLimit} records...`);

        console.log(`⏱️ No timeout limit - scraper will run until completion for ${recordLimit} records`);

        let apifyResponse;
        let retryCount = 0;
        const maxRetries = 2;

        while (retryCount <= maxRetries) {
            try {
                console.log(`🎯 Attempt ${retryCount + 1}/${maxRetries + 1} - Calling Apify scraper...`);
                
                apifyResponse = await axios.post(
                    'https://api.apify.com/v2/acts/code_crafter~apollo-io-scraper/run-sync-get-dataset-items',
                    {
                        cleanOutput: true,
                        totalRecords: recordLimit,
                        url: apolloUrl
                    },
                    {
                        headers: {
                            'Accept': 'application/json',
                            'Authorization': `Bearer ${process.env.APIFY_API_TOKEN}`,
                            'Connection': 'keep-alive',
                            'User-Agent': 'LGA-Lead-Generator/1.0'
                        },
                        timeout: 0, // No timeout - let it run until completion
                        maxRedirects: 5,
                        validateStatus: function (status) {
                            return status < 500; // Resolve only if status is less than 500
                        }
                    }
                ).catch(error => {
                    // Remove sensitive data from error logs
                    if (error.config && error.config.headers && error.config.headers.Authorization) {
                        error.config.headers.Authorization = 'Bearer [REDACTED]';
                    }
                    throw error;
                });
                
                console.log('✅ Apify scraper completed successfully');
                break; // Success, exit retry loop
                
            } catch (error) {
                retryCount++;
                console.error(`❌ Attempt ${retryCount}/${maxRetries + 1} failed:`, error.code || error.message);
                
                if (retryCount > maxRetries) {
                    // All retries exhausted - provide specific error messages
                    if (error.code === 'ECONNABORTED') {
                        throw new Error(`Apollo scraping was interrupted. Please try again with fewer records (current: ${recordLimit}) or check your network connection.`);
                    } else if (error.code === 'ECONNRESET' || error.message.includes('socket hang up') || error.message.includes('ECONNRESET')) {
                        throw new Error(`Network connection lost during scraping. This may be due to high server load. Please try again in a few minutes or reduce the record count (current: ${recordLimit}).`);
                    } else if (error.code === 'ETIMEDOUT' || error.message.includes('ETIMEDOUT')) {
                        throw new Error(`Network timeout during scraping. Please check your internet connection and try again.`);
                    } else if (error.response && error.response.status === 429) {
                        throw new Error(`Apify API rate limit exceeded. Please wait a few minutes before trying again.`);
                    } else if (error.response && error.response.status >= 500) {
                        throw new Error(`Apify server error (${error.response.status}). Please try again in a few minutes.`);
                    } else {
                        throw new Error(`Apify scraper failed after ${maxRetries + 1} attempts: ${error.message}`);
                    }
                } else {
                    // Wait before retry with longer delays for network issues
                    let waitTime;
                    if (error.code === 'ECONNRESET' || error.message.includes('socket hang up')) {
                        // Longer wait for connection issues
                        waitTime = Math.pow(2, retryCount - 1) * 10000; // 10s, 20s delays
                        console.log(`🌐 Network issue detected - waiting ${waitTime/1000}s before retry...`);
                    } else {
                        // Standard exponential backoff
                        waitTime = Math.pow(2, retryCount - 1) * 5000; // 5s, 10s delays
                        console.log(`⏳ Waiting ${waitTime/1000}s before retry...`);
                    }
                    await new Promise(resolve => setTimeout(resolve, waitTime));
                }
            }
        }

        const rawData = apifyResponse.data || [];
        
        console.log(`✅ Successfully scraped ${rawData.length} leads`);

        // Extract total count from Apify response metadata if available
        let totalAvailable = rawData.length;
        let limitReached = false;
        
        if (apifyResponse.headers && apifyResponse.headers['x-apify-total-results']) {
            totalAvailable = parseInt(apifyResponse.headers['x-apify-total-results']);
            limitReached = totalAvailable > rawData.length;
        }

        // Duplicate prevention: Remove duplicates based on email and LinkedIn URL
        const uniqueLeads = [];
        const seen = new Set();
        
        rawData.forEach(lead => {
            // Create unique identifier: email + linkedin_url (fallback to name + company)
            const email = (lead.email || '').toLowerCase().trim();
            const linkedin = (lead.linkedin_url || '').toLowerCase().trim();
            const name = (lead.name || '').toLowerCase().trim();
            const company = (lead.organization_name || '').toLowerCase().trim();
            
            let identifier;
            if (email && email !== '') {
                identifier = email; // Email is most unique
            } else if (linkedin && linkedin !== '') {
                identifier = linkedin; // LinkedIn URL second most unique
            } else {
                identifier = `${name}|${company}`; // Fallback to name+company
            }
            
            if (!seen.has(identifier)) {
                seen.add(identifier);
                uniqueLeads.push(lead);
            } else {
                console.log(`🔄 Removed duplicate: ${lead.name} (${identifier})`);
            }
        });

        const duplicatesRemoved = rawData.length - uniqueLeads.length;
        if (duplicatesRemoved > 0) {
            console.log(`🧹 Removed ${duplicatesRemoved} duplicate records`);
        }

        // Transform leads to match n8n workflow structure
        const transformedLeads = uniqueLeads.map(lead => ({
            name: lead.name || '',
            title: lead.title || '',
            organization_name: lead.organization_name || '',
            organization_website_url: lead.organization_website_url || '',
            estimated_num_employees: lead.estimated_num_employees || '',
            email: lead.email || '',
            email_verified: lead.email ? 'Y' : 'N',
            linkedin_url: lead.linkedin_url || '',
            industry: lead.industry || '',
            country: lead.country || 'Singapore',
            conversion_status: 'Pending'
        }));

        // For large datasets, don't return all leads in the response to avoid memory issues
        if (transformedLeads.length > 250) {
            // Store leads temporarily (in a real app, you'd use Redis or database)
            global.tempLeads = global.tempLeads || new Map();
            const sessionId = Date.now().toString() + Math.random().toString(36).substr(2, 9);
            global.tempLeads.set(sessionId, transformedLeads);
            
            // Clean up old sessions after 30 minutes
            setTimeout(() => {
                global.tempLeads.delete(sessionId);
            }, 30 * 60 * 1000);

            res.json({
                success: true,
                count: transformedLeads.length,
                sessionId: sessionId, // Use this to retrieve leads in chunks
                metadata: {
                    apolloUrl,
                    scrapedAt: new Date().toISOString(),
                    maxRecords: recordLimit,
                    totalAvailable: totalAvailable,
                    rawScraped: rawData.length,
                    duplicatesRemoved: duplicatesRemoved,
                    finalCount: transformedLeads.length,
                    limitReached: limitReached,
                    deduplicationStats: {
                        input: rawData.length,
                        duplicates: duplicatesRemoved,
                        unique: transformedLeads.length,
                        deduplicationRate: rawData.length > 0 ? ((duplicatesRemoved / rawData.length) * 100).toFixed(1) + '%' : '0%'
                    }
                }
            });
        } else {
            // For smaller datasets, return leads directly
            res.json({
                success: true,
                count: transformedLeads.length,
                leads: transformedLeads,
                metadata: {
                    apolloUrl,
                    scrapedAt: new Date().toISOString(),
                    maxRecords: recordLimit,
                    totalAvailable: totalAvailable,
                    rawScraped: rawData.length,
                    duplicatesRemoved: duplicatesRemoved,
                    finalCount: transformedLeads.length,
                    limitReached: limitReached,
                    deduplicationStats: {
                        input: rawData.length,
                        duplicates: duplicatesRemoved,
                        unique: transformedLeads.length,
                        deduplicationRate: rawData.length > 0 ? ((duplicatesRemoved / rawData.length) * 100).toFixed(1) + '%' : '0%'
                    }
                }
            });
        }

    } catch (error) {
        console.error('Apollo scraping error:', error);

        // Handle specific error types
        if (error.code === 'ECONNABORTED' || error.message.includes('timeout')) {
            return res.status(408).json({
                error: 'Timeout Error',
                message: 'Apollo scraping took too long. Try reducing the number of records or try again later.'
            });
        }

        if (error.response?.status === 401) {
            return res.status(401).json({
                error: 'Authentication Error', 
                message: 'Invalid Apify API token'
            });
        }

        if (error.response?.status === 429) {
            return res.status(429).json({
                error: 'Rate Limit Error',
                message: 'Apify API rate limit exceeded. Please try again later.'
            });
        }

        res.status(500).json({
            error: 'Scraping Error',
            message: 'Failed to scrape leads from Apollo',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Get leads in chunks for large datasets
router.post('/get-leads-chunk', async (req, res) => {
    try {
        const { sessionId, offset = 0, limit = 100 } = req.body;
        
        if (!sessionId) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Session ID is required'
            });
        }

        global.tempLeads = global.tempLeads || new Map();
        const allLeads = global.tempLeads.get(sessionId);
        
        if (!allLeads) {
            return res.status(404).json({
                error: 'Session Not Found',
                message: 'Session expired or invalid'
            });
        }

        const chunk = allLeads.slice(offset, offset + limit);
        
        res.json({
            success: true,
            leads: chunk,
            hasMore: offset + limit < allLeads.length,
            total: allLeads.length,
            offset: offset,
            limit: limit
        });

    } catch (error) {
        console.error('Get leads chunk error:', error);
        res.status(500).json({
            error: 'Server Error',
            message: 'Failed to retrieve leads chunk'
        });
    }
});

// Test endpoint for Apollo integration
router.get('/test', async (req, res) => {
    const checks = {
        apifyToken: !!process.env.APIFY_API_TOKEN,
        apifyConnection: false
    };

    // Test Apify connection if token is available
    if (checks.apifyToken) {
        try {
            const testResponse = await axios.get('https://api.apify.com/v2/actor-tasks', {
                headers: {
                    'Authorization': `Bearer ${process.env.APIFY_API_TOKEN}`
                },
                timeout: 5000
            });
            checks.apifyConnection = testResponse.status === 200;
        } catch (error) {
            checks.apifyConnection = false;
            checks.apifyError = error.response?.status || 'Connection failed';
        }
    }

    const allGood = Object.values(checks).every(check => 
        typeof check === 'boolean' ? check : true
    );

    res.status(allGood ? 200 : 500).json({
        status: allGood ? 'OK' : 'Error',
        checks,
        message: allGood ? 'Apollo integration ready' : 'Apollo integration has issues'
    });
});

module.exports = router;