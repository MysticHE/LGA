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

        // Determine timeout based on record count (more records = longer timeout)
        const baseTimeout = 120000; // 2 minutes base
        const additionalTimeout = Math.max(0, (recordLimit - 100) * 500); // Add 0.5s per record over 100
        const finalTimeout = Math.min(baseTimeout + additionalTimeout, 600000); // Max 10 minutes
        
        console.log(`⏱️ Using ${finalTimeout/1000}s timeout for ${recordLimit} records`);

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
                            'Authorization': `Bearer ${process.env.APIFY_API_TOKEN}`
                        },
                        timeout: finalTimeout
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
                    // All retries exhausted
                    if (error.code === 'ECONNABORTED') {
                        throw new Error(`Apollo scraping took too long (>${finalTimeout/1000}s). Try reducing the number of records (current: ${recordLimit}) or try again later.`);
                    } else {
                        throw new Error(`Apify scraper failed after ${maxRetries + 1} attempts: ${error.message}`);
                    }
                } else {
                    // Wait before retry (exponential backoff)
                    const waitTime = Math.pow(2, retryCount - 1) * 5000; // 5s, 10s delays
                    console.log(`⏳ Waiting ${waitTime/1000}s before retry...`);
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