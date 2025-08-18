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

        // Call Apify Apollo scraper (same as n8n workflow)
        const apifyResponse = await axios.post(
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
                timeout: 120000 // 2 minute timeout
            }
        );

        const leads = apifyResponse.data || [];
        
        console.log(`✅ Successfully scraped ${leads.length} leads`);

        // Transform leads to match n8n workflow structure
        const transformedLeads = leads.map(lead => ({
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
                maxRecords: recordLimit
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