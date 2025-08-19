const express = require('express');
const OpenAI = require('openai');
const XLSX = require('xlsx');
const axios = require('axios');
const router = express.Router();

// Initialize OpenAI client
let openai;
if (process.env.OPENAI_API_KEY) {
    openai = new OpenAI({
        apiKey: process.env.OPENAI_API_KEY
    });
}

// Generate personalized outreach for leads
router.post('/generate-outreach', async (req, res) => {
    try {
        const { leads } = req.body;

        // Validation
        if (!leads || !Array.isArray(leads) || leads.length === 0) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Leads array is required and must not be empty'
            });
        }

        if (!openai) {
            return res.status(500).json({
                error: 'Configuration Error',
                message: 'OpenAI API key not configured'
            });
        }

        console.log(`🤖 Generating outreach for ${leads.length} leads...`);

        const enrichedLeads = [];
        const errors = [];

        // Process leads in batches to avoid overwhelming the API
        const batchSize = 5;
        for (let i = 0; i < leads.length; i += batchSize) {
            const batch = leads.slice(i, i + batchSize);
            
            const batchPromises = batch.map(async (lead, index) => {
                try {
                    const globalIndex = i + index;
                    console.log(`📝 Processing lead ${globalIndex + 1}/${leads.length}: ${lead.name}`);

                    // Generate personalized outreach using OpenAI
                    const outreachContent = await generateOutreachContent(lead);
                    
                    return {
                        ...lead,
                        notes: outreachContent,
                        outreach_generated: true,
                        processed_at: new Date().toISOString()
                    };
                } catch (error) {
                    console.error(`❌ Error processing lead ${lead.name}:`, error.message);
                    errors.push({
                        lead: lead.name,
                        error: error.message
                    });
                    
                    return {
                        ...lead,
                        notes: 'Error generating outreach content',
                        outreach_generated: false,
                        error: error.message
                    };
                }
            });

            const batchResults = await Promise.all(batchPromises);
            enrichedLeads.push(...batchResults);

            // Add delay between batches to respect rate limits
            if (i + batchSize < leads.length) {
                await new Promise(resolve => setTimeout(resolve, 1000));
            }
        }

        console.log(`✅ Completed outreach generation. Success: ${enrichedLeads.length - errors.length}, Errors: ${errors.length}`);

        res.json({
            success: true,
            count: enrichedLeads.length,
            leads: enrichedLeads,
            errors: errors.length > 0 ? errors : undefined,
            metadata: {
                processedAt: new Date().toISOString(),
                successRate: ((enrichedLeads.length - errors.length) / enrichedLeads.length * 100).toFixed(1) + '%'
            }
        });

    } catch (error) {
        console.error('Outreach generation error:', error);
        res.status(500).json({
            error: 'Outreach Generation Error',
            message: 'Failed to generate outreach content',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Generate outreach content for a single lead
async function generateOutreachContent(lead) {
    const prompt = `LinkedIn Personalized Outreach Generator
Analyze the LinkedIn profile below and create personalized outreach content.
LinkedIn URL: ${lead.linkedin_url || 'Not available'}

Reference specific details like recent posts, job changes, company news, or shared connections.

Lead Information:
- Name: ${lead.name}
- Title: ${lead.title}
- Company: ${lead.organization_name}
- Industry: ${lead.industry}
- Location: ${lead.country}

Your Product/Service: SME Insurance
Goal: Partnership and sales opportunity

Generate:
🔥 ICE BREAKER (50-80 words):
Casual, genuine opener
Reference 2-3 specific profile details
Natural conversation starter

📧 EMAIL (100-150 words):
Subject: [Personalized 5-8 words]
Opening: Personal hook with specific details
Body: Value proposition for their role/industry
CTA: Clear next step

Rules:
✅ Use specific details from the lead information
✅ Match their seniority level tone
✅ Sound human, not automated
❌ No generic templates
❌ Don't over-compliment
❌ Avoid assumptions about needs`;

    try {
        const response = await openai.chat.completions.create({
            model: 'gpt-4o-mini',
            messages: [
                {
                    role: 'system',
                    content: prompt
                }
            ],
            max_tokens: 500,
            temperature: 0.7
        });

        return response.choices[0]?.message?.content || 'Failed to generate outreach content';
    } catch (error) {
        throw new Error(`OpenAI API error: ${error.message}`);
    }
}

// Export leads to Excel
router.post('/export-excel', async (req, res) => {
    try {
        const { leads, filename } = req.body;

        // Validation
        if (!leads || !Array.isArray(leads) || leads.length === 0) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Leads array is required and must not be empty'
            });
        }

        console.log(`📊 Exporting ${leads.length} leads to Excel...`);

        // Transform leads to Excel format (matching n8n workflow structure)
        const excelData = leads.map(lead => ({
            'Name': lead.name || '',
            'Title': lead.title || '',
            'Company Name': lead.organization_name || '',
            'Company Website': lead.organization_website_url || '',
            'Size': lead.estimated_num_employees || '',
            'Email': lead.email || '',
            'Email Verified': lead.email_verified || 'N',
            'LinkedIn URL': lead.linkedin_url || '',
            'Industry': lead.industry || '',
            'Location': lead.country || '',
            'Notes': lead.notes || '',
            'Conversion Status': lead.conversion_status || 'Pending'
        }));

        // Create workbook
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(excelData);

        // Set column widths for better readability
        ws['!cols'] = [
            {width: 20}, // Name
            {width: 25}, // Title
            {width: 30}, // Company Name
            {width: 35}, // Company Website
            {width: 15}, // Size
            {width: 30}, // Email
            {width: 15}, // Email Verified
            {width: 40}, // LinkedIn URL
            {width: 20}, // Industry
            {width: 15}, // Location
            {width: 50}, // Notes
            {width: 18}  // Conversion Status
        ];

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, 'Leads');

        // Generate Excel file buffer
        const excelBuffer = XLSX.write(wb, { 
            type: 'buffer', 
            bookType: 'xlsx' 
        });

        // Generate filename with timestamp if not provided
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const finalFilename = filename || `singapore-leads-${timestamp}.xlsx`;

        console.log(`✅ Excel file generated: ${finalFilename}`);

        // Send file as download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${finalFilename}"`);
        res.setHeader('Content-Length', excelBuffer.length);

        res.send(excelBuffer);

    } catch (error) {
        console.error('Excel export error:', error);
        res.status(500).json({
            error: 'Export Error',
            message: 'Failed to export leads to Excel',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Streaming lead generation workflow with chunked responses  
router.post('/complete-workflow-stream', async (req, res) => {
    // Add request logging for debugging
    console.log('🔍 Streaming workflow requested:', {
        timestamp: new Date().toISOString(),
        ip: req.ip,
        userAgent: req.get('User-Agent')?.substring(0, 50) + '...'
    });
    try {
        const { jobTitles, companySizes, maxRecords = 0, generateOutreach = true, chunkSize = 100 } = req.body;

        // Set SSE headers for streaming
        res.writeHead(200, {
            'Content-Type': 'text/event-stream',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Allow-Methods': 'POST'
        });

        function sendEvent(type, data) {
            const payload = JSON.stringify(data);
            res.write(`event: ${type}\ndata: ${payload}\n\n`);
        }

        // Handle client disconnect
        req.on('close', () => {
            console.log('Client disconnected from stream');
        });

        console.log('🚀 Starting streaming lead generation workflow...');
        sendEvent('progress', { step: 1, message: 'Starting workflow...', total: 4 });
        
        // Step 1: Generate Apollo URL
        sendEvent('progress', { step: 2, message: 'Generating Apollo URL...', total: 4 });
        const apolloResponse = await axios.post(`${req.protocol}://${req.get('host')}/api/apollo/generate-url`, {
            jobTitles, companySizes
        });
        const { apolloUrl } = apolloResponse.data;
        
        sendEvent('progress', { step: 3, message: 'Scraping leads from Apollo...', total: 4 });
        
        // Step 2: Scrape all leads first, then stream process them
        const scrapeResponse = await axios.post(`${req.protocol}://${req.get('host')}/api/apollo/scrape-leads`, {
            apolloUrl, maxRecords
        }, { timeout: 0 });
        
        const scrapeData = scrapeResponse.data;
        const { metadata: scrapeMetadata } = scrapeData;
        console.log(`✅ Successfully scraped ${scrapeData.count} leads`);
        
        sendEvent('scraped', { 
            count: scrapeData.count, 
            metadata: scrapeMetadata,
            apolloUrl 
        });

        if (scrapeData.count === 0) {
            sendEvent('complete', { 
                success: true, 
                count: 0, 
                message: 'No leads found' 
            });
            res.end();
            return;
        }

        // Step 3: Process leads in chunks
        sendEvent('progress', { step: 4, message: `Processing ${scrapeData.count} leads in batches...`, total: 4 });
        
        let processedLeads = [];
        const totalChunks = Math.ceil(scrapeData.count / chunkSize);
        
        // Handle both direct leads (small datasets) and sessionId (large datasets)
        if (scrapeData.leads) {
            // Small dataset - leads returned directly
            const leads = scrapeData.leads;
            for (let i = 0; i < leads.length; i += chunkSize) {
                const chunk = leads.slice(i, i + chunkSize);
                const chunkNumber = Math.floor(i / chunkSize) + 1;
                
                const finalChunk = await processChunk(chunk, chunkNumber, totalChunks, generateOutreach);
                processedLeads.push(...finalChunk);
                
                sendEvent('chunk_complete', {
                    chunk: chunkNumber,
                    total: totalChunks,
                    leads: finalChunk,
                    processed: processedLeads.length,
                    remaining: leads.length - processedLeads.length
                });

                if (i + chunkSize < leads.length) {
                    await new Promise(resolve => setTimeout(resolve, 500));
                }
            }
        } else if (scrapeData.sessionId) {
            // Large dataset - retrieve in chunks using sessionId
            let offset = 0;
            const chunkLimit = chunkSize;
            
            for (let chunkNumber = 1; chunkNumber <= totalChunks; chunkNumber++) {
                sendEvent('chunk_start', { 
                    chunk: chunkNumber, 
                    total: totalChunks, 
                    size: chunkLimit 
                });

                try {
                    // Get chunk from Apollo
                    const chunkResponse = await axios.post(`${req.protocol}://${req.get('host')}/api/apollo/get-leads-chunk`, {
                        sessionId: scrapeData.sessionId,
                        offset: offset,
                        limit: chunkLimit
                    }, { timeout: 0 });

                    const chunk = chunkResponse.data.leads;
                    const finalChunk = await processChunk(chunk, chunkNumber, totalChunks, generateOutreach);
                    processedLeads.push(...finalChunk);

                    sendEvent('chunk_complete', {
                        chunk: chunkNumber,
                        total: totalChunks,
                        leads: finalChunk,
                        processed: processedLeads.length,
                        remaining: scrapeData.count - processedLeads.length
                    });

                    offset += chunkLimit;
                    
                    // Only add delay if there are more chunks
                    if (!chunkResponse.data.hasMore) break;
                    if (chunkNumber < totalChunks) {
                        await new Promise(resolve => setTimeout(resolve, 500));
                    }
                } catch (error) {
                    console.error(`Error retrieving chunk ${chunkNumber}:`, error);
                    sendEvent('error', { error: 'Chunk retrieval failed', message: error.message });
                    res.end();
                    return;
                }
            }
        }

        // Helper function to process each chunk
        async function processChunk(chunk, chunkNumber, totalChunks, generateOutreach) {
            sendEvent('chunk_start', { 
                chunk: chunkNumber, 
                total: totalChunks, 
                size: chunk.length 
            });

            let finalChunk = chunk;

            // Generate outreach for this chunk if enabled
            if (generateOutreach && openai) {
                try {
                    const outreachResponse = await axios.post(`${req.protocol}://${req.get('host')}/api/leads/generate-outreach`, {
                        leads: chunk
                    }, { timeout: 0 });
                    
                    if (outreachResponse.data && outreachResponse.data.leads) {
                        finalChunk = outreachResponse.data.leads;
                    }
                } catch (error) {
                    console.warn(`⚠️ Outreach generation failed for chunk ${chunkNumber}:`, error.message);
                }
            }

            return finalChunk;
        }

        // Send final completion event
        sendEvent('complete', {
            success: true,
            count: processedLeads.length,
            metadata: {
                apolloUrl,
                jobTitles,
                companySizes,
                maxRecords,
                totalFound: leads.length,
                processed: processedLeads.length,
                outreachGenerated: generateOutreach && !!openai,
                scrapeMetadata,
                completedAt: new Date().toISOString()
            }
        });

        res.end();

    } catch (error) {
        console.error('Streaming workflow error:', error);
        res.write(`event: error\ndata: ${JSON.stringify({
            error: 'Workflow Error',
            message: error.message
        })}\n\n`);
        res.end();
    }
});

// Complete lead generation workflow with progress updates
router.post('/complete-workflow', async (req, res) => {
    try {
        const { jobTitles, companySizes, maxRecords = 0, generateOutreach = true } = req.body;

        console.log('🚀 Starting complete lead generation workflow...');
        console.log('📋 Request body:', { jobTitles, companySizes, maxRecords, generateOutreach });
        
        // Check environment variables first
        console.log('🔍 Environment check:');
        console.log('- APIFY_API_TOKEN:', process.env.APIFY_API_TOKEN ? '✅ Set' : '❌ Missing');
        console.log('- OPENAI_API_KEY:', process.env.OPENAI_API_KEY ? '✅ Set' : '❌ Missing');
        
        // Validate request data
        if (!jobTitles || !Array.isArray(jobTitles) || jobTitles.length === 0) {
            console.error('❌ Validation Error: Job titles missing or invalid');
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Job titles are required and must be a non-empty array'
            });
        }
        
        if (!companySizes || !Array.isArray(companySizes) || companySizes.length === 0) {
            console.error('❌ Validation Error: Company sizes missing or invalid');
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Company sizes are required and must be a non-empty array'
            });
        }

        // Step 1: Generate Apollo URL
        console.log('📋 Step 1: Generating Apollo URL...');
        let apolloData;
        try {
            const apolloResponse = await axios.post(`${req.protocol}://${req.get('host')}/api/apollo/generate-url`, {
                jobTitles,
                companySizes
            }, {
                timeout: 30000
            });
            apolloData = apolloResponse.data;
        } catch (axiosError) {
            console.error('❌ Apollo URL generation error:', axiosError.response?.data || axiosError.message);
            throw new Error(`Failed to generate Apollo URL: ${axiosError.response?.data?.message || axiosError.message}`);
        }

        const { apolloUrl } = apolloData;
        console.log('✅ Apollo URL generated successfully:', apolloUrl);

        // Step 2: Scrape leads from Apollo
        console.log('🔍 Step 2: Scraping leads from Apollo...');
        console.log(`⏱️ No timeout limit - will wait for scraper to complete (${maxRecords} records)`);
        
        let scrapeData;
        let retryCount = 0;
        const maxRetries = 2;

        while (retryCount <= maxRetries) {
            try {
                console.log(`🎯 Internal API call attempt ${retryCount + 1}/${maxRetries + 1} - Calling Apollo scraper...`);
                
                const scrapeResponse = await axios.post(`${req.protocol}://${req.get('host')}/api/apollo/scrape-leads`, {
                    apolloUrl,
                    maxRecords
                }, {
                    timeout: 0, // No timeout - wait for completion
                    headers: {
                        'Connection': 'keep-alive',
                        'User-Agent': 'LGA-Lead-Generator/1.0'
                    }
                });
                
                scrapeData = scrapeResponse.data;
                console.log('✅ Internal API call completed successfully');
                break; // Success, exit retry loop
                
            } catch (axiosError) {
                retryCount++;
                console.error(`❌ Internal API attempt ${retryCount}/${maxRetries + 1} failed:`, axiosError.code || axiosError.message);
                
                if (retryCount > maxRetries) {
                    // All retries exhausted
                    if (axiosError.code === 'ECONNRESET' || axiosError.message.includes('socket hang up')) {
                        throw new Error(`Network connection lost during lead scraping. The Apollo scraper may have completed successfully but the connection was interrupted. Please check the results or try again.`);
                    } else if (axiosError.response && axiosError.response.data) {
                        throw new Error(`Failed to scrape leads: ${axiosError.response.data.message || axiosError.response.data.error || 'Unknown server error'}`);
                    } else {
                        throw new Error(`Failed to scrape leads after ${maxRetries + 1} attempts: ${axiosError.message}`);
                    }
                } else {
                    // Wait before retry with longer delays for network issues
                    let waitTime;
                    if (axiosError.code === 'ECONNRESET' || axiosError.message.includes('socket hang up')) {
                        waitTime = Math.pow(2, retryCount - 1) * 8000; // 8s, 16s delays
                        console.log(`🌐 Network issue in internal API - waiting ${waitTime/1000}s before retry...`);
                    } else {
                        waitTime = Math.pow(2, retryCount - 1) * 3000; // 3s, 6s delays  
                        console.log(`⏳ Waiting ${waitTime/1000}s before internal API retry...`);
                    }
                    await new Promise(resolve => setTimeout(resolve, waitTime));
                }
            }
        }
        const { leads, metadata: scrapeMetadata } = scrapeData;

        console.log(`✅ Successfully scraped ${leads.length} leads from Apollo`);

        if (leads.length === 0) {
            return res.json({
                success: true,
                message: 'No leads found for the specified criteria',
                leads: [],
                apolloUrl,
                metadata: {
                    apolloUrl,
                    jobTitles,
                    companySizes,
                    maxRecords,
                    totalFound: 0,
                    scraped: 0,
                    outreachGenerated: false,
                    completedAt: new Date().toISOString()
                }
            });
        }

        let finalLeads = leads;

        // Step 3: Generate outreach content (optional)
        if (generateOutreach && openai && leads.length > 0) {
            console.log(`🤖 Step 3: Generating AI outreach for ${leads.length} leads...`);
            console.log(`⏱️ No timeout limit - will generate outreach until completion (${leads.length} leads)`);
            
            try {
                const outreachResponse = await axios.post(`${req.protocol}://${req.get('host')}/api/leads/generate-outreach`, {
                    leads
                }, {
                    timeout: 0 // No timeout - wait for completion
                });

                if (outreachResponse.data) {
                    finalLeads = outreachResponse.data.leads;
                    console.log('✅ AI outreach content generated successfully');
                }
            } catch (outreachError) {
                console.warn('⚠️ Outreach generation error:', outreachError.response?.data || outreachError.message);
                // Continue without outreach generation on error
            }
        }

        console.log(`🎉 Workflow completed successfully with ${finalLeads.length} leads`);

        // Ensure scrapeMetadata is serializable by creating a clean copy
        const cleanScrapeMetadata = scrapeMetadata ? {
            apolloUrl: scrapeMetadata.apolloUrl,
            scrapedAt: scrapeMetadata.scrapedAt,
            maxRecords: scrapeMetadata.maxRecords,
            totalAvailable: scrapeMetadata.totalAvailable,
            rawScraped: scrapeMetadata.rawScraped,
            duplicatesRemoved: scrapeMetadata.duplicatesRemoved,
            finalCount: scrapeMetadata.finalCount,
            limitReached: scrapeMetadata.limitReached,
            deduplicationStats: scrapeMetadata.deduplicationStats
        } : null;

        res.json({
            success: true,
            count: finalLeads.length,
            leads: finalLeads,
            metadata: {
                apolloUrl,
                jobTitles,
                companySizes,
                maxRecords: maxRecords || 'unlimited',
                totalFound: leads.length, // This is what Apify actually found
                scraped: finalLeads.length,
                outreachGenerated: generateOutreach && !!openai,
                scrapeMetadata: cleanScrapeMetadata,
                completedAt: new Date().toISOString()
            }
        });

    } catch (error) {
        console.error('Complete workflow error:', error);
        res.status(500).json({
            error: 'Workflow Error',
            message: 'Failed to complete lead generation workflow',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Test streaming endpoint
router.post('/test-stream', async (req, res) => {
    try {
        // Set SSE headers
        res.writeHead(200, {
            'Content-Type': 'text/event-stream',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Access-Control-Allow-Origin': '*'
        });

        function sendEvent(type, data) {
            res.write(`event: ${type}\ndata: ${JSON.stringify(data)}\n\n`);
        }

        // Send test events
        sendEvent('progress', { step: 1, total: 3, message: 'Starting test...' });
        
        await new Promise(resolve => setTimeout(resolve, 1000));
        sendEvent('progress', { step: 2, total: 3, message: 'Processing...' });
        
        await new Promise(resolve => setTimeout(resolve, 1000));
        sendEvent('chunk_complete', { 
            chunk: 1, 
            total: 1, 
            leads: [{ name: 'Test Lead', email: 'test@example.com' }],
            processed: 1,
            remaining: 0
        });
        
        await new Promise(resolve => setTimeout(resolve, 500));
        sendEvent('complete', { success: true, count: 1 });
        
        res.end();
    } catch (error) {
        console.error('Test stream error:', error);
        res.status(500).json({ error: 'Test failed' });
    }
});

// Test endpoint for OpenAI integration
router.get('/test', async (req, res) => {
    const checks = {
        openaiKey: !!process.env.OPENAI_API_KEY,
        openaiConnection: false
    };

    // Test OpenAI connection if key is available
    if (checks.openaiKey && openai) {
        try {
            const testResponse = await openai.chat.completions.create({
                model: 'gpt-4o-mini',
                messages: [{ role: 'user', content: 'Test' }],
                max_tokens: 5
            });
            checks.openaiConnection = !!testResponse.choices[0];
        } catch (error) {
            checks.openaiConnection = false;
            checks.openaiError = error.message;
        }
    }

    const allGood = checks.openaiKey && checks.openaiConnection;

    res.status(allGood ? 200 : 500).json({
        status: allGood ? 'OK' : 'Error',
        checks,
        message: allGood ? 'OpenAI integration ready' : 'OpenAI integration has issues'
    });
});

module.exports = router;