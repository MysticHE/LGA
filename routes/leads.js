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
    let heartbeatInterval;
    
    try {
        const { jobTitles, companySizes, maxRecords = 0, generateOutreach = true, chunkSize = 100 } = req.body;

        // Set SSE headers for streaming
        res.writeHead(200, {
            'Content-Type': 'text/event-stream',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Allow-Methods': 'POST',
            'X-Accel-Buffering': 'no' // Disable nginx buffering
        });

        function sendEvent(type, data) {
            const payload = JSON.stringify(data);
            res.write(`event: ${type}\ndata: ${payload}\n\n`);
        }
        
        // Send heartbeat to keep connection alive
        function sendHeartbeat() {
            res.write(`: heartbeat ${Date.now()}\n\n`);
        }
        
        // Start heartbeat every 30 seconds
        heartbeatInterval = setInterval(sendHeartbeat, 30000);

        // Handle client disconnect
        req.on('close', () => {
            console.log('Client disconnected from stream');
            clearInterval(heartbeatInterval);
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
        
        // Step 2: Start Apollo scraping in background and stream progress
        sendEvent('progress', { step: 3, message: 'Starting Apollo scraper...', total: 4 });
        
        let scrapeData = null;
        let scrapeMetadata = {};
        
        try {
            // Start scraping with periodic progress updates
            const startTime = Date.now();
            let progressInterval;
            
            // Send progress updates every 15 seconds while scraping
            const sendProgressUpdate = () => {
                const elapsed = Math.floor((Date.now() - startTime) / 1000);
                const minutes = Math.floor(elapsed / 60);
                const seconds = elapsed % 60;
                const timeStr = minutes > 0 ? `${minutes}m ${seconds}s` : `${seconds}s`;
                
                sendEvent('progress', { 
                    step: 3, 
                    message: `Apollo still scraping... (${timeStr} elapsed)`, 
                    total: 4,
                    elapsed: elapsed
                });
            };
            
            // Start progress interval
            progressInterval = setInterval(sendProgressUpdate, 15000);
            
            // Make the Apollo request
            const scrapeResponse = await axios.post(`${req.protocol}://${req.get('host')}/api/apollo/scrape-leads`, {
                apolloUrl, maxRecords
            }, { timeout: 0 });
            
            // Clear progress interval
            clearInterval(progressInterval);
            
            scrapeData = scrapeResponse.data;
            scrapeMetadata = scrapeData.metadata || {};
            
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
            
        } catch (error) {
            console.error('Apollo scraping failed:', error);
            sendEvent('error', { 
                error: 'Apollo Scraping Failed', 
                message: error.message || 'Failed to scrape leads from Apollo',
                details: error.response?.data || error.code
            });
            clearInterval(heartbeatInterval);
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
                totalFound: scrapeData.count,
                processed: processedLeads.length,
                outreachGenerated: generateOutreach && !!openai,
                scrapeMetadata,
                completedAt: new Date().toISOString()
            }
        });

        // Clean up intervals
        clearInterval(heartbeatInterval);
        res.end();

    } catch (error) {
        console.error('Streaming workflow error:', error);
        res.write(`event: error\ndata: ${JSON.stringify({
            error: 'Workflow Error',
            message: error.message
        })}\n\n`);
        clearInterval(heartbeatInterval);
        res.end();
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