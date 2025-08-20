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

// Initialize job storage (in production, use Redis or database)
global.backgroundJobs = global.backgroundJobs || new Map();

// Background job processing to avoid 5-minute timeout limit
router.post('/start-workflow-job', async (req, res) => {
    try {
        const { jobTitles, companySizes, maxRecords = 0, generateOutreach = true, chunkSize = 100 } = req.body;

        // Generate unique job ID
        const jobId = Date.now().toString() + Math.random().toString(36).substr(2, 9);
        
        // Initialize job status
        const jobStatus = {
            id: jobId,
            status: 'started',
            progress: { step: 1, message: 'Starting workflow...', total: 4 },
            startTime: new Date().toISOString(),
            params: { jobTitles, companySizes, maxRecords, generateOutreach, chunkSize },
            result: null,
            error: null,
            completedAt: null
        };
        
        // Store job status
        global.backgroundJobs.set(jobId, jobStatus);
        
        // Clean up old jobs after 2 hours
        setTimeout(() => {
            global.backgroundJobs.delete(jobId);
        }, 2 * 60 * 60 * 1000);

        console.log(`🚀 Starting background job ${jobId}...`);
        
        // Return job ID immediately to avoid timeout
        res.json({
            success: true,
            jobId: jobId,
            message: 'Workflow started in background. Use /job-status to check progress.'
        });
        
        // Run the actual workflow in background
        processWorkflowJob(jobId, req.protocol, req.get('host')).catch(error => {
            console.error(`Background job ${jobId} failed:`, error);
            const job = global.backgroundJobs.get(jobId);
            if (job) {
                job.status = 'failed';
                job.error = error.message;
                job.completedAt = new Date().toISOString();
            }
        });
        
    } catch (error) {
        console.error('Job creation error:', error);
        res.status(500).json({
            error: 'Job Creation Error',
            message: 'Failed to start background job'
        });
    }
});

// Background job processor
async function processWorkflowJob(jobId, protocol, host) {
    const job = global.backgroundJobs.get(jobId);
    if (!job) return;
    
    try {
        const { jobTitles, companySizes, maxRecords, generateOutreach, chunkSize } = job.params;
        
        // Step 1: Generate Web URL
        job.progress = { step: 2, message: 'Generating Web URL...', total: 4 };
        job.status = 'generating_url';
        
        const apolloResponse = await axios.post(`${protocol}://${host}/api/apollo/generate-url`, {
            jobTitles, companySizes
        });
        const { apolloUrl } = apolloResponse.data;
        
        // Step 2: Start Web scraping
        job.progress = { step: 3, message: 'Scraping leads from Web...', total: 4 };
        job.status = 'scraping';
        
        let scrapeData = null;
        let scrapeMetadata = {};
        const startTime = Date.now();
        
        // Update progress every 30 seconds during scraping
        const progressInterval = setInterval(() => {
            const elapsed = Math.floor((Date.now() - startTime) / 1000);
            const minutes = Math.floor(elapsed / 60);
            const seconds = elapsed % 60;
            const timeStr = minutes > 0 ? `${minutes}m ${seconds}s` : `${seconds}s`;
            
            const currentJob = global.backgroundJobs.get(jobId);
            if (currentJob && currentJob.status === 'scraping') {
                currentJob.progress = { 
                    step: 3, 
                    message: `Web scraping in progress... (${timeStr} elapsed)`, 
                    total: 4,
                    elapsed: elapsed
                };
            }
        }, 30000);
        
        try {
            // Start Web scraping asynchronously 
            console.log(`🔄 Job ${jobId}: Starting async Web scraping for ${maxRecords} records`);
            
            const apolloJobResponse = await axios.post(`${protocol}://${host}/api/apollo/start-scrape-job`, {
                apolloUrl, maxRecords
            });
            
            const apolloJobId = apolloJobResponse.data.jobId;
            console.log(`✅ Job ${jobId}: Web job started: ${apolloJobId}`);
            
            // Poll Apollo job status
            scrapeData = await pollApolloJob(apolloJobId, protocol, host, jobId, progressInterval);
            
            clearInterval(progressInterval);
            scrapeMetadata = scrapeData.metadata || {};
            
            console.log(`✅ Job ${jobId}: Successfully scraped ${scrapeData.count} leads`);
            
            if (scrapeData.count === 0) {
                job.status = 'completed';
                job.result = { 
                    success: true, 
                    count: 0, 
                    message: 'No leads found',
                    apolloUrl
                };
                job.completedAt = new Date().toISOString();
                return;
            }
            
        } catch (error) {
            clearInterval(progressInterval);
            throw new Error(`Web scraping failed: ${error.message}`);
        }

        // Step 3: Process leads in chunks
        job.progress = { step: 4, message: `Processing ${scrapeData.count} leads in batches...`, total: 4 };
        job.status = 'processing';
        
        let processedLeads = [];
        const totalChunks = Math.ceil(scrapeData.count / chunkSize);
        
        // Handle both direct leads and sessionId approaches
        if (scrapeData.leads) {
            // Small dataset - leads returned directly
            const leads = scrapeData.leads;
            for (let i = 0; i < leads.length; i += chunkSize) {
                const chunk = leads.slice(i, i + chunkSize);
                const chunkNumber = Math.floor(i / chunkSize) + 1;
                
                // Update progress
                job.progress = {
                    step: 4,
                    message: `Processing chunk ${chunkNumber}/${totalChunks}...`,
                    total: 4,
                    chunk: chunkNumber,
                    totalChunks: totalChunks
                };
                
                const finalChunk = await processChunk(chunk, generateOutreach, protocol, host);
                processedLeads.push(...finalChunk);

                if (i + chunkSize < leads.length) {
                    await new Promise(resolve => setTimeout(resolve, 500));
                }
            }
        } else if (scrapeData.sessionId) {
            // Large dataset - retrieve in chunks using sessionId
            let offset = 0;
            const chunkLimit = chunkSize;
            
            for (let chunkNumber = 1; chunkNumber <= totalChunks; chunkNumber++) {
                // Update progress
                job.progress = {
                    step: 4,
                    message: `Processing chunk ${chunkNumber}/${totalChunks}...`,
                    total: 4,
                    chunk: chunkNumber,
                    totalChunks: totalChunks
                };

                // Get chunk from Apollo
                const chunkResponse = await axios.post(`${protocol}://${host}/api/apollo/get-leads-chunk`, {
                    sessionId: scrapeData.sessionId,
                    offset: offset,
                    limit: chunkLimit
                }, { timeout: 0 });

                const chunk = chunkResponse.data.leads;
                const finalChunk = await processChunk(chunk, generateOutreach, protocol, host);
                processedLeads.push(...finalChunk);

                offset += chunkLimit;
                
                if (!chunkResponse.data.hasMore) break;
                if (chunkNumber < totalChunks) {
                    await new Promise(resolve => setTimeout(resolve, 500));
                }
            }
        }

        // Job completed successfully
        job.status = 'completed';
        job.result = {
            success: true,
            count: processedLeads.length,
            leads: processedLeads,
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
        };
        job.completedAt = new Date().toISOString();
        
        console.log(`✅ Background job ${jobId} completed successfully with ${processedLeads.length} leads`);
        
    } catch (error) {
        console.error(`❌ Background job ${jobId} failed:`, error);
        job.status = 'failed';
        job.error = error.message;
        job.completedAt = new Date().toISOString();
    }
}

// Helper function to process chunks in background job
async function processChunk(chunk, generateOutreach, protocol, host) {
    let finalChunk = chunk;

    // Generate outreach for this chunk if enabled
    if (generateOutreach && openai) {
        try {
            const outreachResponse = await axios.post(`${protocol}://${host}/api/leads/generate-outreach`, {
                leads: chunk
            }, { timeout: 0 });
            
            if (outreachResponse.data && outreachResponse.data.leads) {
                finalChunk = outreachResponse.data.leads;
            }
        } catch (error) {
            console.warn(`⚠️ Outreach generation failed for chunk:`, error.message);
        }
    }

    return finalChunk;
}

// Poll Apollo job status until completion
async function pollApolloJob(apolloJobId, protocol, host, jobId, progressInterval) {
    const job = global.backgroundJobs.get(jobId);
    let pollCount = 0;
    const maxPolls = 360; // 30 minutes max (5 second intervals)
    
    while (pollCount < maxPolls) {
        try {
            pollCount++;
            
            // Check Apollo job status
            const statusResponse = await axios.get(`${protocol}://${host}/api/apollo/job-status/${apolloJobId}`);
            const apolloStatus = statusResponse.data;
            
            // Update main job progress
            if (job) {
                const elapsed = Math.floor((Date.now() - new Date(job.startTime).getTime()) / 1000);
                const minutes = Math.floor(elapsed / 60);
                const seconds = elapsed % 60;
                const timeStr = minutes > 0 ? `${minutes}m ${seconds}s` : `${seconds}s`;
                
                job.progress = {
                    step: 3,
                    message: `Web ${apolloStatus.status}... (${timeStr} elapsed)`,
                    total: 4,
                    elapsed: elapsed,
                    apolloJobId: apolloJobId,
                    apolloStatus: apolloStatus.status
                };
            }
            
            console.log(`🔄 Job ${jobId}: Web job ${apolloJobId} status: ${apolloStatus.status} (poll ${pollCount})`);
            
            if (apolloStatus.isComplete) {
                if (apolloStatus.status === 'completed') {
                    // Get Apollo results
                    const resultResponse = await axios.get(`${protocol}://${host}/api/apollo/job-result/${apolloJobId}`);
                    console.log(`✅ Job ${jobId}: Web job completed with ${resultResponse.data.count} leads`);
                    return resultResponse.data;
                } else if (apolloStatus.status === 'failed') {
                    throw new Error(`Web job failed: ${apolloStatus.error || 'Unknown error'}`);
                }
            }
            
            // Wait 5 seconds before next poll
            await new Promise(resolve => setTimeout(resolve, 5000));
            
        } catch (error) {
            console.error(`❌ Job ${jobId}: Web polling error:`, error.message);
            
            // If it's a 404, the Apollo job might not exist
            if (error.response?.status === 404) {
                throw new Error('Web job not found - may have expired');
            }
            
            // For other errors, retry a few times
            if (pollCount < 5) {
                console.log(`⚠️ Job ${jobId}: Retrying Web status check in 10 seconds...`);
                await new Promise(resolve => setTimeout(resolve, 10000));
                continue;
            } else {
                throw new Error(`Web job polling failed: ${error.message}`);
            }
        }
    }
    
    throw new Error('Web job exceeded maximum wait time (30 minutes)');
}

// Get job status for polling
router.get('/job-status/:jobId', async (req, res) => {
    try {
        const { jobId } = req.params;
        
        if (!jobId) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Job ID is required'
            });
        }

        global.backgroundJobs = global.backgroundJobs || new Map();
        const job = global.backgroundJobs.get(jobId);
        
        if (!job) {
            return res.status(404).json({
                error: 'Job Not Found',
                message: 'Job not found or expired'
            });
        }

        // Return current job status
        res.json({
            success: true,
            jobId: jobId,
            status: job.status,
            progress: job.progress,
            startTime: job.startTime,
            completedAt: job.completedAt,
            result: job.result,
            error: job.error,
            isComplete: ['completed', 'failed'].includes(job.status)
        });

    } catch (error) {
        console.error('Job status check error:', error);
        res.status(500).json({
            error: 'Server Error',
            message: 'Failed to check job status'
        });
    }
});

// Get job result (leads data) - separate endpoint to handle large payloads
router.get('/job-result/:jobId', async (req, res) => {
    try {
        const { jobId } = req.params;
        
        if (!jobId) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Job ID is required'
            });
        }

        global.backgroundJobs = global.backgroundJobs || new Map();
        const job = global.backgroundJobs.get(jobId);
        
        if (!job) {
            return res.status(404).json({
                error: 'Job Not Found',
                message: 'Job not found or expired'
            });
        }

        if (job.status !== 'completed') {
            return res.status(400).json({
                error: 'Job Not Complete',
                message: `Job is still ${job.status}. Check job-status first.`
            });
        }

        // Return the full result with leads data
        res.json({
            success: true,
            jobId: jobId,
            ...job.result
        });

    } catch (error) {
        console.error('Job result retrieval error:', error);
        res.status(500).json({
            error: 'Server Error',
            message: 'Failed to retrieve job result'
        });
    }
});

// List all active jobs (for debugging)
router.get('/jobs', async (req, res) => {
    try {
        global.backgroundJobs = global.backgroundJobs || new Map();
        
        const jobs = Array.from(global.backgroundJobs.entries()).map(([id, job]) => ({
            jobId: id,
            status: job.status,
            startTime: job.startTime,
            completedAt: job.completedAt,
            progress: job.progress,
            hasError: !!job.error,
            params: {
                jobTitles: job.params.jobTitles,
                companySizes: job.params.companySizes,
                maxRecords: job.params.maxRecords
            }
        }));

        res.json({
            success: true,
            totalJobs: jobs.length,
            jobs: jobs
        });

    } catch (error) {
        console.error('Jobs list error:', error);
        res.status(500).json({
            error: 'Server Error',
            message: 'Failed to list jobs'
        });
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