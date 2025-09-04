const express = require('express');
const OpenAI = require('openai');
const XLSX = require('xlsx');
const axios = require('axios');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const router = express.Router();

// Import new content processing modules
const PDFContentProcessor = require('../utils/pdfContentProcessor');
const ContentAnalyzer = require('../utils/contentAnalyzer');
const ContentCache = require('../utils/contentCache');
const { getEnvironmentConfig } = require('../config/contentConfig');
const ExcelProcessor = require('../utils/excelProcessor');

// Configure multer for memory storage
const upload = multer({ 
    storage: multer.memoryStorage(),
    limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit per file
});

// Initialize OpenAI client
let openai;
if (process.env.OPENAI_API_KEY) {
    openai = new OpenAI({
        apiKey: process.env.OPENAI_API_KEY
    });
}

// Initialize content processing modules
const config = getEnvironmentConfig();
const contentProcessor = new PDFContentProcessor();
const contentAnalyzer = openai ? new ContentAnalyzer(openai) : null;
const contentCache = new ContentCache();
const excelProcessor = new ExcelProcessor();


// Initialize product materials storage (in-memory, like job storage)
global.productMaterials = global.productMaterials || new Map();

// PDF upload endpoint for product materials
router.post('/upload-materials', upload.array('pdfs'), async (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'No PDF files uploaded'
            });
        }


        const materials = [];
        const errors = [];

        for (const file of req.files) {
            try {
                const materialId = Date.now().toString() + Math.random().toString(36).substr(2, 9);
                
                // Parse PDF content
                const pdfData = await pdfParse(file.buffer);
                
                const material = {
                    id: materialId,
                    filename: file.originalname,
                    content: pdfData.text,
                    uploadedAt: new Date().toISOString(),
                    pages: pdfData.numpages,
                    size: file.size
                };
                
                global.productMaterials.set(materialId, material);
                materials.push({
                    id: materialId,
                    filename: file.originalname,
                    pages: pdfData.numpages,
                    size: file.size,
                    contentLength: pdfData.text.length
                });


            } catch (error) {
                console.error(`❌ Error processing ${file.originalname}:`, error.message);
                errors.push({
                    filename: file.originalname,
                    error: error.message
                });
            }
        }

        // Auto-cleanup materials after 24 hours
        setTimeout(() => {
            materials.forEach(material => {
                global.productMaterials.delete(material.id);
            });
        }, 24 * 60 * 60 * 1000);


        res.json({
            success: true,
            materials: materials,
            errors: errors.length > 0 ? errors : undefined,
            message: `${materials.length} PDF materials uploaded successfully`
        });

    } catch (error) {
        console.error('PDF upload error:', error);
        res.status(500).json({
            error: 'Upload Error',
            message: 'Failed to process PDF files',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// List uploaded materials
router.get('/materials', (req, res) => {
    try {
        global.productMaterials = global.productMaterials || new Map();
        
        const materials = Array.from(global.productMaterials.values()).map(material => ({
            id: material.id,
            filename: material.filename,
            uploadedAt: material.uploadedAt,
            pages: material.pages,
            size: material.size,
            contentLength: material.content.length
        }));

        res.json({
            success: true,
            count: materials.length,
            materials: materials
        });

    } catch (error) {
        console.error('Materials list error:', error);
        res.status(500).json({
            error: 'Server Error',
            message: 'Failed to list materials'
        });
    }
});

// Delete uploaded material
router.delete('/materials/:materialId', (req, res) => {
    try {
        const { materialId } = req.params;
        global.productMaterials = global.productMaterials || new Map();
        
        if (global.productMaterials.has(materialId)) {
            const material = global.productMaterials.get(materialId);
            global.productMaterials.delete(materialId);
            
            
            res.json({
                success: true,
                message: `Material ${material.filename} deleted successfully`
            });
        } else {
            res.status(404).json({
                error: 'Not Found',
                message: 'Material not found'
            });
        }

    } catch (error) {
        console.error('Material deletion error:', error);
        res.status(500).json({
            error: 'Server Error',
            message: 'Failed to delete material'
        });
    }
});

// Generate personalized outreach for leads
router.post('/generate-outreach', async (req, res) => {
    try {
        const { leads, useProductMaterials = false } = req.body;

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


        const enrichedLeads = [];
        const errors = [];

        // Process leads in batches to avoid overwhelming the API
        const batchSize = 5;
        for (let i = 0; i < leads.length; i += batchSize) {
            const batch = leads.slice(i, i + batchSize);
            
            const batchPromises = batch.map(async (lead, index) => {
                try {
                    const globalIndex = i + index;

                    // Generate personalized outreach using OpenAI
                    const outreachContent = await generateOutreachContent(lead, useProductMaterials);
                    
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

// Generate outreach content for a single lead with optimized content processing
async function generateOutreachContent(lead, useProductMaterials = false) {
    const startTime = Date.now();
    let productContext = '';
    let processingMetadata = null;
    
    if (useProductMaterials) {
        global.productMaterials = global.productMaterials || new Map();
        const materials = Array.from(global.productMaterials.values());
        
        if (materials.length > 0) {
            try {
                // Create lead context for optimization
                const leadContext = {
                    industry: lead.industry,
                    role: lead.title,
                    company: lead.organization_name,
                    country: lead.country
                };

                // Check cache first
                const cacheKey = materials.map(m => m.id).join('_');
                let optimizedResult = null;
                
                if (config.cache.enabled) {
                    optimizedResult = contentCache.getByIndustryRole(
                        cacheKey, 
                        leadContext.industry, 
                        leadContext.role
                    );
                }

                if (!optimizedResult) {
                    
                    // Process content with Content Processor (cleaning + intelligent extraction)
                    const processedResult = await contentProcessor.processContent(materials, leadContext);
                    
                    if (processedResult.success) {
                        // Use Content Processor output directly (no AI Analyzer)
                        optimizedResult = {
                            content: processedResult.content,
                            metadata: {
                                ...processedResult.metadata,
                                aiOptimized: false,
                                processingMethod: 'Content Processor Only',
                                productsPreserved: true
                            }
                        };
                        
                        // Cache the result
                        if (config.cache.enabled && optimizedResult) {
                            contentCache.cacheByIndustryRole(
                                cacheKey,
                                leadContext.industry,
                                leadContext.role,
                                optimizedResult
                            );
                        }
                    } else {
                        // Fallback to old processing method
                        console.warn('⚠️ Content processing failed, using fallback method');
                        const allContent = materials.map(m => `${m.filename}:\n${m.content}`).join('\n\n---\n\n');
                        optimizedResult = {
                            content: allContent.substring(0, 3000) + (allContent.length > 3000 ? '\n\n[Content truncated - processing error]' : ''),
                            metadata: {
                                processed: false,
                                fallbackReason: processedResult.error
                            }
                        };
                    }
                }

                productContext = optimizedResult.content;
                processingMetadata = optimizedResult.metadata;


            } catch (error) {
                console.error('❌ Content optimization error:', error);
                // Fallback to simple processing
                const allContent = materials.map(m => `${m.filename}:\n${m.content}`).join('\n\n---\n\n');
                productContext = allContent.substring(0, 3000) + (allContent.length > 3000 ? '\n\n[Content processing error - using fallback]' : '');
            }
        }
    }

    const prompt = `${productContext ? `[IF PDF MATERIALS UPLOADED - up to 3.5K characters:]
PRODUCT MATERIALS & SERVICES:
${productContext}

Use this product information to generate tailored, benefit-focused email content aligned with the prospect's industry challenges and role priorities

` : ''}PROSPECT RESEARCH

Research the company "${lead.organization_name}" using your knowledge base to understand:
Their industry-specific risks, potential insurance gaps, employee retention, wellness, medical needs

LEAD INFORMATION

Name: ${lead.name}
Title: ${lead.title}
Company: ${lead.organization_name}
Industry: ${lead.industry}
Location: ${lead.country}
LinkedIn: ${lead.linkedin_url || 'Not available'}

TASK:

Create a personalized, consultative marketing email for SME insurance solutions focusing on:
Employee protection & retention, medical needs, health and wellness, cost savings

GENERATE:

📧 PROFESSIONAL EMAIL (150–180 words)

1. Subject Line:

Generate (6–10 words) that are:
- Benefit-Oriented or pain point to capture attention
- Personalized with company or industry reference
- Clear, no jargon or vague phrasing

2. Email Body:

EMAIL OPENING RULES:
- Start with a natural sentence linking the company's business goals or industry context to the challenges addressed by insurance solutions.
- Avoid generic phrases like "you are likely aware" or "as the CEO...".
- Make it sound consultative and empathetic, not scripted.

Second paragraph with 2–3 potential pain points in bullets form companies like theirs face (e.g. Abusive claims, rising costs, talent retention, limited coverage etc).

Third paragraph name the product with 2–3 insurance solutions in bullets form from materials, described benefit-first.

Forth paragraph, briefly mention similar companies helped or results achieved.

Lastly, call to action, Professional request for brief meeting/call.

WRITING GUIDELINES:

- Professional, consultative tone — talk with them, not at them
- No generic "we offer insurance" pitches
- No heavy jargon or long product descriptions
- No passive CTAs like "looking forward to your reply"
- Output in plain text, not Markdown
- Use short paragraphs and simple bullet points for readability
- No extra formatting symbols like **bold** or ---; just clear text formatting for email

LANGUAGE & TONE RULES:
- Avoid repetitive use of "yours," "your company," or "your team."
- Maintain a consultative, professional tone — avoid overly casual or familiar language
- Keep language clear, benefit-focused, and respectful of executive-level readers

Make sure that you review the email content to achieve a rating of 10/10;
- Hooks readers faster with benefit-focused subject lines
- Speaks to their challenges clearly and empathetically
- Structures the email for easy reading with bullets and concise paragraphs
- Makes the call-to-action specific and easy to accept`;

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

// Start workflow job with exclusion file upload
router.post('/start-workflow-job-with-exclusions', upload.single('exclusionsFile'), async (req, res) => {
    try {
        const { 
            jobTitles, 
            companySizes, 
            maxRecords = 0, 
            generateOutreach = true, 
            useProductMaterials = false, 
            chunkSize = 100,
            excludeIndustries = [],
            saveToOneDrive = false,
            sendEmailCampaign = false,
            templateChoice = 'AI_Generated',
            emailTemplate = '',
            emailSubject = '',
            trackEmailReads = true
        } = req.body;

        // Parse JSON strings from FormData - all arrays come as strings when using multipart
        const parsedJobTitles = typeof jobTitles === 'string' ? 
            JSON.parse(jobTitles || '[]') : jobTitles;
        const parsedCompanySizes = typeof companySizes === 'string' ? 
            JSON.parse(companySizes || '[]') : companySizes;
        const parsedExcludeIndustries = typeof excludeIndustries === 'string' ? 
            JSON.parse(excludeIndustries || '[]') : excludeIndustries;

        // Extract exclusion domains from uploaded file
        let excludeEmailDomains = [];
        if (req.file) {
            // Check for Excel temporary files
            if (req.file.originalname.startsWith('~$')) {
                return res.status(400).json({
                    error: 'Invalid File',
                    message: 'Cannot process Excel temporary file. Please close Excel and upload the actual file (not the ~$ temporary file).'
                });
            }
            
            console.log(`🚫 Processing exclusions file: ${req.file.originalname} (${req.file.size} bytes)`);
            try {
                excludeEmailDomains = excelProcessor.parseExclusionDomainsFromExcel(req.file.buffer);
                console.log(`✅ Extracted ${excludeEmailDomains.length} exclusion domains from uploaded file`);
            } catch (error) {
                console.error('❌ Failed to extract exclusion domains:', error.message);
                return res.status(400).json({
                    error: 'File Processing Error',
                    message: 'Failed to extract exclusion domains from uploaded file: ' + error.message
                });
            }
        }

        // Generate unique job ID
        const jobId = Date.now().toString() + Math.random().toString(36).substr(2, 9);
        
        // Initialize job status
        const jobStatus = {
            id: jobId,
            status: 'started',
            progress: { step: 1, message: 'Starting workflow with domain exclusions...', total: saveToOneDrive || sendEmailCampaign ? 6 : 4 },
            startTime: new Date().toISOString(),
            params: { 
                jobTitles: parsedJobTitles, 
                companySizes: parsedCompanySizes, 
                maxRecords, 
                generateOutreach, 
                useProductMaterials, 
                chunkSize, 
                excludeEmailDomains, 
                excludeIndustries: parsedExcludeIndustries, 
                saveToOneDrive, 
                sendEmailCampaign, 
                templateChoice, 
                emailTemplate, 
                emailSubject, 
                trackEmailReads,
                exclusionFileUploaded: !!req.file,
                exclusionFileName: req.file?.originalname
            },
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

        // Return job ID immediately to avoid timeout
        res.json({
            success: true,
            jobId: jobId,
            message: 'Workflow started with domain exclusions. Use /job-status to check progress.',
            exclusionStats: {
                domainsExtracted: excludeEmailDomains.length,
                domains: excludeEmailDomains.slice(0, 10),
                fileName: req.file?.originalname
            }
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
            message: 'Failed to start background job with exclusions'
        });
    }
});

// Background job processing to avoid 5-minute timeout limit
router.post('/start-workflow-job', async (req, res) => {
    try {
        const { 
            jobTitles, 
            companySizes, 
            maxRecords = 0, 
            generateOutreach = true, 
            useProductMaterials = false, 
            chunkSize = 100,
            excludeEmailDomains = [],
            excludeIndustries = [],
            saveToOneDrive = false,
            sendEmailCampaign = false,
            templateChoice = 'AI_Generated',
            emailTemplate = '',
            emailSubject = '',
            trackEmailReads = true
        } = req.body;

        // Generate unique job ID
        const jobId = Date.now().toString() + Math.random().toString(36).substr(2, 9);
        
        // Initialize job status
        const jobStatus = {
            id: jobId,
            status: 'started',
            progress: { step: 1, message: 'Starting workflow...', total: saveToOneDrive || sendEmailCampaign ? 6 : 4 },
            startTime: new Date().toISOString(),
            params: { jobTitles, companySizes, maxRecords, generateOutreach, useProductMaterials, chunkSize, excludeEmailDomains, excludeIndustries, saveToOneDrive, sendEmailCampaign, templateChoice, emailTemplate, emailSubject, trackEmailReads },
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
        const { jobTitles, companySizes, maxRecords, generateOutreach, useProductMaterials, chunkSize, excludeEmailDomains, excludeIndustries, saveToOneDrive, sendEmailCampaign, templateChoice, emailTemplate, emailSubject, trackEmailReads } = job.params;
        
        // Step 1: Generate Web URL
        const totalSteps = 4 + (saveToOneDrive ? 1 : 0) + (sendEmailCampaign ? 1 : 0);
        job.progress = { step: 2, message: 'Generating Web URL...', total: totalSteps };
        job.status = 'generating_url';
        
        const apolloResponse = await axios.post(`${protocol}://${host}/api/apollo/generate-url`, {
            jobTitles, companySizes
        });
        const { apolloUrl } = apolloResponse.data;
        
        // Step 2: Start Web scraping
        job.progress = { step: 3, message: 'Scraping leads from Web...', total: totalSteps };
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
                    total: totalSteps,
                    elapsed: elapsed
                };
            }
        }, 30000);
        
        try {
            // Start Web scraping asynchronously
            
            const apolloJobResponse = await axios.post(`${protocol}://${host}/api/apollo/start-scrape-job`, {
                apolloUrl, maxRecords
            });
            
            const apolloJobId = apolloJobResponse.data.jobId;
            
            // Poll Apollo job status
            scrapeData = await pollApolloJob(apolloJobId, protocol, host, jobId, progressInterval, totalSteps);
            
            clearInterval(progressInterval);
            scrapeMetadata = scrapeData.metadata || {};
            
            
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
        job.progress = { step: 4, message: `Processing ${scrapeData.count} leads in batches...`, total: totalSteps };
        job.status = 'processing';
        
        let processedLeads = [];
        let totalFilteredCount = 0;
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
                    total: totalSteps,
                    chunk: chunkNumber,
                    totalChunks: totalChunks
                };
                
                const chunkResult = await processChunk(chunk, generateOutreach, useProductMaterials, excludeEmailDomains, excludeIndustries, protocol, host);
                processedLeads.push(...chunkResult.leads);
                totalFilteredCount += chunkResult.filteredCount;

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
                    total: totalSteps,
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
                const chunkResult = await processChunk(chunk, generateOutreach, useProductMaterials, excludeEmailDomains, excludeIndustries, protocol, host);
                processedLeads.push(...chunkResult.leads);
                totalFilteredCount += chunkResult.filteredCount;

                offset += chunkLimit;
                
                if (!chunkResponse.data.hasMore) break;
                if (chunkNumber < totalChunks) {
                    await new Promise(resolve => setTimeout(resolve, 500));
                }
            }
        }

        // Step 5: Save to OneDrive (if enabled)
        let oneDriveFileId = null;
        if (saveToOneDrive && processedLeads.length > 0) {
            try {
                job.progress = { step: 5, message: 'Saving leads to OneDrive...', total: sendEmailCampaign ? 6 : 5 };
                job.status = 'saving_onedrive';
                
                
                const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
                const filename = `singapore-leads-${timestamp}.xlsx`;
                
                const oneDriveResponse = await axios.post(`${protocol}://${host}/api/microsoft-graph/onedrive/append-to-table`, {
                    leads: processedLeads,
                    filename: filename,
                    folderPath: '/LGA-Leads',
                    useCustomFile: true  // Allow custom filename for compatibility
                });
                
                if (oneDriveResponse.data.success) {
                    oneDriveFileId = oneDriveResponse.data.fileId;
                }
            } catch (oneDriveError) {
                console.error(`⚠️ Job ${jobId}: OneDrive save failed:`, oneDriveError.message);
                // Continue with workflow even if OneDrive fails
            }
        }

        // Step 6: Send Email Campaign (if enabled)
        let campaignResult = null;
        if (sendEmailCampaign && emailSubject && processedLeads.length > 0 && (templateChoice !== 'custom' || emailTemplate)) {
            try {
                const finalStep = saveToOneDrive ? 6 : 5;
                job.progress = { step: finalStep, message: 'Sending email campaign...', total: finalStep };
                job.status = 'sending_emails';
                
                
                const emailResponse = await axios.post(`${protocol}://${host}/api/email-automation/send-campaign`, {
                    leads: processedLeads,
                    templateChoice: templateChoice,
                    emailTemplate: emailTemplate,
                    subject: emailSubject,
                    trackReads: trackEmailReads,
                    oneDriveFileId: oneDriveFileId
                });
                
                if (emailResponse.data.success) {
                    campaignResult = {
                        campaignId: emailResponse.data.campaignId,
                        sent: emailResponse.data.sent,
                        failed: emailResponse.data.failed,
                        trackingEnabled: emailResponse.data.trackingEnabled
                    };
                }
            } catch (emailError) {
                console.error(`⚠️ Job ${jobId}: Email campaign failed:`, emailError.message);
                // Continue with workflow even if email campaign fails
            }
        }

        // Job completed successfully
        job.status = 'completed';
        // Update scrapeMetadata to reflect post-filtering results
        const originalTotalAvailable = scrapeMetadata.totalAvailable || scrapeData.count;
        const originalFinalCount = scrapeMetadata.finalCount || scrapeData.count;
        
        // Calculate corrected numbers after exclusion filters
        const correctedFinalCount = processedLeads.length;
        const correctedTotalAvailable = originalTotalAvailable; // Keep original as "available before filters"
        
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
                totalFilteredOut: totalFilteredCount,
                outreachGenerated: generateOutreach && !!openai,
                usedProductMaterials: useProductMaterials && generateOutreach,
                excludedEmailDomains: excludeEmailDomains,
                excludedIndustries: excludeIndustries,
                filtersApplied: excludeEmailDomains.length > 0 || excludeIndustries.length > 0,
                scrapeMetadata: {
                    ...scrapeMetadata,
                    // Override key metrics to reflect post-filtering reality
                    totalAvailable: correctedTotalAvailable,
                    finalCount: correctedFinalCount,
                    // Add original pre-filter values for reference
                    originalTotalAvailable: originalTotalAvailable,
                    originalFinalCount: originalFinalCount,
                    postFilterAdjustment: true
                },
                oneDriveIntegration: {
                    enabled: saveToOneDrive,
                    success: oneDriveFileId ? true : false,
                    fileId: oneDriveFileId
                },
                emailCampaign: campaignResult,
                completedAt: new Date().toISOString()
            }
        };
        job.completedAt = new Date().toISOString();
        
        
    } catch (error) {
        console.error(`Background job ${jobId} failed:`, error);
        job.status = 'failed';
        job.error = error.message;
        job.completedAt = new Date().toISOString();
    }
}

// Helper function to filter leads based on exclusion criteria
function filterLeads(leads, excludeEmailDomains = [], excludeIndustries = []) {
    if (!leads || leads.length === 0) return { filteredLeads: leads, filteredCount: 0 };
    
    const filteredLeads = leads.filter(lead => {
        // Filter by email domain
        if (excludeEmailDomains.length > 0 && lead.email) {
            const emailDomain = lead.email.split('@')[1]?.toLowerCase();
            if (emailDomain && excludeEmailDomains.some(domain => 
                emailDomain === domain.toLowerCase() || emailDomain.endsWith('.' + domain.toLowerCase())
            )) {
                return false;
            }
        }
        
        // Filter by industry
        if (excludeIndustries.length > 0 && lead.industry) {
            const leadIndustry = lead.industry.toLowerCase();
            if (excludeIndustries.some(industry => 
                leadIndustry.includes(industry.toLowerCase()) || industry.toLowerCase().includes(leadIndustry)
            )) {
                return false;
            }
        }
        
        return true;
    });
    
    const filteredCount = leads.length - filteredLeads.length;
    if (filteredCount > 0) {
    }
    
    return { filteredLeads, filteredCount };
}

// Helper function to process chunks in background job
async function processChunk(chunk, generateOutreach, useProductMaterials, excludeEmailDomains, excludeIndustries, protocol, host) {
    // Apply exclusion filters first
    const { filteredLeads, filteredCount } = filterLeads(chunk, excludeEmailDomains, excludeIndustries);
    let finalChunk = filteredLeads;

    // Generate outreach for this chunk if enabled (only if we have leads after filtering)
    if (generateOutreach && openai && finalChunk.length > 0) {
        try {
            const outreachResponse = await axios.post(`${protocol}://${host}/api/leads/generate-outreach`, {
                leads: finalChunk,
                useProductMaterials: useProductMaterials
            }, { timeout: 0 });
            
            if (outreachResponse.data && outreachResponse.data.leads) {
                finalChunk = outreachResponse.data.leads;
            }
        } catch (error) {
            console.warn(`Outreach generation failed for chunk:`, error.message);
        }
    }

    return { leads: finalChunk, filteredCount };
}

// Poll Apollo job status until completion
async function pollApolloJob(apolloJobId, protocol, host, jobId, progressInterval, totalSteps) {
    const job = global.backgroundJobs.get(jobId);
    let pollCount = 0;
    
    while (true) {
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
                    total: totalSteps,
                    elapsed: elapsed,
                    apolloJobId: apolloJobId,
                    apolloStatus: apolloStatus.status
                };
            }
            
            
            if (apolloStatus.isComplete) {
                if (apolloStatus.status === 'completed') {
                    // Get Apollo results
                    const resultResponse = await axios.get(`${protocol}://${host}/api/apollo/job-result/${apolloJobId}`);
                    return resultResponse.data;
                } else if (apolloStatus.status === 'failed') {
                    throw new Error(`Web job failed: ${apolloStatus.error || 'Unknown error'}`);
                }
            }
            
            // Wait 5 seconds before next poll
            await new Promise(resolve => setTimeout(resolve, 5000));
            
        } catch (error) {
            console.error(`Job ${jobId}: Web polling error:`, error.message);
            
            // If it's a 404, the Apollo job might not exist
            if (error.response?.status === 404) {
                throw new Error('Web job not found - may have expired');
            }
            
            // For other errors, retry a few times
            if (pollCount < 5) {
                await new Promise(resolve => setTimeout(resolve, 10000));
                continue;
            } else {
                throw new Error(`Web job polling failed: ${error.message}`);
            }
        }
    }
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

// Content processing info and cache statistics endpoint
router.get('/content-processing-info', (req, res) => {
    try {
        const cacheInfo = contentCache.getCacheInfo();
        const configInfo = {
            processingEnabled: !!contentProcessor,
            aiAnalysisEnabled: !!contentAnalyzer,
            cacheEnabled: config.cache.enabled,
            features: config.features,
            aiConfig: {
                model: config.ai.model,
                tokenLimits: config.ai.tokenLimits,
                rateLimiting: config.ai.rateLimiting
            }
        };

        res.json({
            success: true,
            system: 'Content Processing Engine v1.0',
            timestamp: new Date().toISOString(),
            config: configInfo,
            cache: cacheInfo,
            materials: {
                count: global.productMaterials?.size || 0,
                totalSize: Array.from(global.productMaterials?.values() || [])
                    .reduce((sum, m) => sum + m.content.length, 0)
            }
        });

    } catch (error) {
        console.error('Content processing info error:', error);
        res.status(500).json({
            error: 'Info Error',
            message: 'Failed to get content processing information',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Clear content cache endpoint (useful for development)
router.post('/clear-content-cache', (req, res) => {
    try {
        const beforeStats = contentCache.getStats();
        contentCache.clear();
        
        res.json({
            success: true,
            message: 'Content cache cleared successfully',
            clearedEntries: beforeStats.currentEntries,
            stats: contentCache.getStats()
        });

    } catch (error) {
        console.error('Clear cache error:', error);
        res.status(500).json({
            error: 'Cache Error',
            message: 'Failed to clear content cache',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Debug endpoint to check Excel file structure
router.post('/debug-excel-structure', upload.single('excelFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({
                error: 'No file provided'
            });
        }

        console.log(`📊 Debugging Excel structure: ${req.file.originalname}`);
        
        const XLSX = require('xlsx');
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        
        const structure = {};
        
        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            if (data.length > 0) {
                structure[sheetName] = {
                    rowCount: data.length,
                    headers: Object.keys(data[0] || {}),
                    sampleRow: data[0]
                };
            }
        }
        
        res.json({
            success: true,
            fileName: req.file.originalname,
            fileSize: req.file.size,
            sheets: workbook.SheetNames,
            structure: structure
        });
        
    } catch (error) {
        console.error('Debug Excel error:', error);
        res.status(500).json({
            error: 'Debug Error',
            message: error.message
        });
    }
});

module.exports = router;