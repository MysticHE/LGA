const express = require('express');
const OpenAI = require('openai');
const XLSX = require('xlsx');
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

// Complete lead generation workflow with progress updates
router.post('/complete-workflow', async (req, res) => {
    try {
        const { jobTitles, companySizes, maxRecords = 0, generateOutreach = true } = req.body;

        console.log('🚀 Starting complete lead generation workflow...');

        // Step 1: Generate Apollo URL
        console.log('📋 Step 1: Generating Apollo URL...');
        const apolloResponse = await fetch(`${req.protocol}://${req.get('host')}/api/apollo/generate-url`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ jobTitles, companySizes })
        });

        if (!apolloResponse.ok) {
            throw new Error('Failed to generate Apollo URL');
        }

        const { apolloUrl } = await apolloResponse.json();
        console.log('✅ Apollo URL generated successfully');

        // Step 2: Scrape leads from Apollo
        console.log('🔍 Step 2: Scraping leads from Apollo...');
        const scrapeResponse = await fetch(`${req.protocol}://${req.get('host')}/api/apollo/scrape-leads`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ apolloUrl, maxRecords })
        });

        if (!scrapeResponse.ok) {
            const scrapeError = await scrapeResponse.json();
            throw new Error(scrapeError.message || 'Failed to scrape leads');
        }

        const scrapeData = await scrapeResponse.json();
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
            
            try {
                const outreachResponse = await fetch(`${req.protocol}://${req.get('host')}/api/leads/generate-outreach`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ leads })
                });

                if (outreachResponse.ok) {
                    const outreachData = await outreachResponse.json();
                    finalLeads = outreachData.leads;
                    console.log('✅ AI outreach content generated successfully');
                } else {
                    console.warn('⚠️ Outreach generation failed, continuing without it');
                }
            } catch (outreachError) {
                console.warn('⚠️ Outreach generation error:', outreachError.message);
            }
        }

        console.log(`🎉 Workflow completed successfully with ${finalLeads.length} leads`);

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
                outreachGenerated: generateOutreach && openai,
                scrapeMetadata,
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