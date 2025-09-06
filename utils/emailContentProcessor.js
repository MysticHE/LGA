/**
 * Email Content Processing Utilities
 * Handles email template processing, variable replacement, and content generation
 */

class EmailContentProcessor {
    constructor() {
        this.variablePattern = /\{([^}]+)\}/g;
        this.supportedVariables = [
            'Name', 'Company', 'Title', 'Industry', 'Location', 'Email',
            'Company_Name', 'LinkedIn_URL', 'Size'
        ];
    }

    /**
     * Process email content based on lead data and email choice
     */
    async processEmailContent(lead, emailChoice, templates = []) {
        try {
            console.log(`📧 Processing email content for ${lead.Email} using ${emailChoice}`);

            switch (emailChoice) {
                case 'AI_Generated':
                    return this.getAIGeneratedContent(lead);
                
                case 'Email_Template_1':
                case 'Email_Template_2':
                    return this.processTemplate(emailChoice, lead, templates);
                
                default:
                    // Default to AI-generated content
                    return this.getAIGeneratedContent(lead);
            }
        } catch (error) {
            console.error('❌ Email content processing error:', error);
            throw new Error('Failed to process email content: ' + error.message);
        }
    }

    /**
     * Get AI-generated email content from lead data
     */
    getAIGeneratedContent(lead) {
        const aiContent = lead.AI_Generated_Email || lead.Notes || '';
        
        if (!aiContent) {
            console.warn(`⚠️ No AI-generated content found for ${lead.Email}`);
            return this.generateFallbackContent(lead);
        }

        // Parse AI content to extract subject and body
        const parsed = this.parseEmailContent(aiContent);
        
        // Fix Issue 2 & 3: Personalize square bracket placeholders and ensure greeting
        let personalizedBody = this.personalizeSquareBrackets(parsed.body, lead);
        personalizedBody = this.ensureProperGreeting(personalizedBody, lead);
        
        // Final subject with proper fallback
        let finalSubject = parsed.subject;
        if (!finalSubject) {
            finalSubject = `Connecting with ${lead['Company Name'] || 'your company'}`;
        }

        console.log(`📧 Email processed for ${lead.Email}`);
        
        return {
            subject: finalSubject,
            body: personalizedBody,
            contentType: 'AI_Generated',
            variables: this.extractVariables(aiContent)
        };
    }

    /**
     * Process template with variable replacement
     */
    processTemplate(templateChoice, lead, templates) {
        console.log(`📝 Processing template: ${templateChoice}`);

        // Find the template
        const template = templates.find(t => 
            t.Template_ID === templateChoice || 
            t.Template_Name === templateChoice ||
            t.Template_Type === templateChoice
        );

        if (!template) {
            console.warn(`⚠️ Template ${templateChoice} not found, using fallback`);
            return this.generateFallbackContent(lead);
        }

        // Replace variables in template
        const processedSubject = this.replaceVariables(template.Subject, lead);
        let processedBody = this.replaceVariables(template.Body, lead);
        
        // Remove placeholder signatures from template body
        processedBody = this.removePlaceholderSignatures(processedBody);

        return {
            subject: processedSubject,
            body: processedBody,
            contentType: templateChoice,
            templateId: template.Template_ID,
            variables: this.extractVariables(template.Subject + ' ' + template.Body)
        };
    }

    /**
     * Replace template variables with lead data
     */
    replaceVariables(content, lead) {
        if (!content) return '';

        return content.replace(this.variablePattern, (match, variableName) => {
            const trimmedVar = variableName.trim();
            
            // Map common variable names to lead properties
            const variableMap = {
                'Name': lead.Name || '',
                'Company': lead['Company Name'] || '',
                'Company_Name': lead['Company Name'] || '',
                'Title': lead.Title || '',
                'Industry': lead.Industry || '',
                'Location': lead.Location || '',
                'Email': lead.Email || '',
                'LinkedIn_URL': lead['LinkedIn URL'] || '',
                'Size': lead.Size || '',
                'Website': lead['Company Website'] || ''
            };

            const replacement = variableMap[trimmedVar] || 
                               lead[trimmedVar] || 
                               match; // Keep original if no replacement found

            console.log(`🔄 Replacing {${trimmedVar}} with "${replacement}"`);
            return replacement;
        });
    }

    /**
     * Parse email content to extract subject and body
     */
    parseEmailContent(aiContent) {
        if (!aiContent || typeof aiContent !== 'string') return { subject: '', body: '' };

        // Remove BOM & trim
        aiContent = aiContent.replace(/^\uFEFF/, '').trim();

        console.log(`📧 Parsing AI content...`);

        // Improved regex: captures "Subject Line:" even with extra spaces or after numbers
        const subjectMatch = aiContent.match(/(?:Subject\s*Line:|^\s*\d+\.\s*Subject\s*Line:)\s*(.+)/im);

        const subject = subjectMatch ? subjectMatch[1].trim() : '';

        // Remove subject line from body if found
        let body = aiContent;
        if (subjectMatch) {
            body = aiContent.replace(subjectMatch[0], '').trim();
        }

        // Additional cleanup: remove "Email Body:" or numbered body labels
        body = body.replace(/^(?:\d+\.\s*)?Email\s*Body:\s*/im, '').trim();
        
        // Remove placeholder signatures from parsed content
        body = this.removePlaceholderSignatures(body);

        console.log(`📧 Content parsed successfully`);

        return {
            subject: subject,
            body: body
        };
    }

    /**
     * Clean and format email body content
     */
    cleanEmailBody(body) {
        if (!body) return '';

        return body
            // Remove AI-generated labels and formatting
            .replace(/^(Subject Line:|Email Body:|Body:|Message:)\s*/mi, '') // Remove content labels
            .replace(/\d+\.\s*(Subject Line:|Email Body:)/gi, '') // Remove numbered labels
            .replace(/^(Subject:|Subj:)\s*.+$/mi, '') // Remove any remaining subject lines
            .replace(/^Email Body:\s*/mi, '') // Remove "Email Body:" prefix if it exists
            .replace(/^Dear\s+[^,]+,\s*/mi, '$&') // Preserve greeting but clean extra spaces
            .replace(/^Hi\s+[^,]+,\s*/mi, '$&') // Preserve greeting but clean extra spaces
            // Remove placeholder signature components
            .replace(/\[Your Name\]/gi, '') // Remove placeholder name
            .replace(/\[Your Position\]/gi, '') // Remove placeholder position
            .replace(/\[Your Contact Information\]/gi, '') // Remove placeholder contact
            .replace(/Inspro Insurance Brokers(?:\s*\n)?/gi, '') // Remove standalone company name
            .replace(/^\s*\n+/g, '') // Remove leading empty lines
            .replace(/\n{3,}/g, '\n\n') // Limit consecutive line breaks
            .replace(/^\s+|\s+$/g, '') // Trim whitespace
            .replace(/\r\n/g, '\n') // Normalize line endings
            .trim(); // Final trim
    }

    /**
     * Extract variables used in content
     */
    extractVariables(content) {
        if (!content) return [];

        const variables = [];
        let match;
        
        const regex = new RegExp(this.variablePattern);
        while ((match = regex.exec(content)) !== null) {
            const variable = match[1].trim();
            if (!variables.includes(variable)) {
                variables.push(variable);
            }
        }

        return variables;
    }

    /**
     * Generate fallback content when no template or AI content is available
     */
    generateFallbackContent(lead) {
        const companyName = lead['Company Name'] || 'your company';
        const leadName = lead.Name || 'Sir/Madam';

        const fallbackSubject = `Partnership opportunity with ${companyName}`;
        const fallbackBody = `Dear ${leadName},

I hope this email finds you well. I'm reaching out because ${companyName} could benefit from our services.

I'd love to schedule a brief call to discuss how we can help ${companyName} achieve its goals.

Best regards,
Joel Lee`;

        return {
            subject: fallbackSubject,
            body: fallbackBody,
            contentType: 'Fallback',
            variables: ['Name', 'Company']
        };
    }

    /**
     * Validate email content before sending
     */
    validateEmailContent(emailContent) {
        const errors = [];

        if (!emailContent.subject || emailContent.subject.trim() === '') {
            errors.push('Subject line is required');
        }

        if (!emailContent.body || emailContent.body.trim() === '') {
            errors.push('Email body is required');
        }

        // Check for unresolved variables
        const unresolvedVars = this.findUnresolvedVariables(emailContent.subject + ' ' + emailContent.body);
        if (unresolvedVars.length > 0) {
            errors.push(`Unresolved variables: ${unresolvedVars.join(', ')}`);
        }

        // Check subject line length
        if (emailContent.subject && emailContent.subject.length > 100) {
            errors.push('Subject line is too long (max 100 characters)');
        }

        // Check body length
        if (emailContent.body && emailContent.body.length > 10000) {
            errors.push('Email body is too long (max 10,000 characters)');
        }

        return {
            isValid: errors.length === 0,
            errors: errors
        };
    }

    /**
     * Find unresolved template variables
     */
    findUnresolvedVariables(content) {
        const unresolved = [];
        let match;
        
        const regex = new RegExp(this.variablePattern);
        while ((match = regex.exec(content)) !== null) {
            unresolved.push(match[1].trim());
        }

        return unresolved;
    }

    /**
     * Preview email content with sample data
     */
    previewEmailContent(template, sampleLead = null) {
        const defaultSampleLead = {
            Name: 'John Smith',
            'Company Name': 'ABC Corporation',
            Title: 'Marketing Manager',
            Industry: 'Technology',
            Location: 'Singapore',
            Email: 'john.smith@abccorp.com',
            'LinkedIn URL': 'https://linkedin.com/in/johnsmith',
            Size: '100-500'
        };

        const lead = sampleLead || defaultSampleLead;
        
        return {
            subject: this.replaceVariables(template.Subject, lead),
            body: this.replaceVariables(template.Body, lead),
            variables: this.extractVariables(template.Subject + ' ' + template.Body)
        };
    }

    /**
     * Convert email content to HTML format with tracking and professional signature
     */
    convertToHTML(emailContent, leadEmail = null, leadData = null) {
        let htmlBody = emailContent.body || '';
        
        // Clean placeholder signatures from body
        htmlBody = this.removePlaceholderSignatures(htmlBody);
        
        // Convert line breaks to HTML
        htmlBody = htmlBody.replace(/\n/g, '<br>');
        
        // Add professional CTA button for direct reply
        const ctaButton = this.generateCTAButton(leadData);
        
        // Add professional Inspro signature
        const professionalSignature = this.generateProfessionalSignature();
        
        // Add tracking pixel if email is provided
        let trackingPixel = '';
        if (leadEmail) {
            const trackingId = `${leadEmail}-${Date.now()}`;
            const baseUrl = process.env.RENDER_EXTERNAL_URL || 'http://localhost:3000';
            trackingPixel = `<img src="${baseUrl}/api/email/track-read?id=${encodeURIComponent(trackingId)}" width="1" height="1" style="display:none;" alt="" />`;
        }
        
        // Wrap in enhanced HTML structure with CTA button and professional signature
        const html = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${emailContent.subject || 'Email'}</title>
    <style>
        .email-container { max-width: 600px; margin: 0 auto; font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .email-body { margin-bottom: 30px; }
        .cta-section { text-align: center; margin: 30px 0; padding: 20px; background-color: #f8f9fa; border-radius: 8px; }
        .cta-button { background-color: #28a745; color: white !important; padding: 15px 30px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; font-size: 16px; margin: 10px 0; }
        .cta-button:hover { background-color: #218838; }
        .cta-text { font-size: 14px; color: #666; margin-top: 10px; }
        .signature { border-top: 1px solid #e0e0e0; padding-top: 20px; margin-top: 30px; }
        .logo { margin-bottom: 10px; }
        .contact-info { font-size: 13px; color: #666; }
        .legal-text { font-size: 11px; color: #999; margin-top: 15px; border-top: 1px solid #f0f0f0; padding-top: 10px; }
    </style>
</head>
<body>
    <div class="email-container">
        <div class="email-body">
            ${htmlBody}
        </div>
        ${ctaButton}
        ${professionalSignature}
        ${trackingPixel}
    </div>
</body>
</html>`;

        return html;
    }

    /**
     * Generate professional CTA button for direct email reply
     */
    generateCTAButton(leadData = null) {
        const companyName = leadData?.['Company Name'] || leadData?.Company || 'your company';
        const leadName = leadData?.Name || '';
        
        // Create mailto link with pre-filled reply
        const subject = encodeURIComponent('Re: Insurance Inquiry - Keen to Know More');
        const body = encodeURIComponent(`Hi Joel,

I'm interested in learning more about Inspro's insurance solutions for ${companyName}. Please send me detailed information about your services.

Best regards,
${leadName}`);
        
        const mailtoLink = `mailto:BenefitsCare@inspro.com.sg?subject=${subject}&body=${body}`;
        
        return `
        <div class="cta-section">
            <div style="margin-bottom: 15px; font-size: 16px; color: #333;">
                <strong>Interested in our insurance solutions?</strong>
            </div>
            <a href="${mailtoLink}" class="cta-button">
                📞 Yes, I'm Keen to Know More
            </a>
            <div class="cta-text">
                Click the button above to send us a quick reply, or simply respond to this email.
            </div>
        </div>`;
    }

    /**
     * Remove placeholder signature components from email body
     */
    removePlaceholderSignatures(body) {
        if (!body) return '';
        
        return body
            // Remove common placeholder patterns
            .replace(/\[Your Name\]\s*/gi, '')
            .replace(/\[Your Title\]\s*/gi, '')
            .replace(/\[Your Position\]\s*/gi, '')  
            .replace(/\[Your Contact Information\]\s*/gi, '')
            .replace(/\[Your Company\]\s*/gi, '') // Fix Issue 1: Remove [Your Company]
            .replace(/Inspro Insurance Brokers(?:\s*\n)?\s*/gi, '')
            // Remove "Best regards," if followed by placeholders
            .replace(/Best regards,\s*\n\s*\[Your Name\]/gi, '')
            .replace(/Best regards,\s*\n\s*\[Your Title\]/gi, '')
            // Clean up extra whitespace
            .replace(/\n\s*\n\s*\n/g, '\n\n')
            .trim();
    }

    /**
     * Generate professional Inspro Insurance Brokers signature
     */
    generateProfessionalSignature() {
        return `
        <div class="signature">
            <div class="logo">
                <img src="https://ik.imagekit.io/ofkmpd3cb/inspro%20logo.jpg?updatedAt=1756520750006" 
                     alt="Inspro Insurance Brokers" 
                     style="height: 40px; max-width: 200px;" />
            </div>
            
            <div class="contact-info">
                <strong>Joel Lee – Client Relations Manager</strong><br>
                <strong>Inspro Insurance Brokers Pte Ltd (199307139Z)</strong><br><br>
                
                38 Jalan Pemimpin M38 #02-08 Singapore 577178<br>
                E: <a href="mailto:joellee@inspro.com.sg" style="color: #0066cc;">joellee@inspro.com.sg</a> 
                W: <a href="https://www.inspro.com.sg" style="color: #0066cc;">www.inspro.com.sg</a>
            </div>
            
            <div class="legal-text">
                <p><strong>Privacy Statement:</strong> In accordance with data protection law we do not use or disclose personal information for any purpose that is unrelated to our services. In providing your data you have agreed to the use of this related to our services. A copy of our Privacy statement is available on request.</p>
                
                <p><strong>Confidentiality Notice:</strong> This e-mail is intended for the named addressee only. It contains information which may be privileged and confidential. Unless you are the named addressee you may neither use it, copy it nor disclose it to anyone else. If you have received it in error please notify the sender immediately by email or telephone. Thank You.</p>
            </div>
        </div>`;
    }

    /**
     * Create email message object for Microsoft Graph API
     */
    createEmailMessage(emailContent, leadEmail, leadData = null, trackReads = false) {
        return {
            subject: emailContent.subject,
            body: {
                contentType: 'HTML',
                content: this.convertToHTML(emailContent, trackReads ? leadEmail : null, leadData)
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: leadEmail,
                        name: leadData?.Name || leadEmail
                    }
                }
            ]
        };
    }

    /**
     * Add tracking pixel to email content
     */
    addTrackingPixel(emailContent, trackingUrl) {
        if (!trackingUrl) return emailContent;

        const trackingPixel = `<img src="${trackingUrl}" width="1" height="1" style="display:none;" />`;
        
        // If content is HTML, add before closing body tag
        if (emailContent.includes('<body') && emailContent.includes('</body>')) {
            return emailContent.replace('</body>', trackingPixel + '</body>');
        } else {
            // Add to end of content
            return emailContent + trackingPixel;
        }
    }

    /**
     * Personalize email content based on lead information
     */
    personalizeContent(content, lead, personalizationLevel = 'standard') {
        let personalizedContent = content;

        switch (personalizationLevel) {
            case 'basic':
                // Only replace name and company
                personalizedContent = this.replaceVariables(content, {
                    Name: lead.Name,
                    Company: lead['Company Name']
                });
                break;

            case 'standard':
                // Replace all available variables
                personalizedContent = this.replaceVariables(content, lead);
                break;

            case 'advanced':
                // Advanced personalization with industry-specific content
                personalizedContent = this.replaceVariables(content, lead);
                personalizedContent = this.addIndustrySpecificContent(personalizedContent, lead.Industry);
                break;
        }

        return personalizedContent;
    }

    /**
     * Add industry-specific content
     */
    addIndustrySpecificContent(content, industry) {
        if (!industry) return content;

        const industryInsights = {
            'Technology': 'digital transformation and scalability',
            'Manufacturing': 'operational efficiency and supply chain optimization',
            'Healthcare': 'patient care and regulatory compliance',
            'Finance': 'financial security and regulatory requirements',
            'Retail': 'customer experience and inventory management',
            'Education': 'student engagement and administrative efficiency'
        };

        const insight = industryInsights[industry];
        if (insight) {
            // Add industry insight if placeholder exists
            content = content.replace(
                '{Industry_Insight}', 
                `particularly in ${insight}`
            );
        }

        return content;
    }

    /**
     * Calculate email content statistics
     */
    getContentStats(emailContent) {
        const subject = emailContent.subject || '';
        const body = emailContent.body || '';

        return {
            subjectLength: subject.length,
            bodyLength: body.length,
            wordCount: body.split(/\s+/).filter(word => word.length > 0).length,
            variableCount: this.extractVariables(subject + ' ' + body).length,
            estimatedReadTime: Math.ceil(body.split(/\s+/).length / 200) // Average 200 words per minute
        };
    }

    /**
     * Fix Issue 2: Personalize square bracket placeholders in AI-generated content
     */
    personalizeSquareBrackets(body, lead) {
        if (!body || !lead) return body;

        const leadName = lead.Name || null;
        const companyName = lead['Company Name'] || lead.Company || 'your company';
        
        // For greetings, use "Sir/Madam" when name is not available
        const greetingName = leadName || 'Sir/Madam';
        // For other contexts, use "there" as fallback
        const fallbackName = leadName || 'there';
        
        return body
            .replace(/\[Owner's Name\]/gi, fallbackName)
            .replace(/\[Owner Name\]/gi, fallbackName)
            .replace(/\[Name\]/gi, fallbackName)
            .replace(/\[Recipient's Name\]/gi, greetingName)
            .replace(/\[Recipient Name\]/gi, greetingName)
            .replace(/\[Your Company\]/gi, companyName)
            .replace(/\[Company Name\]/gi, companyName)
            .replace(/\[Company\]/gi, companyName);
    }

    /**
     * Fix Issue 3: Ensure proper greeting at start of email
     */
    ensureProperGreeting(body, lead) {
        if (!body || !lead) return body;

        const leadName = lead.Name || 'Sir/Madam';
        const bodyTrimmed = body.trim();
        
        // Check if body already starts with a proper greeting
        const hasGreeting = /^(Dear|Hi|Hello|Hey)\s+/i.test(bodyTrimmed);
        
        if (!hasGreeting) {
            // Add proper greeting at the start
            return `Dear ${leadName},\n\n${bodyTrimmed}`;
        }
        
        return body;
    }

}

module.exports = EmailContentProcessor;