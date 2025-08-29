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
        
        return {
            subject: parsed.subject || `Connecting with ${lead['Company Name'] || 'your company'}`,
            body: parsed.body || aiContent,
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
        const processedBody = this.replaceVariables(template.Body, lead);

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
    parseEmailContent(content) {
        if (!content) return { subject: '', body: '' };

        // Try to extract subject line
        const subjectMatch = content.match(/(?:Subject:|Subj:)\s*(.+?)(?:\n|$)/i);
        let subject = subjectMatch ? subjectMatch[1].trim() : '';

        // Extract body (everything after subject or full content if no subject)
        let body = content;
        if (subjectMatch) {
            body = content.replace(subjectMatch[0], '').trim();
        }

        // Clean up body formatting
        body = this.cleanEmailBody(body);

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
            .replace(/^(Body:|Message:)\s*/i, '') // Remove body labels
            .replace(/\n{3,}/g, '\n\n') // Limit consecutive line breaks
            .replace(/^\s+|\s+$/g, '') // Trim whitespace
            .replace(/\r\n/g, '\n'); // Normalize line endings
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
        const leadName = lead.Name || 'there';

        const fallbackSubject = `Partnership opportunity with ${companyName}`;
        const fallbackBody = `Hi ${leadName},

I hope this email finds you well. I'm reaching out because ${companyName} could benefit from our services.

I'd love to schedule a brief call to discuss how we can help ${companyName} achieve its goals.

Best regards,
[Your Name]`;

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
     * Convert email content to HTML format with tracking
     */
    convertToHTML(emailContent, leadEmail = null) {
        let htmlBody = emailContent.body || '';
        
        // Convert line breaks to HTML
        htmlBody = htmlBody.replace(/\n/g, '<br>');
        
        // Add tracking pixel if email is provided
        let trackingPixel = '';
        if (leadEmail) {
            const trackingId = `${leadEmail}-${Date.now()}`;
            const baseUrl = process.env.RENDER_EXTERNAL_URL || 'http://localhost:3000';
            trackingPixel = `<img src="${baseUrl}/api/email/track-read?id=${encodeURIComponent(trackingId)}" width="1" height="1" style="display:none;" alt="" />`;
        }
        
        // Wrap in basic HTML structure with tracking
        const html = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>${emailContent.subject || 'Email'}</title>
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    ${htmlBody}
    ${trackingPixel}
</body>
</html>`;

        return html;
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
}

module.exports = EmailContentProcessor;