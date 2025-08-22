/**
 * AI-Powered Content Analyzer
 * 
 * Uses OpenAI to intelligently summarize, analyze, and optimize PDF content
 * for maximum relevance in insurance email generation.
 */

class ContentAnalyzer {
    constructor(openaiClient) {
        this.openai = openaiClient;
        
        this.analysisPrompts = {
            summarize: `You are an expert insurance content analyst. Your task is to create concise, sales-focused summaries of insurance product materials.

INSTRUCTIONS:
- Extract key selling points and business benefits
- Preserve specific product names and coverage types
- Focus on value propositions that matter to businesses
- Remove legal disclaimers and boilerplate text
- Keep technical insurance terms only if they add business value
- Maintain professional, consultative tone

OUTPUT FORMAT:
Create a structured summary with:
1. Key Products/Services (bullet points)
2. Primary Benefits (business-focused)
3. Target Industries/Use Cases
4. Unique Value Propositions

Keep the summary under 500 words while preserving all critical selling information.`,

            industryOptimize: `You are an insurance sales expert specializing in industry-specific solutions. Optimize the following content for a prospect in the {INDUSTRY} industry.

TASK:
- Highlight insurance products most relevant to {INDUSTRY} businesses
- Emphasize risks and challenges specific to this industry
- Include industry-specific benefits and use cases
- Use terminology familiar to {INDUSTRY} professionals
- Focus on compliance and regulatory aspects relevant to this sector

TARGET ROLE: {ROLE}
- Tailor language appropriateness for this seniority level
- Focus on strategic vs. operational concerns based on role
- Highlight decision-making factors relevant to this position

Keep output focused and under 400 words.`,

            emailOptimize: `You are an expert email copywriter for insurance sales. Transform this product content into compelling email material that drives prospect engagement.

REQUIREMENTS:
- Create content that naturally fits into a professional business email
- Focus on problems this insurance solves for the prospect
- Include specific product benefits that create urgency
- Make it conversational but professional
- Remove all sales jargon and technical complexity
- Include quantifiable benefits where possible

OUTPUT:
Create 3-4 key talking points that can be woven into an email:
1. Problem/Risk identification
2. Solution overview
3. Business benefits
4. Credibility/proof points

Each point should be 1-2 sentences maximum. Total under 300 words.`
        };
    }

    /**
     * Analyze and optimize content using AI
     * @param {string} content - Content to analyze
     * @param {Object} options - Analysis options
     * @returns {Object} Analyzed content with metadata
     */
    async analyzeContent(content, options = {}) {
        try {
            const {
                type = 'summarize',
                industry = null,
                role = null,
                maxTokens = 400,
                temperature = 0.3
            } = options;

            // Choose appropriate analysis type
            let analysisResult;
            switch (type) {
                case 'summarize':
                    analysisResult = await this.summarizeContent(content, maxTokens, temperature);
                    break;
                case 'industry':
                    analysisResult = await this.optimizeForIndustry(content, industry, role, maxTokens, temperature);
                    break;
                case 'email':
                    analysisResult = await this.optimizeForEmail(content, maxTokens, temperature);
                    break;
                default:
                    throw new Error(`Unknown analysis type: ${type}`);
            }

            return {
                success: true,
                type,
                originalLength: content.length,
                analyzedContent: analysisResult.content,
                analyzedLength: analysisResult.content.length,
                compressionRatio: analysisResult.content.length / content.length,
                tokensUsed: analysisResult.tokensUsed,
                confidence: this.calculateConfidence(analysisResult.content),
                metadata: {
                    analysisType: type,
                    industry,
                    role,
                    timestamp: new Date().toISOString()
                }
            };

        } catch (error) {
            console.error('Content analysis error:', error);
            return {
                success: false,
                error: error.message,
                fallbackContent: this.createFallbackSummary(content)
            };
        }
    }

    /**
     * Summarize content for general use
     * @param {string} content - Content to summarize
     * @param {number} maxTokens - Maximum tokens for response
     * @param {number} temperature - AI temperature setting
     * @returns {Object} Summarization result
     */
    async summarizeContent(content, maxTokens, temperature) {
        const prompt = this.analysisPrompts.summarize;
        
        try {
            const response = await this.openai.chat.completions.create({
                model: 'gpt-4o-mini',
                messages: [
                    { role: 'system', content: prompt },
                    { role: 'user', content: `Please summarize this insurance content:\n\n${content}` }
                ],
                max_tokens: maxTokens,
                temperature: temperature
            });

            return {
                content: response.choices[0]?.message?.content || '',
                tokensUsed: response.usage?.total_tokens || 0
            };

        } catch (error) {
            throw new Error(`Summarization failed: ${error.message}`);
        }
    }

    /**
     * Optimize content for specific industry and role
     * @param {string} content - Content to optimize
     * @param {string} industry - Target industry
     * @param {string} role - Target role
     * @param {number} maxTokens - Maximum tokens
     * @param {number} temperature - Temperature setting
     * @returns {Object} Optimization result
     */
    async optimizeForIndustry(content, industry, role, maxTokens, temperature) {
        if (!industry) {
            // Fall back to general summarization
            return this.summarizeContent(content, maxTokens, temperature);
        }

        const prompt = this.analysisPrompts.industryOptimize
            .replace(/\{INDUSTRY\}/g, industry)
            .replace(/\{ROLE\}/g, role || 'business decision maker');

        try {
            const response = await this.openai.chat.completions.create({
                model: 'gpt-4o-mini',
                messages: [
                    { role: 'system', content: prompt },
                    { role: 'user', content: `Insurance content to optimize:\n\n${content}` }
                ],
                max_tokens: maxTokens,
                temperature: temperature
            });

            return {
                content: response.choices[0]?.message?.content || '',
                tokensUsed: response.usage?.total_tokens || 0
            };

        } catch (error) {
            throw new Error(`Industry optimization failed: ${error.message}`);
        }
    }

    /**
     * Optimize content specifically for email use
     * @param {string} content - Content to optimize
     * @param {number} maxTokens - Maximum tokens
     * @param {number} temperature - Temperature setting
     * @returns {Object} Email optimization result
     */
    async optimizeForEmail(content, maxTokens, temperature) {
        const prompt = this.analysisPrompts.emailOptimize;

        try {
            const response = await this.openai.chat.completions.create({
                model: 'gpt-4o-mini',
                messages: [
                    { role: 'system', content: prompt },
                    { role: 'user', content: `Transform this insurance content for email use:\n\n${content}` }
                ],
                max_tokens: maxTokens,
                temperature: temperature
            });

            return {
                content: response.choices[0]?.message?.content || '',
                tokensUsed: response.usage?.total_tokens || 0
            };

        } catch (error) {
            throw new Error(`Email optimization failed: ${error.message}`);
        }
    }

    /**
     * Batch analyze multiple content pieces efficiently
     * @param {Array} contentPieces - Array of content to analyze
     * @param {Object} options - Analysis options
     * @returns {Array} Array of analysis results
     */
    async batchAnalyze(contentPieces, options = {}) {
        const {
            batchSize = 3,
            delayBetweenBatches = 1000
        } = options;

        const results = [];
        
        // Process in batches to respect rate limits
        for (let i = 0; i < contentPieces.length; i += batchSize) {
            const batch = contentPieces.slice(i, i + batchSize);
            
            const batchPromises = batch.map(async (content, index) => {
                try {
                    // Add small delay to avoid rate limiting
                    await this.delay(index * 200);
                    return await this.analyzeContent(content.text, {
                        ...options,
                        type: content.type || 'summarize'
                    });
                } catch (error) {
                    return {
                        success: false,
                        error: error.message,
                        fallbackContent: this.createFallbackSummary(content.text)
                    };
                }
            });

            const batchResults = await Promise.all(batchPromises);
            results.push(...batchResults);

            // Delay between batches
            if (i + batchSize < contentPieces.length) {
                await this.delay(delayBetweenBatches);
            }
        }

        return results;
    }

    /**
     * Extract key insights from analyzed content
     * @param {string} analyzedContent - Content that was analyzed
     * @returns {Object} Extracted insights
     */
    extractInsights(analyzedContent) {
        const insights = {
            products: [],
            benefits: [],
            industries: [],
            keyTerms: []
        };

        // Extract products mentioned
        const productMatches = analyzedContent.match(/(?:liability|property|cyber|professional|directors|officers|employment|commercial|general|auto|workers|compensation|umbrella|excess)\s+(?:insurance|coverage|policy|protection)/gi);
        if (productMatches) {
            insights.products = [...new Set(productMatches.map(p => p.toLowerCase()))];
        }

        // Extract benefit keywords
        const benefitKeywords = ['protect', 'cover', 'reduce', 'mitigate', 'compliance', 'peace of mind', 'financial protection'];
        benefitKeywords.forEach(keyword => {
            if (analyzedContent.toLowerCase().includes(keyword)) {
                insights.benefits.push(keyword);
            }
        });

        // Extract industry mentions
        const industryKeywords = ['technology', 'manufacturing', 'healthcare', 'finance', 'retail', 'construction', 'professional services'];
        industryKeywords.forEach(industry => {
            if (analyzedContent.toLowerCase().includes(industry)) {
                insights.industries.push(industry);
            }
        });

        return insights;
    }

    /**
     * Calculate confidence score for analyzed content
     * @param {string} content - Analyzed content
     * @returns {number} Confidence score (0-1)
     */
    calculateConfidence(content) {
        let score = 0.5; // Base score

        // Check for structure and completeness
        if (content.includes('•') || content.includes('-')) score += 0.1; // Has bullet points
        if (content.length > 200) score += 0.1; // Adequate length
        if (content.length < 100) score -= 0.2; // Too short

        // Check for insurance-specific content
        const insuranceTerms = ['insurance', 'coverage', 'protection', 'policy', 'claim', 'premium'];
        const termCount = insuranceTerms.filter(term => content.toLowerCase().includes(term)).length;
        score += Math.min(termCount * 0.05, 0.2);

        // Check for business value language
        const valueTerms = ['benefit', 'protect', 'reduce', 'cost', 'risk', 'solution'];
        const valueCount = valueTerms.filter(term => content.toLowerCase().includes(term)).length;
        score += Math.min(valueCount * 0.03, 0.15);

        return Math.min(Math.max(score, 0), 1); // Clamp between 0 and 1
    }

    /**
     * Create fallback summary when AI analysis fails
     * @param {string} content - Original content
     * @returns {string} Fallback summary
     */
    createFallbackSummary(content) {
        // Extract first few sentences and clean them
        const sentences = content.split(/[.!?]+/).slice(0, 5);
        const cleanedSentences = sentences
            .map(s => s.trim())
            .filter(s => s.length > 20 && !s.toLowerCase().includes('disclaimer'))
            .slice(0, 3);

        return cleanedSentences.join('. ') + '.';
    }

    /**
     * Simple delay utility for rate limiting
     * @param {number} ms - Milliseconds to delay
     * @returns {Promise} Promise that resolves after delay
     */
    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Validate content quality before analysis
     * @param {string} content - Content to validate
     * @returns {Object} Validation result
     */
    validateContent(content) {
        const issues = [];
        
        if (!content || content.trim().length === 0) {
            issues.push('Content is empty');
        } else if (content.length < 50) {
            issues.push('Content too short for meaningful analysis');
        } else if (content.length > 10000) {
            issues.push('Content too long, may need pre-processing');
        }

        // Check for common PDF extraction issues
        if (content.includes('������') || content.includes('???')) {
            issues.push('Content contains extraction artifacts');
        }

        const wordCount = content.split(/\s+/).length;
        if (wordCount < 10) {
            issues.push('Insufficient word count for analysis');
        }

        return {
            isValid: issues.length === 0,
            issues,
            wordCount,
            characterCount: content.length
        };
    }
}

module.exports = ContentAnalyzer;