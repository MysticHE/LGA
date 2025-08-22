/**
 * PDF Content Processing Engine
 * 
 * Intelligent content extraction, cleaning, and optimization for insurance materials.
 * Converts raw PDF text into structured, relevant content for AI email generation.
 */

class PDFContentProcessor {
    constructor() {
        this.contentPatterns = {
            // Headers and footers to remove
            headers: [
                /^page \d+.*$/gmi,
                /^\d+\s*$/gm,
                /^.*confidential.*$/gmi,
                /^.*proprietary.*$/gmi,
                /^.*copyright.*$/gmi,
                /^.*all rights reserved.*$/gmi
            ],
            
            // Insurance-specific content patterns
            insurance: {
                products: /(?:liability|property|cyber|professional|directors|officers|employment|commercial|general|auto|workers|compensation|umbrella|excess)\s+(?:insurance|coverage|policy|protection)/gi,
                coverageTypes: /(?:coverage|protection|benefits|limits|deductibles|premiums|claims|exclusions)/gi,
                businessTerms: /(?:risk|exposure|compliance|regulations|audit|assessment|mitigation)/gi
            },
            
            // Section identifiers
            sections: {
                products: /(?:products?|services?|solutions?|offerings?)\s*:?/gi,
                benefits: /(?:benefits?|advantages?|features?)\s*:?/gi,
                coverage: /(?:coverage|protection|what'?s covered)\s*:?/gi,
                pricing: /(?:pricing|rates?|premiums?|costs?)\s*:?/gi,
                contact: /(?:contact|reach|call|email)\s*:?/gi
            },
            
            // Content to de-prioritize
            lowValue: [
                /^.*disclaimer.*$/gmi,
                /^.*terms and conditions.*$/gmi,
                /^.*legal notice.*$/gmi,
                /^.*regulatory.*$/gmi,
                /^\s*\d+\.\d+[\.\d]*\s/gm, // Legal numbering
                /^.*table of contents.*$/gmi
            ]
        };
        
        this.industryMappings = {
            'technology': ['cyber', 'professional', 'employment', 'directors'],
            'manufacturing': ['property', 'liability', 'workers', 'commercial'],
            'healthcare': ['professional', 'cyber', 'employment', 'liability'],
            'finance': ['cyber', 'professional', 'directors', 'employment'],
            'retail': ['property', 'liability', 'commercial', 'employment'],
            'construction': ['liability', 'workers', 'property', 'commercial'],
            'professional': ['professional', 'cyber', 'employment', 'directors'],
            'default': ['liability', 'property', 'professional', 'cyber']
        };
    }

    /**
     * Main processing method - converts raw PDF text to optimized content
     * @param {Array} materials - Array of material objects with content
     * @param {Object} leadContext - Lead information for content customization
     * @returns {Object} Processed and optimized content
     */
    async processContent(materials, leadContext = {}) {
        try {
            // Step 1: Clean and normalize all content
            const cleanedMaterials = materials.map(material => ({
                ...material,
                cleanedContent: this.cleanText(material.content),
                filename: material.filename
            }));

            // Step 2: Segment content into structured sections
            const segmentedContent = cleanedMaterials.map(material => ({
                ...material,
                segments: this.segmentContent(material.cleanedContent)
            }));

            // Step 3: Score and rank content segments
            const scoredContent = segmentedContent.map(material => ({
                ...material,
                segments: this.scoreSegments(material.segments, leadContext)
            }));

            // Step 4: Select and compile best content
            const optimizedContent = this.compileOptimalContent(scoredContent, leadContext);

            return {
                success: true,
                originalLength: materials.reduce((sum, m) => sum + m.content.length, 0),
                optimizedLength: optimizedContent.length,
                compressionRatio: this.calculateCompressionRatio(materials, optimizedContent),
                content: optimizedContent,
                metadata: this.generateMetadata(scoredContent, leadContext)
            };

        } catch (error) {
            console.error('Content processing error:', error);
            return {
                success: false,
                error: error.message,
                fallbackContent: this.createFallbackContent(materials)
            };
        }
    }

    /**
     * Clean and normalize raw PDF text
     * @param {string} rawText - Raw extracted PDF text
     * @returns {string} Cleaned text
     */
    cleanText(rawText) {
        let cleaned = rawText;

        // Remove headers and footers
        this.contentPatterns.headers.forEach(pattern => {
            cleaned = cleaned.replace(pattern, '');
        });

        // Normalize whitespace
        cleaned = cleaned
            .replace(/\r\n/g, '\n')  // Normalize line endings
            .replace(/\n{3,}/g, '\n\n')  // Collapse multiple newlines
            .replace(/[ \t]{2,}/g, ' ')  // Collapse multiple spaces
            .replace(/^\s+|\s+$/gm, '');  // Trim line whitespace

        // Preserve structured content markers
        cleaned = this.preserveStructure(cleaned);

        return cleaned;
    }

    /**
     * Preserve structured content like bullet points and sections
     * @param {string} text - Text to process
     * @returns {string} Text with preserved structure
     */
    preserveStructure(text) {
        return text
            .replace(/^[\s]*[•·‣⁃▪▫‣]\s*/gm, '• ')  // Normalize bullet points
            .replace(/^[\s]*\d+\.\s*/gm, (match, offset, string) => {
                // Keep numbered lists but clean formatting
                const num = match.match(/\d+/)[0];
                return `${num}. `;
            })
            .replace(/^([A-Z][A-Z\s&]{2,}):?\s*$/gm, '\n**$1**\n');  // Section headers
    }

    /**
     * Segment content into logical sections
     * @param {string} cleanedText - Cleaned text
     * @returns {Array} Array of content segments
     */
    segmentContent(cleanedText) {
        const segments = [];
        const lines = cleanedText.split('\n');
        let currentSegment = { type: 'general', content: '', confidence: 0 };

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;

            // Detect section type
            const sectionType = this.detectSectionType(line);
            
            if (sectionType !== currentSegment.type && currentSegment.content.length > 50) {
                // Save current segment and start new one
                segments.push(currentSegment);
                currentSegment = { type: sectionType, content: line + '\n', confidence: 0 };
            } else {
                currentSegment.content += line + '\n';
            }
        }

        // Add final segment
        if (currentSegment.content.length > 50) {
            segments.push(currentSegment);
        }

        return segments.filter(seg => seg.content.length > 30);  // Filter very short segments
    }

    /**
     * Detect the type of content section
     * @param {string} line - Line of text to analyze
     * @returns {string} Section type
     */
    detectSectionType(line) {
        const lowerLine = line.toLowerCase();

        // Check against section patterns
        for (const [type, pattern] of Object.entries(this.contentPatterns.sections)) {
            if (pattern.test(line)) {
                return type;
            }
        }

        // Check for insurance-specific content
        if (this.contentPatterns.insurance.products.test(line)) {
            return 'products';
        }
        if (this.contentPatterns.insurance.coverageTypes.test(line)) {
            return 'coverage';
        }
        if (this.contentPatterns.insurance.businessTerms.test(line)) {
            return 'business';
        }

        return 'general';
    }

    /**
     * Score content segments based on relevance and quality
     * @param {Array} segments - Content segments
     * @param {Object} leadContext - Lead context for scoring
     * @returns {Array} Scored segments
     */
    scoreSegments(segments, leadContext) {
        return segments.map(segment => {
            let score = 0;
            const content = segment.content.toLowerCase();

            // Base scoring by section type
            const typeScores = {
                'products': 10,
                'benefits': 8,
                'coverage': 9,
                'business': 7,
                'pricing': 6,
                'general': 3,
                'contact': 2
            };
            score += typeScores[segment.type] || 3;

            // Industry relevance scoring
            if (leadContext.industry) {
                const industryKey = this.getIndustryKey(leadContext.industry);
                const relevantProducts = this.industryMappings[industryKey] || this.industryMappings.default;
                
                relevantProducts.forEach(product => {
                    if (content.includes(product)) {
                        score += 5;
                    }
                });
            }

            // Insurance keyword density
            const insuranceMatches = (content.match(this.contentPatterns.insurance.products) || []).length;
            score += Math.min(insuranceMatches * 2, 10);

            // Penalize low-value content
            this.contentPatterns.lowValue.forEach(pattern => {
                if (pattern.test(segment.content)) {
                    score -= 3;
                }
            });

            // Content quality factors
            const wordCount = segment.content.split(/\s+/).length;
            if (wordCount < 20) score -= 2;  // Too short
            if (wordCount > 200) score += 2;  // Substantial content

            return {
                ...segment,
                score: Math.max(score, 0),  // Ensure non-negative score
                wordCount
            };
        });
    }

    /**
     * Compile optimal content from scored segments
     * @param {Array} scoredMaterials - Materials with scored segments
     * @param {Object} leadContext - Lead context
     * @returns {string} Optimized content string
     */
    compileOptimalContent(scoredMaterials, leadContext) {
        // Collect all segments and sort by score
        const allSegments = [];
        scoredMaterials.forEach(material => {
            material.segments.forEach(segment => {
                allSegments.push({
                    ...segment,
                    source: material.filename
                });
            });
        });

        // Sort by score (highest first)
        allSegments.sort((a, b) => b.score - a.score);

        // Build content within character limit
        const maxChars = 2500;  // Leave room for other prompt content
        let compiledContent = '';
        let currentLength = 0;
        const usedSources = new Set();

        // Group by source for better organization
        const contentBySource = {};
        
        for (const segment of allSegments) {
            if (segment.score < 3) break;  // Skip low-quality content
            
            const segmentLength = segment.content.length;
            if (currentLength + segmentLength > maxChars) {
                // Try to fit partial content if it's high-value
                if (segment.score >= 8 && currentLength < maxChars * 0.8) {
                    const remainingChars = maxChars - currentLength - 50;  // Buffer
                    const partialContent = this.truncateIntelligently(segment.content, remainingChars);
                    if (partialContent.length > 100) {
                        if (!contentBySource[segment.source]) contentBySource[segment.source] = [];
                        contentBySource[segment.source].push(partialContent);
                        currentLength += partialContent.length;
                    }
                }
                break;
            }

            if (!contentBySource[segment.source]) contentBySource[segment.source] = [];
            contentBySource[segment.source].push(segment.content);
            currentLength += segmentLength;
            usedSources.add(segment.source);
        }

        // Format final content
        for (const [source, contents] of Object.entries(contentBySource)) {
            compiledContent += `**${source}:**\n`;
            compiledContent += contents.join('\n\n') + '\n\n---\n\n';
        }

        return compiledContent.trim();
    }

    /**
     * Intelligently truncate content at sentence boundaries
     * @param {string} content - Content to truncate
     * @param {number} maxLength - Maximum length
     * @returns {string} Truncated content
     */
    truncateIntelligently(content, maxLength) {
        if (content.length <= maxLength) return content;

        // Try to cut at sentence boundaries
        const sentences = content.split(/[.!?]+/);
        let truncated = '';
        
        for (const sentence of sentences) {
            if (truncated.length + sentence.length + 1 > maxLength) {
                break;
            }
            truncated += sentence + '. ';
        }

        return truncated.trim() || content.substring(0, maxLength - 20) + '...';
    }

    /**
     * Get industry key for mapping
     * @param {string} industry - Industry string
     * @returns {string} Mapped industry key
     */
    getIndustryKey(industry) {
        const lowerIndustry = industry.toLowerCase();
        
        if (lowerIndustry.includes('tech') || lowerIndustry.includes('software')) return 'technology';
        if (lowerIndustry.includes('manufact') || lowerIndustry.includes('industrial')) return 'manufacturing';
        if (lowerIndustry.includes('health') || lowerIndustry.includes('medical')) return 'healthcare';
        if (lowerIndustry.includes('finance') || lowerIndustry.includes('bank')) return 'finance';
        if (lowerIndustry.includes('retail') || lowerIndustry.includes('consumer')) return 'retail';
        if (lowerIndustry.includes('construct') || lowerIndustry.includes('building')) return 'construction';
        if (lowerIndustry.includes('professional') || lowerIndustry.includes('services')) return 'professional';
        
        return 'default';
    }

    /**
     * Calculate compression ratio
     * @param {Array} originalMaterials - Original materials
     * @param {string} optimizedContent - Optimized content
     * @returns {number} Compression ratio
     */
    calculateCompressionRatio(originalMaterials, optimizedContent) {
        const originalLength = originalMaterials.reduce((sum, m) => sum + m.content.length, 0);
        return originalLength > 0 ? optimizedContent.length / originalLength : 0;
    }

    /**
     * Generate metadata about the processing
     * @param {Array} scoredMaterials - Processed materials
     * @param {Object} leadContext - Lead context
     * @returns {Object} Metadata object
     */
    generateMetadata(scoredMaterials, leadContext) {
        const totalSegments = scoredMaterials.reduce((sum, m) => sum + m.segments.length, 0);
        const highQualitySegments = scoredMaterials.reduce((sum, m) => 
            sum + m.segments.filter(s => s.score >= 7).length, 0);

        return {
            totalSegments,
            highQualitySegments,
            qualityRatio: totalSegments > 0 ? highQualitySegments / totalSegments : 0,
            industryOptimized: !!leadContext.industry,
            sourcesUsed: scoredMaterials.length,
            processingTimestamp: new Date().toISOString()
        };
    }

    /**
     * Create fallback content when processing fails
     * @param {Array} materials - Original materials
     * @returns {string} Fallback content
     */
    createFallbackContent(materials) {
        // Simple concatenation with basic cleaning as fallback
        const allContent = materials.map(m => `${m.filename}:\n${m.content}`).join('\n\n---\n\n');
        return this.cleanText(allContent).substring(0, 3000) + 
               (allContent.length > 3000 ? '\n\n[Content truncated - processing error]' : '');
    }
}

module.exports = PDFContentProcessor;