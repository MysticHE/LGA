/**
 * Email Bounce Detection Utilities
 * Monitors inbox for bounce notifications and updates master list
 */

class BounceDetector {
    constructor() {
        // Common bounce detection patterns
        this.bouncePatterns = {
            // Subject line patterns
            subjects: [
                /undelivered mail returned to sender/i,
                /delivery status notification/i,
                /returned mail/i,
                /mail delivery failed/i,
                /undeliverable/i,
                /bounce/i,
                /delivery failure/i,
                /message not delivered/i,
                /mail delivery subsystem/i,
                /automatic reply/i,
                /delivery has failed to these recipients/i, // Microsoft Outlook
                /delivery failed/i
            ],
            
            // Sender patterns (bounce notifications usually come from these)
            senders: [
                /@.*postmaster/i,
                /@.*mailer-daemon/i,
                /@.*noreply/i,
                /@.*no-reply/i,
                /delivery-status/i,
                /bounce/i,
                /mail.*delivery/i,
                /outlook@microsoft\.com/i, // Microsoft Outlook bounces
                /microsoftexchange/i
            ],
            
            // Body content patterns for bounce reasons
            reasons: {
                // Hard bounces (permanent failures)
                hard: [
                    { pattern: /user unknown/i, reason: 'User does not exist' },
                    { pattern: /no such user/i, reason: 'User does not exist' },
                    { pattern: /invalid recipient/i, reason: 'Invalid recipient address' },
                    { pattern: /recipient address rejected/i, reason: 'Recipient address rejected' },
                    { pattern: /mailbox unavailable/i, reason: 'Mailbox unavailable' },
                    { pattern: /domain not found/i, reason: 'Domain does not exist' },
                    { pattern: /host unknown/i, reason: 'Host unknown' },
                    { pattern: /permanent failure/i, reason: 'Permanent delivery failure' },
                    { pattern: /550/i, reason: 'Mailbox unavailable (550)' },
                    { pattern: /551/i, reason: 'User not local (551)' },
                    { pattern: /553/i, reason: 'Invalid address syntax (553)' },
                    { pattern: /554/i, reason: 'Transaction failed (554)' }
                ],
                
                // Soft bounces (temporary failures)
                soft: [
                    { pattern: /mailbox full/i, reason: 'Mailbox full' },
                    { pattern: /over quota/i, reason: 'Mailbox over quota' },
                    { pattern: /temporary failure/i, reason: 'Temporary failure' },
                    { pattern: /try again later/i, reason: 'Temporary failure, try again later' },
                    { pattern: /service unavailable/i, reason: 'Service temporarily unavailable' },
                    { pattern: /421/i, reason: 'Service not available (421)' },
                    { pattern: /450/i, reason: 'Requested action not taken (450)' },
                    { pattern: /451/i, reason: 'Requested action aborted (451)' },
                    { pattern: /452/i, reason: 'Insufficient storage (452)' }
                ],
                
                // Temporary issues
                temporary: [
                    { pattern: /greylisted/i, reason: 'Greylisted - will retry' },
                    { pattern: /deferred/i, reason: 'Message deferred' },
                    { pattern: /queue/i, reason: 'Queued for retry' },
                    { pattern: /timeout/i, reason: 'Connection timeout' }
                ]
            }
        };
    }

    /**
     * Check if an email message is a bounce notification
     * @param {object} message - Email message from Microsoft Graph
     * @returns {object|null} Bounce information or null if not a bounce
     */
    detectBounce(message) {
        const subject = message.subject || '';
        const sender = message.from?.emailAddress?.address || '';
        const body = message.body?.content || '';
        
        // Check if this looks like a bounce notification
        const isBounceSubject = this.bouncePatterns.subjects.some(pattern => pattern.test(subject));
        const isBounceSender = this.bouncePatterns.senders.some(pattern => pattern.test(sender));
        
        if (!isBounceSubject && !isBounceSender) {
            return null;
        }
        
        console.log(`🔍 Potential bounce detected: ${subject} from ${sender}`);
        
        // Extract original recipient email from bounce message
        const originalRecipient = this.extractOriginalRecipient(body, subject);
        
        if (!originalRecipient) {
            console.log(`⚠️ Could not extract original recipient from bounce message`);
            return null;
        }
        
        // Analyze bounce type and reason
        const bounceAnalysis = this.analyzeBounceReason(body);
        
        return {
            originalRecipient: originalRecipient,
            bounceType: bounceAnalysis.type,
            bounceReason: bounceAnalysis.reason,
            bounceDate: message.receivedDateTime || new Date().toISOString(),
            bounceSubject: subject,
            bounceSender: sender,
            messageId: message.id
        };
    }

    /**
     * Extract the original recipient email from bounce message
     * @param {string} body - Bounce message body
     * @param {string} subject - Bounce message subject
     * @returns {string|null} Original recipient email
     */
    extractOriginalRecipient(body, subject) {
        // Common patterns for extracting email addresses from bounce messages
        const emailPatterns = [
            // Standard bounce format: "The following address(es) failed:"
            /following.*(?:address|recipient).*failed[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i,
            
            // Postfix format: "to <email@domain.com>"
            /to\s*[<"]([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})[>"]/i,
            
            // Exchange format: "The following recipient(s) could not be reached:"
            /recipient.*could not be reached[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i,
            
            // Generic patterns
            /original.*recipient[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i,
            /failed.*delivery.*to[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i,
            
            // Microsoft Outlook format: "Name (email@domain.com)"
            /\(([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})\)/i,
            
            // Extract from subject line as fallback
            /undelivered.*to[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i
        ];
        
        // Try body patterns first
        for (const pattern of emailPatterns) {
            const match = body.match(pattern);
            if (match) {
                return match[1].toLowerCase();
            }
        }
        
        // Try subject line patterns
        for (const pattern of emailPatterns) {
            const match = subject.match(pattern);
            if (match) {
                return match[1].toLowerCase();
            }
        }
        
        // Fallback: extract any email address from the bounce message
        const generalEmailPattern = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
        const emails = body.match(generalEmailPattern) || [];
        
        // Filter out common bounce system emails
        const filteredEmails = emails.filter(email => 
            !email.includes('postmaster') &&
            !email.includes('mailer-daemon') &&
            !email.includes('noreply') &&
            !email.includes('no-reply') &&
            !email.includes('bounce')
        );
        
        return filteredEmails.length > 0 ? filteredEmails[0].toLowerCase() : null;
    }

    /**
     * Analyze bounce message to determine type and reason
     * @param {string} body - Bounce message body
     * @returns {object} Bounce analysis
     */
    analyzeBounceReason(body) {
        const bodyLower = body.toLowerCase();
        
        // Check for hard bounces first (permanent failures)
        for (const bounceRule of this.bouncePatterns.reasons.hard) {
            if (bounceRule.pattern.test(bodyLower)) {
                return {
                    type: 'Hard',
                    reason: bounceRule.reason
                };
            }
        }
        
        // Check for soft bounces (temporary failures)
        for (const bounceRule of this.bouncePatterns.reasons.soft) {
            if (bounceRule.pattern.test(bodyLower)) {
                return {
                    type: 'Soft',
                    reason: bounceRule.reason
                };
            }
        }
        
        // Check for temporary issues
        for (const bounceRule of this.bouncePatterns.reasons.temporary) {
            if (bounceRule.pattern.test(bodyLower)) {
                return {
                    type: 'Temporary',
                    reason: bounceRule.reason
                };
            }
        }
        
        // Default to soft bounce if we can't determine the specific type
        return {
            type: 'Soft',
            reason: 'Unknown bounce reason'
        };
    }

    /**
     * Check inbox for bounce notifications using Microsoft Graph
     * @param {object} graphClient - Microsoft Graph client
     * @param {number} hoursBack - How many hours back to check (default: 24)
     * @returns {Array} Array of detected bounces
     */
    async checkInboxForBounces(graphClient, hoursBack = 24) {
        try {
            console.log(`🔍 Checking inbox for bounce notifications (last ${hoursBack} hours)...`);
            
            // Calculate date filter
            const cutoffDate = new Date();
            cutoffDate.setHours(cutoffDate.getHours() - hoursBack);
            const dateFilter = cutoffDate.toISOString();
            
            // Query inbox for recent messages that might be bounces
            const messages = await graphClient
                .api('/me/messages')
                .filter(`receivedDateTime ge ${dateFilter}`)
                .select('id,subject,from,receivedDateTime,body')
                .top(100)
                .get();
            
            const bounces = [];
            
            for (const message of messages.value) {
                const bounceInfo = this.detectBounce(message);
                if (bounceInfo) {
                    bounces.push(bounceInfo);
                }
            }
            
            console.log(`📋 Found ${bounces.length} bounce notifications in the last ${hoursBack} hours`);
            return bounces;
            
        } catch (error) {
            console.error('❌ Error checking inbox for bounces:', error.message);
            throw error;
        }
    }

    /**
     * Process detected bounces and update master list
     * @param {Array} bounces - Array of detected bounces
     * @param {Function} updateLeadCallback - Callback to update lead in master list
     */
    async processBounces(bounces, updateLeadCallback) {
        console.log(`📧 Processing ${bounces.length} detected bounces...`);
        
        const results = {
            processed: 0,
            bounced: 0,
            errors: []
        };
        
        for (const bounce of bounces) {
            try {
                // Simple bounce tracking - just mark as bounced
                const updates = {
                    'Email Bounce': 'Yes',
                    'Status': 'Bounced',
                    'Last Updated': new Date().toISOString()
                };
                
                // Update the lead via callback
                const updateSuccess = await updateLeadCallback(bounce.originalRecipient, updates);
                
                if (updateSuccess) {
                    results.processed++;
                    results.bounced++;
                    console.log(`✅ Marked ${bounce.originalRecipient} as bounced: ${bounce.bounceReason}`);
                } else {
                    console.log(`⚠️ Could not find lead for bounced email: ${bounce.originalRecipient}`);
                }
                
            } catch (error) {
                console.error(`❌ Error processing bounce for ${bounce.originalRecipient}: ${error.message}`);
                results.errors.push({
                    email: bounce.originalRecipient,
                    error: error.message
                });
            }
        }
        
        return results;
    }

    /**
     * Get bounce statistics for reporting
     * @param {Array} leads - Array of leads to analyze
     * @returns {object} Bounce statistics
     */
    getBounceStatistics(leads) {
        const stats = {
            totalLeads: leads.length,
            bouncedEmails: 0,
            validEmails: 0,
            bounceRate: 0
        };
        
        for (const lead of leads) {
            // Count bounces
            if (lead['Email Bounce'] === 'Yes') {
                stats.bouncedEmails++;
            } else {
                stats.validEmails++;
            }
        }
        
        // Calculate bounce rate
        stats.bounceRate = stats.totalLeads > 0 ? 
            (stats.bouncedEmails / stats.totalLeads * 100).toFixed(2) : 0;
        
        return stats;
    }
}

module.exports = BounceDetector;