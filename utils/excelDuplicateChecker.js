/**
 * Excel-Based Duplicate Checker
 * Uses Excel as the single source of truth for duplicate prevention
 * More reliable than separate campaign state management
 */

const { getLeadsViaGraphAPI } = require('./excelGraphAPI');

class ExcelDuplicateChecker {
    constructor() {
        this.cache = new Map(); // Cache lead data for performance
        this.cacheExpiry = 5 * 60 * 1000; // 5 minutes cache
    }

    /**
     * Check if an email has already been sent by looking at Excel data
     * @param {object} graphClient - Microsoft Graph client
     * @param {string} email - Email address to check
     * @returns {Promise<{alreadySent: boolean, leadData: object|null, reason: string}>}
     */
    async isEmailAlreadySent(graphClient, email) {
        try {
            console.log(`🔍 EXCEL DUPLICATE CHECK: Checking if ${email} has already been sent...`);

            // Get fresh Excel data
            const allLeads = await this.getLeadsWithCache(graphClient);
            
            if (!allLeads) {
                console.log(`⚠️ Could not retrieve Excel data for duplicate check`);
                return {
                    alreadySent: false,
                    leadData: null,
                    reason: 'Excel data not available - allowing send'
                };
            }

            // Find the lead by email
            const lead = allLeads.find(l => 
                l.Email && l.Email.toLowerCase().trim() === email.toLowerCase().trim()
            );

            if (!lead) {
                console.log(`❌ Email ${email} not found in Excel - cannot send`);
                return {
                    alreadySent: true, // Prevent sending if not in Excel
                    leadData: null,
                    reason: 'Email not found in Excel master list'
                };
            }

            // Check various "sent" indicators
            const sentIndicators = this.checkSentIndicators(lead);
            
            if (sentIndicators.alreadySent) {
                console.log(`⚠️ DUPLICATE PREVENTED: ${email} - ${sentIndicators.reason}`);
                console.log(`📊 Lead status: Status=${lead.Status}, Last_Email_Date=${lead.Last_Email_Date}, Email_Count=${lead.Email_Count}`);
                
                return {
                    alreadySent: true,
                    leadData: lead,
                    reason: sentIndicators.reason
                };
            }

            console.log(`✅ SAFE TO SEND: ${email} - No previous send indicators found`);
            return {
                alreadySent: false,
                leadData: lead,
                reason: 'No previous send indicators found'
            };

        } catch (error) {
            console.error(`❌ Excel duplicate check error for ${email}:`, error.message);
            
            // Fail safe - if we can't check, don't send
            return {
                alreadySent: true,
                leadData: null,
                reason: `Excel check failed: ${error.message}`
            };
        }
    }

    /**
     * Check multiple indicators to determine if email was already sent
     * @param {object} lead - Lead object from Excel
     * @returns {object} {alreadySent: boolean, reason: string}
     */
    checkSentIndicators(lead) {
        const today = new Date().toISOString().split('T')[0];

        // Priority 1: Check if Status indicates already sent
        if (lead.Status) {
            const status = lead.Status.toString().toLowerCase();
            if (['sent', 'read', 'replied', 'clicked'].includes(status)) {
                return {
                    alreadySent: true,
                    reason: `Status is '${lead.Status}' - already processed`
                };
            }
        }

        // Priority 2: Check Last_Email_Date (most reliable indicator)
        if (lead.Last_Email_Date) {
            const lastEmailDate = this.parseExcelDate(lead.Last_Email_Date);
            if (lastEmailDate) {
                return {
                    alreadySent: true,
                    reason: `Last_Email_Date is ${lastEmailDate} - email already sent`
                };
            }
        }

        // Priority 3: Check Email_Count > 0
        if (lead.Email_Count && parseInt(lead.Email_Count) > 0) {
            return {
                alreadySent: true,
                reason: `Email_Count is ${lead.Email_Count} - emails already sent`
            };
        }

        // Priority 4: 'Email Sent' column removed - using Status field instead
        // This check is now handled by Priority 2: Status-based check

        // Priority 5: 'Sent Date' column removed - using other indicators
        // This check was removed as 'Sent Date' is no longer updated by the system

        // Priority 6: Check Template_Used (indicates email was generated/sent)
        if (lead.Template_Used && lead.Template_Used !== '' && lead.Template_Used !== 'None') {
            return {
                alreadySent: true,
                reason: `Template_Used is '${lead.Template_Used}' - email already sent`
            };
        }

        return {
            alreadySent: false,
            reason: 'No sent indicators found'
        };
    }

    /**
     * Get leads with caching for performance
     * @param {object} graphClient - Microsoft Graph client
     * @returns {Promise<Array|null>}
     */
    async getLeadsWithCache(graphClient) {
        const now = new Date().getTime();
        const cacheKey = 'leads_data';

        // Check cache first
        const cached = this.cache.get(cacheKey);
        if (cached && (now - cached.timestamp) < this.cacheExpiry) {
            console.log(`📋 Using cached leads data (${cached.data.length} leads)`);
            return cached.data;
        }

        // Fetch fresh data
        console.log(`📋 Fetching fresh leads data from Excel...`);
        const leads = await getLeadsViaGraphAPI(graphClient);
        
        if (leads) {
            // Cache the data
            this.cache.set(cacheKey, {
                data: leads,
                timestamp: now
            });
            console.log(`📋 Cached ${leads.length} leads for duplicate checking`);
        }

        return leads;
    }

    /**
     * Clear the cache (useful for testing or manual refresh)
     */
    clearCache() {
        this.cache.clear();
        console.log('🧹 Excel duplicate checker cache cleared');
    }

    /**
     * Parse Excel date values (handles both serial numbers and date strings)
     * @param {*} dateValue - Date value from Excel
     * @returns {string|null} ISO date string or null
     */
    parseExcelDate(dateValue) {
        if (!dateValue) return null;
        
        try {
            // Handle Excel serial numbers (like 45907)
            if (typeof dateValue === 'number' && dateValue > 40000) {
                const excelEpoch = new Date(1900, 0, 1);
                const jsDate = new Date(excelEpoch.getTime() + (dateValue - 2) * 24 * 60 * 60 * 1000);
                return jsDate.toISOString().split('T')[0];
            } else {
                // Handle regular date strings
                return new Date(dateValue).toISOString().split('T')[0];
            }
        } catch (error) {
            return null;
        }
    }

    /**
     * Get comprehensive duplicate check report for debugging
     * @param {object} graphClient - Microsoft Graph client
     * @param {Array} emails - Array of email addresses to check
     * @returns {Promise<object>} Detailed report
     */
    async getDuplicateReport(graphClient, emails) {
        const report = {
            totalChecked: emails.length,
            alreadySent: 0,
            safeToSend: 0,
            errors: 0,
            details: []
        };

        for (const email of emails) {
            try {
                const result = await this.isEmailAlreadySent(graphClient, email);
                
                report.details.push({
                    email: email,
                    alreadySent: result.alreadySent,
                    reason: result.reason,
                    leadData: result.leadData ? {
                        Status: result.leadData.Status,
                        Last_Email_Date: result.leadData.Last_Email_Date,
                        Email_Count: result.leadData.Email_Count
                    } : null
                });

                if (result.alreadySent) {
                    report.alreadySent++;
                } else {
                    report.safeToSend++;
                }

            } catch (error) {
                report.errors++;
                report.details.push({
                    email: email,
                    error: error.message
                });
            }
        }

        return report;
    }
}

// Export singleton instance
module.exports = new ExcelDuplicateChecker();