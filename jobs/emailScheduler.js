const cron = require('node-cron');
const axios = require('axios');
const { getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const { advancedExcelUpload } = require('../routes/excel-upload-fix');

/**
 * Background Email Scheduler
 * Automated email sending based on master Excel file schedules
 */

class EmailScheduler {
    constructor() {
        this.isRunning = false;
        this.lastRun = null;
        this.schedule = '0 */1 * * *'; // Every hour
        this.baseURL = process.env.RENDER_EXTERNAL_URL || 'http://localhost:3000';
        this.excelProcessor = new ExcelProcessor();
        this.authProvider = getDelegatedAuthProvider();
        
        console.log('📅 Email Scheduler initialized');
    }

    /**
     * Start the automated email scheduler (DISABLED - Manual control only)
     */
    start() {
        console.log('📅 Email Scheduler set to MANUAL MODE - automatic scheduling disabled');
        console.log('💡 Use manual triggers via frontend or API endpoints to send emails');
        this.isRunning = false;
    }

    /**
     * Stop the email scheduler
     */
    stop() {
        if (!this.isRunning) {
            console.log('⚠️ Email Scheduler is not running');
            return;
        }

        if (this.cronJob) {
            this.cronJob.stop();
            this.cronJob.destroy();
        }

        this.isRunning = false;
        console.log('🛑 Email Scheduler stopped');
    }

    /**
     * Process scheduled emails for all authenticated sessions
     */
    async processScheduledEmails() {
        try {
            console.log('⏰ Processing scheduled emails...');
            this.lastRun = new Date();

            // Get all active sessions
            const activeSessions = this.authProvider.getActiveSessions();
            
            if (activeSessions.length === 0) {
                console.log('📭 No active sessions found, skipping scheduled email processing');
                return;
            }

            console.log(`👥 Processing scheduled emails for ${activeSessions.length} active sessions`);

            const results = [];
            
            for (const sessionId of activeSessions) {
                try {
                    const sessionResult = await this.processSessionScheduledEmails(sessionId);
                    results.push({
                        sessionId,
                        success: true,
                        ...sessionResult
                    });
                } catch (sessionError) {
                    console.error(`❌ Error processing session ${sessionId}:`, sessionError.message);
                    results.push({
                        sessionId,
                        success: false,
                        error: sessionError.message
                    });
                }
            }

            // Log summary
            const successful = results.filter(r => r.success).length;
            const totalEmails = results.reduce((sum, r) => sum + (r.emailsSent || 0), 0);
            
            console.log(`✅ Scheduled email processing completed: ${successful}/${activeSessions.length} sessions, ${totalEmails} emails sent`);

        } catch (error) {
            console.error('❌ Scheduled email processing error:', error);
        }
    }

    /**
     * Process scheduled emails for a specific session
     */
    async processSessionScheduledEmails(sessionId) {
        try {
            console.log(`📧 Processing scheduled emails for session: ${sessionId}`);

            // Get Graph client for this session
            const graphClient = await this.authProvider.getGraphClient(sessionId);
            
            if (!graphClient) {
                throw new Error('Unable to get Graph client for session');
            }

            // Download master file for this session
            const masterWorkbook = await this.downloadMasterFile(graphClient);
            
            if (!masterWorkbook) {
                console.log(`📋 No master file found for session: ${sessionId}`);
                return { emailsSent: 0, leadsProcessed: 0 };
            }

            // Get leads due for email today
            const dueLeads = this.excelProcessor.getLeadsDueToday(masterWorkbook);
            
            if (dueLeads.length === 0) {
                console.log(`📭 No leads due for email in session: ${sessionId}`);
                return { emailsSent: 0, leadsProcessed: 0 };
            }

            console.log(`📋 Found ${dueLeads.length} leads due for email in session: ${sessionId}`);

            // Get templates for email processing
            const templates = this.excelProcessor.getTemplates(masterWorkbook);

            // Process emails in batches to respect rate limits
            const batchSize = 5;
            let emailsSent = 0;
            let leadsProcessed = 0;
            const errors = [];

            for (let i = 0; i < dueLeads.length; i += batchSize) {
                const batch = dueLeads.slice(i, i + batchSize);
                
                const batchResults = await this.processBatch(
                    graphClient, 
                    batch, 
                    templates, 
                    sessionId
                );
                
                emailsSent += batchResults.emailsSent;
                leadsProcessed += batchResults.leadsProcessed;
                errors.push(...batchResults.errors);

                // Update master file with batch results
                if (batchResults.updates.length > 0) {
                    await this.updateMasterFileWithResults(
                        graphClient, 
                        masterWorkbook, 
                        batchResults.updates
                    );
                }

                // Add delay between batches to respect rate limits
                if (i + batchSize < dueLeads.length) {
                    await this.delay(2000); // 2 second delay
                }
            }

            console.log(`✅ Session ${sessionId} processing complete: ${emailsSent} emails sent, ${leadsProcessed} leads processed`);

            return {
                emailsSent,
                leadsProcessed,
                errors: errors.length,
                errorDetails: errors.slice(0, 5) // Limit error details
            };

        } catch (error) {
            console.error(`❌ Session processing error for ${sessionId}:`, error);
            throw error;
        }
    }

    /**
     * Process a batch of leads
     */
    async processBatch(graphClient, leads, templates, sessionId) {
        const EmailContentProcessor = require('../utils/emailContentProcessor');
        const emailContentProcessor = new EmailContentProcessor();
        
        const results = {
            emailsSent: 0,
            leadsProcessed: 0,
            errors: [],
            updates: []
        };

        for (const lead of leads) {
            try {
                console.log(`📧 Processing lead: ${lead.Email} (${lead.Name})`);
                results.leadsProcessed++;

                // Determine email content type
                const emailChoice = lead.Email_Choice || 'AI_Generated';

                // Process email content
                const emailContent = await emailContentProcessor.processEmailContent(
                    lead, 
                    emailChoice, 
                    templates
                );

                // Validate email content
                const validation = emailContentProcessor.validateEmailContent(emailContent);
                if (!validation.isValid) {
                    results.errors.push({
                        email: lead.Email,
                        error: 'Invalid email content: ' + validation.errors.join(', ')
                    });
                    continue;
                }

                // Send email using Microsoft Graph
                const emailMessage = {
                    subject: emailContent.subject,
                    body: {
                        contentType: 'HTML',
                        content: emailContentProcessor.convertToHTML(emailContent, lead.Email)
                    },
                    toRecipients: [
                        {
                            emailAddress: {
                                address: lead.Email,
                                name: lead.Name
                            }
                        }
                    ]
                };

                await graphClient.api('/me/sendMail').post({
                    message: emailMessage,
                    saveToSentItems: true
                });

                results.emailsSent++;

                // Prepare update for master file
                const updates = {
                    Status: 'Sent',
                    Last_Email_Date: new Date().toISOString().split('T')[0],
                    Email_Count: (lead.Email_Count || 0) + 1,
                    Template_Used: emailContent.contentType,
                    Email_Content_Sent: emailContent.subject + '\n\n' + emailContent.body,
                    Next_Email_Date: this.excelProcessor.calculateNextEmailDate(
                        new Date(), 
                        lead.Follow_Up_Days || 7
                    ),
                    'Email Sent': 'Yes',
                    'Email Status': 'Sent',
                    'Sent Date': new Date().toISOString(),
                    'Last Updated': new Date().toISOString()
                };

                results.updates.push({
                    email: lead.Email,
                    updates: updates
                });

                console.log(`✅ Email sent successfully to: ${lead.Email}`);

                // Small delay between emails
                await this.delay(500);

            } catch (error) {
                console.error(`❌ Failed to process lead ${lead.Email}:`, error.message);
                results.errors.push({
                    email: lead.Email,
                    error: error.message
                });
            }
        }

        return results;
    }

    /**
     * Update master file with email sending results
     */
    async updateMasterFileWithResults(graphClient, masterWorkbook, updates) {
        try {
            for (const update of updates) {
                masterWorkbook = this.excelProcessor.updateLeadInMaster(
                    masterWorkbook, 
                    update.email, 
                    update.updates
                );
            }

            // Save updated master file
            const masterBuffer = this.excelProcessor.workbookToBuffer(masterWorkbook);
            await advancedExcelUpload(
                graphClient, 
                masterBuffer, 
                'LGA-Master-Email-List.xlsx', 
                '/LGA-Email-Automation'
            );

            console.log(`📊 Updated master file with ${updates.length} lead updates`);

        } catch (error) {
            console.error('❌ Error updating master file:', error);
            throw error;
        }
    }

    /**
     * Download master file from OneDrive
     */
    async downloadMasterFile(graphClient) {
        try {
            const masterFileName = 'LGA-Master-Email-List.xlsx';
            const masterFolderPath = '/LGA-Email-Automation';
            
            const files = await graphClient
                .api(`/me/drive/root:${masterFolderPath}:/children`)
                .filter(`name eq '${masterFileName}'`)
                .get();

            if (files.value.length === 0) {
                return null;
            }

            const fileContent = await graphClient
                .api(`/me/drive/items/${files.value[0].id}/content`)
                .get();

            return this.excelProcessor.bufferToWorkbook(fileContent);
        } catch (error) {
            console.error('❌ Master file download error:', error);
            return null;
        }
    }


    /**
     * Manual trigger for scheduled email processing
     */
    async triggerManualProcessing() {
        console.log('🔧 Manual trigger of scheduled email processing...');
        await this.processScheduledEmails();
    }

    /**
     * Get scheduler status
     */
    getStatus() {
        return {
            isRunning: this.isRunning,
            schedule: this.schedule,
            lastRun: this.lastRun,
            activeSessions: this.authProvider.getActiveSessions().length
        };
    }

    /**
     * Update scheduler configuration
     */
    updateSchedule(newSchedule) {
        if (this.isRunning) {
            this.stop();
        }
        
        this.schedule = newSchedule;
        console.log(`📅 Updated scheduler to cron: ${this.schedule}`);
        
        this.start();
    }

    /**
     * Utility function for delays
     */
    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * Get processing statistics
     */
    async getProcessingStats() {
        try {
            const activeSessions = this.authProvider.getActiveSessions();
            const stats = {
                activeSessions: activeSessions.length,
                totalLeadsDue: 0,
                lastProcessingTime: this.lastRun,
                isRunning: this.isRunning,
                schedule: this.schedule
            };

            // Get due leads count for all sessions
            for (const sessionId of activeSessions) {
                try {
                    const graphClient = await this.authProvider.getGraphClient(sessionId);
                    if (graphClient) {
                        const masterWorkbook = await this.downloadMasterFile(graphClient);
                        if (masterWorkbook) {
                            const dueLeads = this.excelProcessor.getLeadsDueToday(masterWorkbook);
                            stats.totalLeadsDue += dueLeads.length;
                        }
                    }
                } catch (error) {
                    console.error(`Error getting stats for session ${sessionId}:`, error.message);
                }
            }

            return stats;
        } catch (error) {
            console.error('Error getting processing stats:', error);
            return {
                activeSessions: 0,
                totalLeadsDue: 0,
                lastProcessingTime: this.lastRun,
                isRunning: this.isRunning,
                schedule: this.schedule,
                error: error.message
            };
        }
    }

    /**
     * Clean up expired sessions
     */
    cleanupExpiredSessions() {
        try {
            const cleanedCount = this.authProvider.cleanupExpiredSessions();
            if (cleanedCount > 0) {
                console.log(`🧹 Cleaned up ${cleanedCount} expired sessions`);
            }
        } catch (error) {
            console.error('Error cleaning up expired sessions:', error);
        }
    }
}

// Create and export singleton instance
const emailScheduler = new EmailScheduler();

module.exports = emailScheduler;