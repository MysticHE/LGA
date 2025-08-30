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
        
        // Start webhook renewal job
        this.startWebhookRenewalJob();
        
        // Start reply detection job
        this.startReplyDetectionJob();
        
        console.log('📅 Email Scheduler initialized with webhook renewal and reply detection');
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

        if (this.webhookRenewalJob) {
            this.webhookRenewalJob.stop();
            this.webhookRenewalJob.destroy();
        }

        if (this.replyDetectionJob) {
            this.replyDetectionJob.stop();
            this.replyDetectionJob.destroy();
        }

        this.isRunning = false;
        console.log('🛑 Email Scheduler stopped');
    }

    /**
     * Start webhook subscription renewal job
     * Runs every 20 hours to renew 24-hour subscriptions before expiry
     */
    startWebhookRenewalJob() {
        // Run every 20 hours (4 hours before expiry)
        this.webhookRenewalJob = cron.schedule('0 */20 * * *', async () => {
            await this.renewWebhookSubscriptions();
        }, {
            scheduled: true,
            timezone: "Asia/Singapore"
        });
        
        console.log('🔄 Webhook renewal job started (runs every 20 hours)');
    }

    /**
     * Start reply detection job
     * Runs every 5 minutes to check inbox for replies to sent emails
     */
    startReplyDetectionJob() {
        // Run every 5 minutes
        this.replyDetectionJob = cron.schedule('*/5 * * * *', async () => {
            await this.checkInboxForReplies();
        }, {
            scheduled: true,
            timezone: "Asia/Singapore"
        });
        
        console.log('💬 Reply detection job started (runs every 5 minutes)');
    }

    /**
     * Renew webhook subscriptions for all active sessions
     */
    async renewWebhookSubscriptions() {
        try {
            console.log('🔄 Starting webhook subscription renewal...');
            
            const activeSessions = this.authProvider.getActiveSessions();
            
            if (activeSessions.length === 0) {
                console.log('📭 No active sessions for webhook renewal');
                return;
            }
            
            const renewalResults = [];
            
            for (const sessionId of activeSessions) {
                try {
                    const result = await this.renewSessionWebhooks(sessionId);
                    renewalResults.push({ sessionId, success: true, ...result });
                } catch (error) {
                    console.error(`❌ Webhook renewal failed for session ${sessionId}:`, error.message);
                    renewalResults.push({ sessionId, success: false, error: error.message });
                }
            }
            
            const successful = renewalResults.filter(r => r.success).length;
            console.log(`✅ Webhook renewal completed: ${successful}/${activeSessions.length} sessions`);
            
        } catch (error) {
            console.error('❌ Webhook renewal job error:', error);
        }
    }

    /**
     * Renew webhooks for a specific session
     */
    async renewSessionWebhooks(sessionId) {
        try {
            const response = await axios.post(`${this.baseURL}/api/email/webhook/auto-setup`, {}, {
                headers: {
                    'X-Session-Id': sessionId,
                    'Content-Type': 'application/json'
                },
                timeout: 10000
            });
            
            if (response.data.success) {
                console.log(`✅ Webhook renewed for session ${sessionId}: ${response.data.subscriptionId}`);
                return {
                    subscriptionId: response.data.subscriptionId,
                    expirationDateTime: response.data.expirationDateTime
                };
            } else {
                throw new Error(response.data.message || 'Unknown renewal error');
            }
            
        } catch (error) {
            if (error.response?.status === 400 && error.response?.data?.message?.includes('already active')) {
                console.log(`📡 Webhook subscriptions already active for session ${sessionId}`);
                return { message: 'Already active' };
            }
            throw error;
        }
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
     * Check inbox for replies to sent emails across all active sessions
     */
    async checkInboxForReplies() {
        try {
            console.log('💬 Checking inbox for replies...');
            
            const activeSessions = this.authProvider.getActiveSessions();
            
            if (activeSessions.length === 0) {
                console.log('📭 No active sessions for reply detection');
                return;
            }
            
            let totalRepliesFound = 0;
            
            for (const sessionId of activeSessions) {
                try {
                    const repliesFound = await this.checkSessionInboxForReplies(sessionId);
                    totalRepliesFound += repliesFound;
                } catch (sessionError) {
                    console.error(`❌ Reply detection failed for session ${sessionId}:`, sessionError.message);
                }
            }
            
            if (totalRepliesFound > 0) {
                console.log(`✅ Reply detection completed: ${totalRepliesFound} replies found`);
            }
            
        } catch (error) {
            console.error('❌ Reply detection job error:', error);
        }
    }

    /**
     * Check inbox for replies in a specific session
     */
    async checkSessionInboxForReplies(sessionId) {
        try {
            console.log(`💬 Checking replies for session: ${sessionId}`);
            
            const graphClient = await this.authProvider.getGraphClient(sessionId);
            
            if (!graphClient) {
                console.log(`❌ Unable to get Graph client for session: ${sessionId}`);
                return 0;
            }
            
            // Get messages from the last 6 hours (to catch recent replies)
            const sixHoursAgo = new Date();
            sixHoursAgo.setHours(sixHoursAgo.getHours() - 6);
            const filterDate = sixHoursAgo.toISOString();
            
            // Query inbox for received messages in the last 6 hours
            const messages = await graphClient
                .api('/me/messages')
                .filter(`receivedDateTime ge ${filterDate} and isDraft eq false`)
                .select('id,subject,from,receivedDateTime,isRead,conversationId,parentFolderId')
                .top(50)
                .get();
            
            if (messages.value.length === 0) {
                console.log(`📭 No recent messages found in session: ${sessionId}`);
                return 0;
            }
            
            console.log(`📧 Found ${messages.value.length} recent messages to check for replies`);
            
            // Download master file to get list of sent emails
            const masterWorkbook = await this.downloadMasterFile(graphClient);
            if (!masterWorkbook) {
                console.log(`📋 No master file found for session: ${sessionId}`);
                return 0;
            }
            
            // Get leads from master file
            const sheetInfo = this.excelProcessor.findLeadsSheet(masterWorkbook);
            if (!sheetInfo) {
                console.log(`📋 No valid lead data sheet found in session: ${sessionId}`);
                return 0;
            }
            
            const leadsData = require('xlsx').utils.sheet_to_json(sheetInfo.sheet);
            const sentEmails = leadsData.filter(lead => 
                lead.Last_Email_Date && !lead.Reply_Date
            ).map(lead => lead.Email);
            
            console.log(`📧 Checking ${sentEmails.length} sent emails for replies...`);
            
            let repliesFound = 0;
            
            // Check each message to see if it's a reply to our sent emails
            for (const message of messages.value) {
                try {
                    const fromEmail = message.from?.emailAddress?.address?.toLowerCase();
                    const subject = message.subject || '';
                    
                    if (!fromEmail) continue;
                    
                    // Check if this email is from someone we sent emails to
                    const isFromSentEmail = sentEmails.some(sentEmail => 
                        sentEmail && sentEmail.toLowerCase() === fromEmail
                    );
                    
                    // Additional checks to confirm it's likely a reply:
                    // 1. It's from someone we emailed
                    // 2. Subject might contain "Re:" or similar reply indicators
                    const subjectIndicatesReply = subject.toLowerCase().includes('re:') || 
                                                 subject.toLowerCase().includes('reply') ||
                                                 subject.toLowerCase().includes('response');
                    
                    if (isFromSentEmail) {
                        console.log(`💬 Potential reply detected from: ${fromEmail}`);
                        console.log(`📧 Subject: "${subject}"`);
                        console.log(`🔍 Subject indicates reply: ${subjectIndicatesReply}`);
                        
                        // For now, consider any email from someone we contacted as a potential reply
                        // This is a conservative approach to catch replies
                        
                        // Update Excel with reply status using Graph API direct updates
                        await this.updateReplyStatus(graphClient, fromEmail, message.receivedDateTime);
                        repliesFound++;
                    }
                    
                } catch (messageError) {
                    console.error(`❌ Error processing message:`, messageError.message);
                }
            }
            
            console.log(`✅ Session ${sessionId}: Found ${repliesFound} new replies`);
            return repliesFound;
            
        } catch (error) {
            console.error(`❌ Session reply check error for ${sessionId}:`, error);
            return 0;
        }
    }

    /**
     * Update reply status in Excel using Graph API
     */
    async updateReplyStatus(graphClient, email, receivedDateTime) {
        try {
            const updates = {
                Status: 'Replied',
                Reply_Date: new Date(receivedDateTime).toISOString().split('T')[0],
                'Last Updated': new Date().toISOString()
            };
            
            // Use direct Graph API update method
            const success = await this.updateExcelViaGraphAPI(graphClient, email, updates);
            
            if (success) {
                console.log(`✅ Reply status updated for: ${email}`);
            } else {
                console.log(`❌ Failed to update reply status for: ${email}`);
            }
            
        } catch (error) {
            console.error(`❌ Reply status update error for ${email}:`, error);
        }
    }

    /**
     * Direct Graph API Excel update - find email and update cells
     */
    async updateExcelViaGraphAPI(graphClient, email, updates) {
        try {
            const masterFileName = 'LGA-Master-Email-List.xlsx';
            const masterFolderPath = '/LGA-Email-Automation';
            
            console.log(`🔍 Graph API: Searching for ${email} in Excel file...`);
            
            // Get the Excel file ID
            const files = await graphClient
                .api(`/me/drive/root:${masterFolderPath}:/children`)
                .filter(`name eq '${masterFileName}'`)
                .get();

            if (files.value.length === 0) {
                console.log(`❌ Master file not found: ${masterFileName}`);
                return false;
            }

            const fileId = files.value[0].id;
            
            // Get worksheet info
            const worksheets = await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets`)
                .get();
                
            if (worksheets.value.length === 0) {
                console.log(`❌ No worksheets found in Excel file`);
                return false;
            }
            
            const worksheetName = worksheets.value[0].name;
            console.log(`📊 Using worksheet: ${worksheetName}`);
            
            // Get all data from the worksheet to find the email
            const usedRange = await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/usedRange`)
                .get();
            
            if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
                console.log(`❌ No data found in worksheet`);
                return false;
            }
            
            const headers = usedRange.values[0];
            const rows = usedRange.values.slice(1); // Skip header row
            
            console.log(`🔍 Found ${rows.length} data rows, searching for email: ${email}`);
            
            // Find email column index
            const emailColumnIndex = headers.findIndex(header => 
                header && typeof header === 'string' && 
                header.toLowerCase().includes('email') && 
                !header.toLowerCase().includes('date') &&
                !header.toLowerCase().includes('count')
            );
            
            if (emailColumnIndex === -1) {
                console.log(`❌ Email column not found in headers`);
                return false;
            }
            
            console.log(`📧 Email column found at index: ${emailColumnIndex} (${headers[emailColumnIndex]})`);
            
            // Find the row with matching email
            let targetRowIndex = -1;
            for (let i = 0; i < rows.length; i++) {
                const rowEmail = rows[i][emailColumnIndex];
                if (rowEmail && typeof rowEmail === 'string' && rowEmail.toLowerCase().trim() === email.toLowerCase().trim()) {
                    targetRowIndex = i;
                    console.log(`✅ Found matching email in row ${i + 2} (Excel row, including header)`);
                    break;
                }
            }
            
            if (targetRowIndex === -1) {
                console.log(`❌ Email ${email} not found in Excel file`);
                return false;
            }
            
            // Excel row number (1-based, including header)
            const excelRowNumber = targetRowIndex + 2;
            
            // Find column indices for fields we want to update
            const fieldColumnMap = {};
            for (const field of Object.keys(updates)) {
                const columnIndex = headers.findIndex(header => 
                    header && typeof header === 'string' && 
                    (header === field || header.replace(/[_\s]/g, '').toLowerCase() === field.replace(/[_\s]/g, '').toLowerCase())
                );
                
                if (columnIndex !== -1) {
                    // Convert column index to Excel column letter
                    const columnLetter = this.getExcelColumnLetter(columnIndex);
                    fieldColumnMap[field] = { index: columnIndex, letter: columnLetter };
                    console.log(`📍 Field '${field}' found at column ${columnIndex} (${columnLetter}): ${headers[columnIndex]}`);
                } else {
                    console.log(`⚠️ Field '${field}' not found in headers`);
                }
            }
            
            // Update each field directly via Graph API
            let updatedCount = 0;
            for (const [field, value] of Object.entries(updates)) {
                if (fieldColumnMap[field]) {
                    const columnLetter = fieldColumnMap[field].letter;
                    const cellAddress = `${columnLetter}${excelRowNumber}`;
                    
                    try {
                        console.log(`🔄 Updating cell ${cellAddress} with '${value}'`);
                        
                        await graphClient
                            .api(`/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/range(address='${cellAddress}')`)
                            .patch({
                                values: [[value]]
                            });
                        
                        console.log(`✅ Updated ${field} in cell ${cellAddress}`);
                        updatedCount++;
                        
                    } catch (cellUpdateError) {
                        console.error(`❌ Failed to update cell ${cellAddress}:`, cellUpdateError.message);
                    }
                }
            }
            
            if (updatedCount > 0) {
                console.log(`🎉 Successfully updated ${updatedCount} fields for ${email} via Graph API!`);
                return true;
            } else {
                console.log(`❌ No fields were successfully updated for ${email}`);
                return false;
            }
            
        } catch (error) {
            console.error(`❌ Graph API Excel update failed:`, error);
            return false;
        }
    }

    /**
     * Helper function to convert column index to Excel column letter
     */
    getExcelColumnLetter(columnIndex) {
        let result = '';
        let index = columnIndex;
        
        while (index >= 0) {
            result = String.fromCharCode(65 + (index % 26)) + result;
            index = Math.floor(index / 26) - 1;
        }
        
        return result;
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