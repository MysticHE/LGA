const XLSX = require('xlsx');

/**
 * Excel Processing Utilities for Email Automation
 * Handles master Excel file operations for lead management
 */

class ExcelProcessor {
    constructor() {
        this.masterFileStructure = {
            // Existing scrape data columns
            'Name': 'text',
            'Title': 'text', 
            'Company Name': 'text',
            'Company Website': 'text',
            'Size': 'text',
            'Email': 'text', // Primary key
            'Email Verified': 'text',
            'LinkedIn URL': 'text',
            'Industry': 'text',
            'Location': 'text',
            'Last Updated': 'date',
            
            // Email automation columns
            'AI_Generated_Email': 'text',
            'Status': 'text', // New|Sent|Read|Replied|Bounced|Unsubscribed
            'Campaign_Stage': 'text', // First_Contact|Follow_Up_1|Follow_Up_2|Completed
            'Email_Choice': 'text', // AI_Generated|Email_Template_1|Email_Template_2
            'Template_Used': 'text',
            'Email_Content_Sent': 'text',
            'Last_Email_Date': 'date',
            'Next_Email_Date': 'date',
            'Follow_Up_Days': 'number',
            'Email_Count': 'number',
            'Max_Emails': 'number',
            'Auto_Send_Enabled': 'text', // Yes|No
            'Read_Date': 'date',
            'Reply_Date': 'date',
            'Email Sent': 'text', // Yes|No (legacy compatibility)
            'Email Status': 'text', // Status text (legacy compatibility)
            'Sent Date': 'date' // Send timestamp (legacy compatibility)
        };
    }

    /**
     * Create a new master Excel file with proper structure
     */
    createMasterFile(leads = []) {
        const wb = XLSX.utils.book_new();
        
        // Sheet 1: Leads (Main data)
        const leadsData = leads.length > 0 ? this.normalizeLeadsData(leads) : [this.getEmptyLeadRow()];
        const leadsSheet = XLSX.utils.json_to_sheet(leadsData);
        
        // Set column widths for better readability
        leadsSheet['!cols'] = this.getColumnWidths();
        
        XLSX.utils.book_append_sheet(wb, leadsSheet, 'Leads');
        
        // Sheet 2: Templates
        const templatesData = [this.getEmptyTemplateRow()];
        const templatesSheet = XLSX.utils.json_to_sheet(templatesData);
        templatesSheet['!cols'] = [
            {width: 20}, // Template_ID
            {width: 30}, // Template_Name
            {width: 20}, // Template_Type
            {width: 50}, // Subject
            {width: 80}, // Body
            {width: 10}  // Active
        ];
        
        XLSX.utils.book_append_sheet(wb, templatesSheet, 'Templates');
        
        // Sheet 3: Campaign History
        const campaignData = [this.getEmptyCampaignRow()];
        const campaignSheet = XLSX.utils.json_to_sheet(campaignData);
        campaignSheet['!cols'] = [
            {width: 20}, // Campaign_ID
            {width: 30}, // Campaign_Name
            {width: 20}, // Start_Date
            {width: 15}, // Emails_Sent
            {width: 15}, // Emails_Read
            {width: 15}, // Replies
            {width: 20}  // Status
        ];
        
        XLSX.utils.book_append_sheet(wb, campaignSheet, 'Campaign_History');
        
        return wb;
    }

    /**
     * Parse uploaded Excel file and extract leads
     */
    parseUploadedFile(fileBuffer) {
        try {
            const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
            const sheetName = workbook.SheetNames[0]; // Use first sheet
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`📊 Parsed ${data.length} rows from uploaded file`);
            
            // Normalize and validate data
            const validLeads = data.filter(row => this.isValidLead(row));
            
            console.log(`✅ ${validLeads.length} valid leads found`);
            
            return validLeads;
        } catch (error) {
            console.error('❌ Excel parsing error:', error);
            throw new Error('Failed to parse Excel file: ' + error.message);
        }
    }

    /**
     * Merge uploaded leads with existing master data
     */
    mergeLeadsWithMaster(uploadedLeads, existingData = []) {
        const results = {
            newLeads: [],
            duplicates: [],
            totalProcessed: uploadedLeads.length
        };

        // Create a Set of existing emails for fast lookup
        const existingEmails = new Set(
            existingData.map(lead => (lead.Email || '').toLowerCase().trim())
        );

        uploadedLeads.forEach(lead => {
            const email = (lead.Email || lead.email || '').toLowerCase().trim();
            
            if (!email) {
                console.warn('⚠️ Skipping lead without email:', lead.Name || 'Unknown');
                return;
            }

            if (existingEmails.has(email)) {
                results.duplicates.push({
                    email: email,
                    name: lead.Name || lead.name || '',
                    reason: 'Email already exists'
                });
            } else {
                // Normalize and add default automation settings
                const normalizedLead = this.normalizeLeadData(lead);
                results.newLeads.push(normalizedLead);
                existingEmails.add(email); // Prevent duplicates within upload
            }
        });

        console.log(`🔄 Merge results: ${results.newLeads.length} new, ${results.duplicates.length} duplicates`);
        
        return results;
    }

    /**
     * Normalize lead data to match master file structure
     */
    normalizeLeadData(lead) {
        const normalized = {};
        
        // Map various input formats to standard columns
        normalized['Name'] = lead.Name || lead.name || lead['Full Name'] || '';
        normalized['Title'] = lead.Title || lead.title || lead['Job Title'] || '';
        normalized['Company Name'] = lead['Company Name'] || lead.organization_name || lead.company || lead.Company || '';
        normalized['Company Website'] = lead['Company Website'] || lead.organization_website_url || lead.website || '';
        normalized['Size'] = lead.Size || lead.estimated_num_employees || lead.size || '';
        normalized['Email'] = lead.Email || lead.email || '';
        normalized['Email Verified'] = lead['Email Verified'] || lead.email_verified || 'N';
        normalized['LinkedIn URL'] = lead['LinkedIn URL'] || lead.linkedin_url || lead.linkedin || '';
        normalized['Industry'] = lead.Industry || lead.industry || '';
        normalized['Location'] = lead.Location || lead.country || lead.location || '';
        normalized['Last Updated'] = new Date().toISOString();

        // Move AI-generated content from Notes to AI_Generated_Email
        normalized['AI_Generated_Email'] = lead.Notes || lead.notes || lead.AI_Generated_Email || '';
        
        // Set default automation settings
        normalized['Status'] = 'New';
        normalized['Campaign_Stage'] = 'First_Contact';
        normalized['Email_Choice'] = 'AI_Generated';
        normalized['Template_Used'] = '';
        normalized['Email_Content_Sent'] = '';
        normalized['Last_Email_Date'] = '';
        normalized['Next_Email_Date'] = this.calculateNextEmailDate(new Date(), 7);
        normalized['Follow_Up_Days'] = 7;
        normalized['Email_Count'] = 0;
        normalized['Max_Emails'] = 3;
        normalized['Auto_Send_Enabled'] = 'Yes';
        normalized['Read_Date'] = '';
        normalized['Reply_Date'] = '';
        
        // Legacy compatibility
        normalized['Email Sent'] = '';
        normalized['Email Status'] = 'Not Sent';
        normalized['Sent Date'] = '';

        return normalized;
    }

    /**
     * Normalize multiple leads data
     */
    normalizeLeadsData(leads) {
        return leads.map(lead => this.normalizeLeadData(lead));
    }

    /**
     * Update master Excel file with new leads
     */
    updateMasterFileWithLeads(existingWorkbook, newLeads) {
        try {
            const leadsSheet = existingWorkbook.Sheets['Leads'];
            const existingData = XLSX.utils.sheet_to_json(leadsSheet);
            
            // Combine existing and new data
            const combinedData = [...existingData, ...newLeads];
            
            // Create new sheet with combined data
            const newSheet = XLSX.utils.json_to_sheet(combinedData);
            newSheet['!cols'] = this.getColumnWidths();
            
            // Replace the leads sheet
            existingWorkbook.Sheets['Leads'] = newSheet;
            
            return existingWorkbook;
        } catch (error) {
            console.error('❌ Master file update error:', error);
            throw new Error('Failed to update master file: ' + error.message);
        }
    }

    /**
     * Update specific lead in master file
     */
    updateLeadInMaster(workbook, email, updates) {
        try {
            const leadsSheet = workbook.Sheets['Leads'];
            const data = XLSX.utils.sheet_to_json(leadsSheet);
            
            // Find and update the lead
            let updated = false;
            for (let i = 0; i < data.length; i++) {
                if ((data[i].Email || '').toLowerCase() === email.toLowerCase()) {
                    Object.assign(data[i], updates);
                    data[i]['Last Updated'] = new Date().toISOString();
                    updated = true;
                    break;
                }
            }
            
            if (!updated) {
                throw new Error(`Lead with email ${email} not found`);
            }
            
            // Recreate sheet with updated data
            const newSheet = XLSX.utils.json_to_sheet(data);
            newSheet['!cols'] = this.getColumnWidths();
            workbook.Sheets['Leads'] = newSheet;
            
            return workbook;
        } catch (error) {
            console.error('❌ Lead update error:', error);
            throw error;
        }
    }

    /**
     * Get leads that are due for email today
     */
    getLeadsDueToday(workbook) {
        try {
            const leadsSheet = workbook.Sheets['Leads'];
            const data = XLSX.utils.sheet_to_json(leadsSheet);
            
            const today = new Date().toISOString().split('T')[0];
            
            return data.filter(lead => {
                if (lead.Auto_Send_Enabled !== 'Yes') return false;
                if (['Replied', 'Unsubscribed', 'Bounced'].includes(lead.Status)) return false;
                
                const nextEmailDate = lead.Next_Email_Date ? 
                    new Date(lead.Next_Email_Date).toISOString().split('T')[0] : null;
                
                return nextEmailDate && nextEmailDate <= today;
            });
        } catch (error) {
            console.error('❌ Due leads query error:', error);
            return [];
        }
    }

    /**
     * Get master file statistics
     */
    getMasterFileStats(workbook) {
        try {
            const leadsSheet = workbook.Sheets['Leads'];
            const data = XLSX.utils.sheet_to_json(leadsSheet);
            
            const stats = {
                totalLeads: data.length,
                dueToday: 0,
                emailsSent: 0,
                emailsRead: 0,
                repliesReceived: 0,
                statusBreakdown: {}
            };
            
            const today = new Date().toISOString().split('T')[0];
            
            data.forEach(lead => {
                // Count due today
                const nextEmailDate = lead.Next_Email_Date ? 
                    new Date(lead.Next_Email_Date).toISOString().split('T')[0] : null;
                
                if (nextEmailDate && nextEmailDate <= today && 
                    lead.Auto_Send_Enabled === 'Yes' && 
                    !['Replied', 'Unsubscribed', 'Bounced'].includes(lead.Status)) {
                    stats.dueToday++;
                }
                
                // Count by status
                const status = lead.Status || 'New';
                stats.statusBreakdown[status] = (stats.statusBreakdown[status] || 0) + 1;
                
                // Count actions
                if (lead.Status === 'Sent' || lead.Status === 'Read' || lead.Status === 'Replied') {
                    stats.emailsSent++;
                }
                if (lead.Status === 'Read' || lead.Status === 'Replied') {
                    stats.emailsRead++;
                }
                if (lead.Status === 'Replied') {
                    stats.repliesReceived++;
                }
            });
            
            return stats;
        } catch (error) {
            console.error('❌ Stats calculation error:', error);
            return {
                totalLeads: 0,
                dueToday: 0,
                emailsSent: 0,
                emailsRead: 0,
                repliesReceived: 0,
                statusBreakdown: {}
            };
        }
    }

    /**
     * Manage templates in master file
     */
    getTemplates(workbook) {
        try {
            const templatesSheet = workbook.Sheets['Templates'];
            if (!templatesSheet) return [];
            
            const data = XLSX.utils.sheet_to_json(templatesSheet);
            return data.filter(template => template.Template_ID && template.Template_ID !== '');
        } catch (error) {
            console.error('❌ Templates retrieval error:', error);
            return [];
        }
    }

    addTemplate(workbook, templateData) {
        try {
            const templatesSheet = workbook.Sheets['Templates'];
            const existingData = XLSX.utils.sheet_to_json(templatesSheet);
            
            // Generate template ID
            const templateId = `Template_${Date.now()}`;
            
            const newTemplate = {
                Template_ID: templateId,
                Template_Name: templateData.Template_Name,
                Template_Type: templateData.Template_Type,
                Subject: templateData.Subject,
                Body: templateData.Body,
                Active: templateData.Active || 'Yes'
            };
            
            const updatedData = [...existingData.filter(t => t.Template_ID), newTemplate];
            
            const newSheet = XLSX.utils.json_to_sheet(updatedData);
            newSheet['!cols'] = [
                {width: 20}, {width: 30}, {width: 20}, {width: 50}, {width: 80}, {width: 10}
            ];
            
            workbook.Sheets['Templates'] = newSheet;
            
            return templateId;
        } catch (error) {
            console.error('❌ Template addition error:', error);
            throw error;
        }
    }

    // Helper methods
    isValidLead(lead) {
        const email = lead.Email || lead.email || '';
        return email && email.includes('@');
    }

    calculateNextEmailDate(fromDate, days) {
        const nextDate = new Date(fromDate);
        nextDate.setDate(nextDate.getDate() + days);
        return nextDate.toISOString().split('T')[0];
    }

    getColumnWidths() {
        return [
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
            {width: 20}, // Last Updated
            {width: 60}, // AI_Generated_Email
            {width: 15}, // Status
            {width: 20}, // Campaign_Stage
            {width: 20}, // Email_Choice
            {width: 20}, // Template_Used
            {width: 40}, // Email_Content_Sent
            {width: 18}, // Last_Email_Date
            {width: 18}, // Next_Email_Date
            {width: 15}, // Follow_Up_Days
            {width: 12}, // Email_Count
            {width: 12}, // Max_Emails
            {width: 18}, // Auto_Send_Enabled
            {width: 18}, // Read_Date
            {width: 18}, // Reply_Date
            {width: 12}, // Email Sent
            {width: 15}, // Email Status
            {width: 18}  // Sent Date
        ];
    }

    getEmptyLeadRow() {
        const row = {};
        Object.keys(this.masterFileStructure).forEach(col => {
            row[col] = '';
        });
        return row;
    }

    getEmptyTemplateRow() {
        return {
            Template_ID: '',
            Template_Name: '',
            Template_Type: '',
            Subject: '',
            Body: '',
            Active: ''
        };
    }

    getEmptyCampaignRow() {
        return {
            Campaign_ID: '',
            Campaign_Name: '',
            Start_Date: '',
            Emails_Sent: '',
            Emails_Read: '',
            Replies: '',
            Status: ''
        };
    }

    /**
     * Convert workbook to buffer for saving
     */
    workbookToBuffer(workbook) {
        return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    }

    /**
     * Create workbook from buffer
     */
    bufferToWorkbook(buffer) {
        return XLSX.read(buffer, { type: 'buffer' });
    }
}

module.exports = ExcelProcessor;