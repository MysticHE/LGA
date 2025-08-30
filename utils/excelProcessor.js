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
        let leadsData;
        if (leads.length > 0) {
            leadsData = this.normalizeLeadsData(leads);
        } else {
            // Create sheet with headers only (no empty row)
            leadsData = [];
        }
        
        const leadsSheet = XLSX.utils.json_to_sheet(leadsData, { 
            header: Object.keys(this.masterFileStructure) 
        });
        
        // Set column widths for better readability
        leadsSheet['!cols'] = this.getColumnWidths();
        
        XLSX.utils.book_append_sheet(wb, leadsSheet, 'Leads');
        
        // Sheet 2: Templates (headers only, no empty rows)
        const templatesSheet = XLSX.utils.json_to_sheet([], {
            header: ['Template_ID', 'Template_Name', 'Template_Type', 'Subject', 'Body', 'Active']
        });
        templatesSheet['!cols'] = [
            {width: 20}, // Template_ID
            {width: 30}, // Template_Name
            {width: 20}, // Template_Type
            {width: 50}, // Subject
            {width: 80}, // Body
            {width: 10}  // Active
        ];
        
        XLSX.utils.book_append_sheet(wb, templatesSheet, 'Templates');
        
        // Sheet 3: Campaign History (headers only, no empty rows)
        const campaignSheet = XLSX.utils.json_to_sheet([], {
            header: ['Campaign_ID', 'Campaign_Name', 'Start_Date', 'Emails_Sent', 'Emails_Read', 'Replies', 'Status']
        });
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
            
            // Try to intelligently find the leads sheet, fallback to first sheet
            const sheetInfo = this.findLeadsSheet(workbook);
            let worksheet;
            let sheetName;
            
            if (sheetInfo) {
                worksheet = sheetInfo.sheet;
                sheetName = sheetInfo.name;
                console.log(`📊 Using intelligently detected sheet: "${sheetName}"`);
            } else {
                // Fallback to first sheet if no intelligent match found
                sheetName = workbook.SheetNames[0];
                worksheet = workbook.Sheets[sheetName];
                console.log(`📊 Using first sheet as fallback: "${sheetName}"`);
            }
            
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

        console.log(`🔄 Merging uploaded and existing leads...`);
        console.log(`📊 Found ${existingData.length} existing leads for duplicate checking`);

        // Create a Set of existing emails for fast lookup with better normalization
        const existingEmails = new Set();
        existingData.forEach(lead => {
            const email = this.normalizeEmail(lead.Email || lead.email || '');
            if (email) {
                existingEmails.add(email);
            }
        });
        
        console.log(`📧 Created lookup set with ${existingEmails.size} unique existing emails`);

        uploadedLeads.forEach((lead, index) => {
            const email = this.normalizeEmail(lead.Email || lead.email || '');
            
            if (!email) {
                console.warn(`⚠️ Row ${index + 1}: Skipping lead without valid email:`, {
                    name: lead.Name || lead.name || 'Unknown',
                    originalEmail: lead.Email || lead.email || 'None'
                });
                return;
            }

            if (existingEmails.has(email)) {
                results.duplicates.push({
                    email: email,
                    name: lead.Name || lead.name || '',
                    reason: 'Email already exists in master list'
                });
            } else {
                // Clear progress logging every 10 leads or at key milestones
                if ((index + 1) % 10 === 0 || index + 1 === uploadedLeads.length) {
                    console.log(`✅ Processing new leads... (${index + 1}/${uploadedLeads.length})`);
                }
                // Normalize and add default automation settings
                const normalizedLead = this.normalizeLeadData(lead);
                results.newLeads.push(normalizedLead);
                existingEmails.add(email); // Prevent duplicates within current upload
            }
        });

        // Merge results summary
        console.log(`📊 Processing results: ${uploadedLeads.length} uploaded, ${results.duplicates.length} duplicates skipped, ${results.newLeads.length} new leads ready for upload`);
        
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
        // Ensure email is properly normalized
        normalized['Email'] = this.normalizeEmail(lead.Email || lead.email || '');
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
     * @param {object} existingWorkbook - The existing Excel workbook
     * @param {array} newLeads - Array of new lead objects to add
     * @param {array} existingData - Optional: Pre-extracted existing data (avoids re-extraction)
     */
    updateMasterFileWithLeads(existingWorkbook, newLeads, existingData = null) {
        try {
            const leadsSheet = existingWorkbook.Sheets['Leads'];
            let finalExistingData = [];
            
            // Use pre-extracted data if provided, otherwise extract from workbook
            if (existingData !== null) {
                console.log(`✅ USING PRE-EXTRACTED DATA: ${existingData.length} existing leads provided`);
                finalExistingData = existingData;
            } else if (leadsSheet && leadsSheet['!ref']) {
                console.log(`🔄 EXTRACTING FROM WORKBOOK: Re-extracting existing data from sheet`);
                // Extract existing data from the workbook
                let rawExistingData = XLSX.utils.sheet_to_json(leadsSheet);
                console.log(`📊 RAW EXISTING DATA: Found ${rawExistingData.length} rows in master file`);
                
                // DEBUG: Show first few rows to understand data structure
                if (rawExistingData.length > 0) {
                    console.log(`🔍 SAMPLE DATA STRUCTURE:`, Object.keys(rawExistingData[0]));
                    console.log(`🔍 FIRST ROW EMAIL FIELD:`, {
                        'Email': rawExistingData[0].Email,
                        'email': rawExistingData[0].email,
                        'Email_normalized': this.normalizeEmail(rawExistingData[0].Email || rawExistingData[0].email || '')
                    });
                }
                
                // More lenient filtering - check multiple email field variations and data indicators
                const originalCount = rawExistingData.length;
                finalExistingData = rawExistingData.filter(row => {
                    // Check for email field variations
                    const email = this.normalizeEmail(row.Email || row.email || '');
                    
                    // Check for other data indicators (Name, Company, Title) - if any exist, keep the row
                    const hasData = email || 
                                  (row.Name && row.Name.toString().trim().length > 0) ||
                                  (row.name && row.name.toString().trim().length > 0) ||
                                  (row['Company Name'] && row['Company Name'].toString().trim().length > 0) ||
                                  (row.Title && row.Title.toString().trim().length > 0);
                    
                    if (!hasData) {
                        console.log(`🗑️ FILTERED OUT EMPTY ROW:`, row);
                    }
                    
                    return hasData;
                });
                console.log(`📊 FILTERED EXISTING DATA: ${finalExistingData.length} valid rows after filtering (removed ${originalCount - finalExistingData.length} empty rows)`);
            } else {
                console.log(`⚠️ CRITICAL: No Leads sheet found or empty sheet reference`);
                console.log(`📊 Leads sheet exists: ${!!leadsSheet}`);
                console.log(`📊 Sheet reference: ${leadsSheet ? leadsSheet['!ref'] : 'null'}`);
            }
            
            console.log(`📊 Merging ${finalExistingData.length} existing + ${newLeads.length} new leads`);
            
            if (finalExistingData.length === 0) {
                console.log(`⚠️ WARNING: No existing data found - this will result in data replacement!`);
            } else {
                console.log(`✅ APPEND MODE: Will preserve ${finalExistingData.length} existing leads`);
            }
            
            // Combine existing and new data (APPEND mode)
            const combinedData = [...finalExistingData, ...newLeads];
            
            
            // Ensure all rows have the complete structure
            const normalizedData = combinedData.map(row => {
                const normalized = {};
                // Start with master file structure defaults
                Object.keys(this.masterFileStructure).forEach(key => {
                    normalized[key] = row[key] || '';
                });
                return normalized;
            });
            
            // Create new sheet with normalized combined data
            const newSheet = XLSX.utils.json_to_sheet(normalizedData, {
                header: Object.keys(this.masterFileStructure)
            });
            
            newSheet['!cols'] = this.getColumnWidths();
            
            // Replace the leads sheet with updated data
            existingWorkbook.Sheets['Leads'] = newSheet;
            
            // Simple verification
            const verificationData = XLSX.utils.sheet_to_json(newSheet);
            console.log(`✅ FINAL: ${verificationData.length} total leads in sheet`);
            
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
            // Use the intelligent sheet finder
            const sheetInfo = this.findLeadsSheet(workbook);
            
            if (!sheetInfo) {
                throw new Error(`No valid lead data sheet found in workbook. Available sheets: ${Object.keys(workbook.Sheets).join(', ')}`);
            }
            
            const leadsSheet = sheetInfo.sheet;
            const sheetName = sheetInfo.name;
            
            const data = XLSX.utils.sheet_to_json(leadsSheet);
            
            console.log(`🔍 Looking for email "${email}" in ${data.length} leads`);
            
            // Find and update the lead - try multiple email field variations
            let updated = false;
            const searchEmail = email.toLowerCase().trim();
            
            for (let i = 0; i < data.length; i++) {
                const lead = data[i];
                
                // Try multiple possible email field names
                const leadEmails = [
                    lead.Email,
                    lead.email, 
                    lead['Email Address'],
                    lead['email_address'],
                    lead.EmailAddress,
                    lead['Contact Email'],
                    lead['Primary Email']
                ].filter(Boolean).map(e => String(e).toLowerCase().trim());
                
                
                if (leadEmails.includes(searchEmail)) {
                    console.log(`✅ FOUND: Updating lead ${i} with email ${searchEmail}`);
                    Object.assign(data[i], updates);
                    data[i]['Last Updated'] = new Date().toISOString();
                    updated = true;
                    break;
                }
            }
            
            if (!updated) {
                // Enhanced error with actual data for debugging
                const availableEmails = data.map(lead => 
                    lead.Email || lead.email || lead['Email Address'] || 'NO_EMAIL'
                ).slice(0, 10);
                
                console.error(`❌ DETAILED DEBUG: Lead not found`);
                console.error(`   - Searching for: "${email}"`);
                console.error(`   - Total leads in file: ${data.length}`);
                console.error(`   - First 10 emails in file:`, availableEmails);
                console.error(`   - Available columns:`, Object.keys(data[0] || {}));
                
                throw new Error(`Lead with email ${email} not found. Total leads: ${data.length}, First emails: ${availableEmails.join(', ')}`);
            }
            
            // Recreate sheet with updated data using the detected sheet name
            console.log(`💾 Recreating sheet "${sheetName}" with updated data`);
            
            const newSheet = XLSX.utils.json_to_sheet(data);
            newSheet['!cols'] = this.getColumnWidths();
            workbook.Sheets[sheetName] = newSheet;
            
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
                newRecords: 0,
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
                
                // Count new records specifically
                if (status === 'New') {
                    stats.newRecords++;
                }
                
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
                newRecords: 0,
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

    /**
     * Helper method to intelligently find the leads sheet in a workbook
     */
    findLeadsSheet(workbook) {
        console.log(`📊 Available sheets in workbook:`, Object.keys(workbook.Sheets || {}));
        
        // First try the expected sheet names
        const expectedSheetNames = ['Leads', 'leads', 'LEADS'];
        for (const name of expectedSheetNames) {
            if (workbook.Sheets[name]) {
                console.log(`✅ Found expected sheet: "${name}"`);
                return { sheet: workbook.Sheets[name], name: name };
            }
        }
        
        // If not found, intelligently search for sheet containing lead data
        console.log(`⚠️ Expected sheet names not found. Searching for sheet with lead data...`);
        
        const sheetNames = Object.keys(workbook.Sheets);
        for (const name of sheetNames) {
            const sheet = workbook.Sheets[name];
            const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Get raw headers
            
            if (data.length > 0) {
                const headers = data[0] || [];
                const hasEmailColumn = headers.some(header => 
                    header && typeof header === 'string' && 
                    header.toLowerCase().includes('email')
                );
                const hasNameColumn = headers.some(header => 
                    header && typeof header === 'string' && 
                    header.toLowerCase().includes('name')
                );
                
                // This sheet likely contains lead data if it has email and name columns
                if (hasEmailColumn && hasNameColumn) {
                    console.log(`✅ Found lead data in sheet: "${name}" (has email and name columns)`);
                    return { sheet: sheet, name: name };
                }
            }
        }
        
        console.error(`❌ No valid lead data sheet found in workbook. Available sheets: ${Object.keys(workbook.Sheets).join(', ')}`);
        return null;
    }

    // Helper methods
    isValidLead(lead) {
        const email = this.normalizeEmail(lead.Email || lead.email || '');
        return email && email.includes('@') && this.isValidEmail(email);
    }
    
    /**
     * Normalize email for consistent comparison
     */
    normalizeEmail(email) {
        if (!email || typeof email !== 'string') {
            return '';
        }
        return email.toLowerCase().trim();
    }
    
    /**
     * Basic email validation
     */
    isValidEmail(email) {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return emailRegex.test(email);
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
     * DEBUG: Validate and inspect workbook contents  
     */
    debugWorkbook(workbook, description = 'Unknown') {
        
        const sheets = Object.keys(workbook.Sheets);
        
        sheets.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const range = sheet['!ref'];
            const data = XLSX.utils.sheet_to_json(sheet);
            
            console.log(`   - Range: ${range}`);
            console.log(`   - Row count: ${data.length}`);
            
        });
    }

    /**
     * Convert workbook to buffer for saving
     */
    workbookToBuffer(workbook) {
        
        // Use xlsx format for better OneDrive compatibility
        const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        
        return buffer;
    }

    /**
     * Create workbook from buffer
     */
    bufferToWorkbook(buffer) {
        
        const workbook = XLSX.read(buffer, { type: 'buffer' });
        
        return workbook;
    }
}

module.exports = ExcelProcessor;