const express = require('express');
const multer = require('multer');
const { requireDelegatedAuth } = require('../middleware/delegatedGraphAuth');
const EmailContentProcessor = require('../utils/emailContentProcessor');
const { getExcelColumnLetter } = require('../utils/excelGraphAPI');
const router = express.Router();

// Configure multer for attachment uploads
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 25 * 1024 * 1024, // 25MB limit (Microsoft Graph attachment limit)
        files: 5 // Maximum 5 files per template
    },
    fileFilter: (req, file, cb) => {
        // Allow common business file types
        const allowedMimes = [
            'application/pdf',
            'application/msword',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-powerpoint',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'image/jpeg',
            'image/png',
            'image/gif',
            'text/plain',
            'text/csv'
        ];

        if (allowedMimes.includes(file.mimetype)) {
            cb(null, true);
        } else {
            cb(new Error(`File type ${file.mimetype} not allowed. Allowed types: PDF, Word, Excel, PowerPoint, Images, Text files.`), false);
        }
    }
});

// Initialize processors
const emailContentProcessor = new EmailContentProcessor();

/**
 * Email Template Management
 * Handles template CRUD operations within the master Excel file
 */

// Get all templates
router.get('/', requireDelegatedAuth, async (req, res) => {
    try {
        console.log('📝 Retrieving email templates...');

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Get templates using Graph API
        const templates = await getTemplatesViaGraphAPI(graphClient);
        
        if (!templates) {
            return res.json({
                success: true,
                templates: [],
                message: 'No master file found'
            });
        }

        res.json({
            success: true,
            templates: templates,
            total: templates.length
        });

    } catch (error) {
        console.error('❌ Templates retrieval error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve templates',
            error: error.message
        });
    }
});

// Get specific template by ID
router.get('/:templateId', requireDelegatedAuth, async (req, res) => {
    try {
        const { templateId } = req.params;
        console.log(`📝 Retrieving template: ${templateId}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Get templates using Graph API
        const templates = await getTemplatesViaGraphAPI(graphClient);
        
        if (!templates) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Find specific template
        const template = templates.find(t => t.Template_ID === templateId);

        if (!template) {
            return res.status(404).json({
                success: false,
                message: 'Template not found'
            });
        }

        res.json({
            success: true,
            template: template
        });

    } catch (error) {
        console.error('❌ Template retrieval error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve template',
            error: error.message
        });
    }
});

// Create new template with optional file attachments
router.post('/', requireDelegatedAuth, upload.array('attachments', 5), async (req, res) => {
    try {
        const templateData = req.body;
        console.log('📝 Creating new template:', templateData.Template_Name);

        // Validate required fields
        const requiredFields = ['Template_Name', 'Template_Type', 'Subject', 'Body'];
        const missingFields = requiredFields.filter(field => !templateData[field]);
        
        if (missingFields.length > 0) {
            return res.status(400).json({
                success: false,
                message: 'Missing required fields',
                missingFields: missingFields
            });
        }

        // Process uploaded attachments
        let attachments = [];
        if (req.files && req.files.length > 0) {
            console.log(`📎 Processing ${req.files.length} attachments for template`);

            // Validate each attachment
            for (const file of req.files) {
                const validation = emailContentProcessor.validateAttachment({
                    name: file.originalname,
                    size: file.size,
                    contentType: file.mimetype
                });

                if (!validation.isValid) {
                    return res.status(400).json({
                        success: false,
                        message: `Attachment validation failed: ${validation.errors.join(', ')}`,
                        attachment: file.originalname
                    });
                }

                // Convert file to attachment format for storage
                attachments.push({
                    name: file.originalname,
                    contentType: file.mimetype,
                    size: file.size,
                    contentBytes: file.buffer.toString('base64'),
                    uploadDate: new Date().toISOString()
                });
            }

            console.log(`✅ ${attachments.length} attachments processed successfully`);
        }

        // Add attachments to template data
        if (attachments.length > 0) {
            templateData.attachments = attachments;
        }

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Add template using Graph API
        const templateId = await addTemplateViaGraphAPI(graphClient, templateData);
        
        if (!templateId) {
            throw new Error('Failed to create template');
        }

        console.log(`✅ Template created: ${templateId}`);

        res.json({
            success: true,
            message: 'Template created successfully',
            templateId: templateId,
            template: {
                Template_ID: templateId,
                ...templateData,
                attachmentCount: attachments.length,
                attachments: attachments.map(att => ({
                    name: att.name,
                    size: att.size,
                    contentType: att.contentType
                })) // Don't include base64 content in response
            }
        });

    } catch (error) {
        console.error('❌ Template creation error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to create template',
            error: error.message
        });
    }
});

// Update existing template
router.put('/:templateId', requireDelegatedAuth, async (req, res) => {
    try {
        const { templateId } = req.params;
        const updates = req.body;
        console.log(`📝 Updating template: ${templateId}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Update template using Graph API
        const updateSuccess = await updateTemplateViaGraphAPI(graphClient, templateId, updates);
        
        if (!updateSuccess) {
            return res.status(404).json({
                success: false,
                message: 'Template not found or update failed'
            });
        }

        console.log(`✅ Template updated: ${templateId}`);

        res.json({
            success: true,
            message: 'Template updated successfully',
            templateId: templateId,
            updatedFields: Object.keys(updates)
        });

    } catch (error) {
        console.error('❌ Template update error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to update template',
            error: error.message
        });
    }
});

// Delete template
router.delete('/:templateId', requireDelegatedAuth, async (req, res) => {
    try {
        const { templateId } = req.params;
        console.log(`📝 Deleting template: ${templateId}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Delete template using Graph API
        const deleteSuccess = await deleteTemplateViaGraphAPI(graphClient, templateId);
        
        if (!deleteSuccess) {
            return res.status(404).json({
                success: false,
                message: 'Template not found or delete failed'
            });
        }

        console.log(`✅ Template deleted: ${templateId}`);

        res.json({
            success: true,
            message: 'Template deleted successfully',
            templateId: templateId
        });

    } catch (error) {
        console.error('❌ Template deletion error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to delete template',
            error: error.message
        });
    }
});

// Preview template with sample data
router.post('/:templateId/preview', requireDelegatedAuth, async (req, res) => {
    try {
        const { templateId } = req.params;
        const { sampleLead } = req.body;
        console.log(`👁️ Previewing template: ${templateId}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Get template using Graph API
        const templates = await getTemplatesViaGraphAPI(graphClient);
        
        if (!templates) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Find specific template
        const template = templates.find(t => t.Template_ID === templateId);

        if (!template) {
            return res.status(404).json({
                success: false,
                message: 'Template not found'
            });
        }

        // Generate preview
        const preview = emailContentProcessor.previewEmailContent(template, sampleLead);

        res.json({
            success: true,
            template: template,
            preview: preview,
            sampleLead: sampleLead || 'Default sample data used'
        });

    } catch (error) {
        console.error('❌ Template preview error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to preview template',
            error: error.message
        });
    }
});

// Toggle template active status
router.patch('/:templateId/toggle', requireDelegatedAuth, async (req, res) => {
    try {
        const { templateId } = req.params;
        console.log(`🔄 Toggling template status: ${templateId}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Toggle template status using Graph API
        const toggleResult = await toggleTemplateStatusViaGraphAPI(graphClient, templateId);
        
        if (!toggleResult.success) {
            return res.status(404).json({
                success: false,
                message: 'Template not found or toggle failed'
            });
        }
        
        const newStatus = toggleResult.newStatus;

        console.log(`✅ Template status toggled: ${templateId} -> ${newStatus}`);

        res.json({
            success: true,
            message: `Template ${newStatus === 'Yes' ? 'activated' : 'deactivated'} successfully`,
            templateId: templateId,
            newStatus: newStatus
        });

    } catch (error) {
        console.error('❌ Template toggle error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to toggle template status',
            error: error.message
        });
    }
});

// Get templates by type
router.get('/type/:templateType', requireDelegatedAuth, async (req, res) => {
    try {
        const { templateType } = req.params;
        console.log(`📝 Retrieving templates of type: ${templateType}`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Get templates using Graph API
        const allTemplates = await getTemplatesViaGraphAPI(graphClient);
        
        if (!allTemplates) {
            return res.json({
                success: true,
                templates: [],
                message: 'No master file found'
            });
        }

        // Filter templates by type
        const filteredTemplates = allTemplates.filter(template => 
            template.Template_Type === templateType
        );

        res.json({
            success: true,
            templates: filteredTemplates,
            total: filteredTemplates.length,
            templateType: templateType
        });

    } catch (error) {
        console.error('❌ Templates by type retrieval error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to retrieve templates by type',
            error: error.message
        });
    }
});

// Validate template content
router.post('/validate', requireDelegatedAuth, async (req, res) => {
    try {
        const { templateData } = req.body;
        console.log('✅ Validating template content...');

        if (!templateData || !templateData.Subject || !templateData.Body) {
            return res.status(400).json({
                success: false,
                message: 'Template data with Subject and Body is required'
            });
        }

        // Create a mock email content object for validation
        const emailContent = {
            subject: templateData.Subject,
            body: templateData.Body
        };

        // Validate using email content processor
        const validation = emailContentProcessor.validateEmailContent(emailContent);

        // Extract variables used in template
        const variables = emailContentProcessor.extractVariables(
            templateData.Subject + ' ' + templateData.Body
        );

        // Get content statistics
        const stats = emailContentProcessor.getContentStats(emailContent);

        res.json({
            success: true,
            validation: validation,
            variables: variables,
            stats: stats,
            recommendations: generateTemplateRecommendations(templateData, validation, stats)
        });

    } catch (error) {
        console.error('❌ Template validation error:', error);
        res.status(500).json({
            success: false,
            message: 'Failed to validate template',
            error: error.message
        });
    }
});

// Helper function to generate template recommendations
function generateTemplateRecommendations(templateData, validation, stats) {
    const recommendations = [];

    // Subject line recommendations
    if (stats.subjectLength > 60) {
        recommendations.push({
            type: 'warning',
            field: 'subject',
            message: 'Subject line is longer than 60 characters. Consider shortening for better mobile display.'
        });
    }

    if (stats.subjectLength < 20) {
        recommendations.push({
            type: 'info',
            field: 'subject',
            message: 'Subject line is quite short. Consider adding more descriptive content.'
        });
    }

    // Body recommendations
    if (stats.wordCount < 50) {
        recommendations.push({
            type: 'info',
            field: 'body',
            message: 'Email body is quite short. Consider adding more value proposition.'
        });
    }

    if (stats.wordCount > 300) {
        recommendations.push({
            type: 'warning',
            field: 'body',
            message: 'Email body is quite long. Consider breaking it into shorter paragraphs.'
        });
    }

    // Variable recommendations
    if (stats.variableCount === 0) {
        recommendations.push({
            type: 'warning',
            field: 'personalization',
            message: 'No personalization variables found. Consider adding {Name} or {Company} for better engagement.'
        });
    }

    return recommendations;
}

// Graph API template management functions

// Get templates using Graph API
async function getTemplatesViaGraphAPI(graphClient) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Get Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            console.log('📋 Master file not found');
            return null;
        }

        const fileId = files.value[0].id;
        
        // Get Templates worksheet data
        const usedRange = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets('Templates')/usedRange`)
            .get();
        
        if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
            return [];
        }
        
        // Convert to template objects
        const headers = usedRange.values[0];
        const rows = usedRange.values.slice(1);
        
        return rows.map(row => {
            const template = {};
            headers.forEach((header, index) => {
                template[header] = row[index] || '';
            });

            // Parse attachments JSON if present
            if (template.Attachments && template.Attachments.trim()) {
                try {
                    template.attachments = JSON.parse(template.Attachments);
                    // Don't include full base64 content in list responses for performance
                    template.attachmentSummary = template.attachments.map(att => ({
                        name: att.name,
                        size: att.size,
                        contentType: att.contentType,
                        uploadDate: att.uploadDate
                    }));
                } catch (parseError) {
                    console.warn(`⚠️ Failed to parse attachments for template ${template.Template_ID}:`, parseError.message);
                    template.attachments = [];
                    template.attachmentSummary = [];
                }
            } else {
                template.attachments = [];
                template.attachmentSummary = [];
            }

            // Ensure attachment count is a number
            template.Attachment_Count = parseInt(template.Attachment_Count) || 0;

            return template;
        }).filter(template => template.Template_ID);
        
    } catch (error) {
        console.error('❌ Get templates via Graph API error:', error);
        return null;
    }
}

// Add template using Graph API
async function addTemplateViaGraphAPI(graphClient, templateData) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        const worksheetName = 'Templates';
        const tableName = 'TemplatesTable';
        
        // Get Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            throw new Error('Master file not found');
        }

        const fileId = files.value[0].id;
        
        // Generate template ID
        const templateId = `template_${Date.now()}`;
        
        // Prepare template row data
        const attachmentInfo = templateData.attachments ? JSON.stringify(templateData.attachments) : '';
        const attachmentCount = templateData.attachments ? templateData.attachments.length : 0;

        const templateRow = [
            templateId,
            templateData.Template_Name || '',
            templateData.Template_Type || '',
            templateData.Subject || '',
            templateData.Body || '',
            templateData.Active || 'Yes',
            attachmentInfo,
            attachmentCount
        ];
        
        // Discover what table to use
        let actualTableName = tableName;
        try {
            const existingTables = await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/tables`)
                .get();
            
            if (existingTables.value.length > 0) {
                const targetTable = existingTables.value.find(t => t.name === tableName);
                if (!targetTable) {
                    actualTableName = existingTables.value[0].name;
                    console.log(`🔄 Using existing table '${actualTableName}' for templates`);
                }
            }
        } catch (discoverError) {
            // Will use default table name
        }

        try {
            // Try to add row to existing table first
            await graphClient
                .api(`/me/drive/items/${fileId}/workbook/tables/${actualTableName}/rows/add`)
                .post({
                    values: [templateRow]
                });
                
            console.log(`✅ Template added to existing table: ${templateId}`);
            
        } catch (tableError) {
            try {
                // Create table if it doesn't exist, or get existing table name
                const returnedTableName = await createTemplatesTable(graphClient, fileId, worksheetName, tableName);
                const tableNameToUse = returnedTableName || tableName;
                
                // Wait a moment for table to be ready
                await new Promise(resolve => setTimeout(resolve, 1000));
                
                // Now add the row to the table
                await graphClient
                    .api(`/me/drive/items/${fileId}/workbook/tables/${tableNameToUse}/rows/add`)
                    .post({
                        values: [templateRow]
                    });
                    
                console.log(`✅ Template added: ${templateId}`);
                
            } catch (createError) {
                console.error('❌ Error creating table and adding template:', createError);
                throw createError;
            }
        }
        
        return templateId;
        
    } catch (error) {
        console.error('❌ Add template via Graph API error:', error);
        return null;
    }
}

// Update template using Graph API
async function updateTemplateViaGraphAPI(graphClient, templateId, updates) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Get Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            return false;
        }

        const fileId = files.value[0].id;
        
        // Get Templates worksheet data to find the row
        const usedRange = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets('Templates')/usedRange`)
            .get();
        
        if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
            return false;
        }
        
        const headers = usedRange.values[0];
        const rows = usedRange.values.slice(1);
        
        // Find template row
        const templateIdIndex = headers.findIndex(h => h === 'Template_ID');
        let targetRowIndex = -1;
        
        for (let i = 0; i < rows.length; i++) {
            if (rows[i][templateIdIndex] === templateId) {
                targetRowIndex = i + 2; // +2 for 1-based and header row
                break;
            }
        }
        
        if (targetRowIndex === -1) {
            return false;
        }
        
        // Update each field
        for (const [field, value] of Object.entries(updates)) {
            const colIndex = headers.findIndex(h => h === field);
            if (colIndex !== -1) {
                const cellAddress = `${getExcelColumnLetter(colIndex)}${targetRowIndex}`;
                
                await graphClient
                    .api(`/me/drive/items/${fileId}/workbook/worksheets('Templates')/range(address='${cellAddress}')`)
                    .patch({
                        values: [[value]]
                    });
            }
        }
        
        return true;
        
    } catch (error) {
        console.error('❌ Update template via Graph API error:', error);
        return false;
    }
}

// Delete template using Graph API
async function deleteTemplateViaGraphAPI(graphClient, templateId) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        // Get Excel file ID
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            return false;
        }

        const fileId = files.value[0].id;
        
        // Get Templates worksheet data to find the row
        const usedRange = await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets('Templates')/usedRange`)
            .get();
        
        if (!usedRange || !usedRange.values || usedRange.values.length <= 1) {
            return false;
        }
        
        const headers = usedRange.values[0];
        const rows = usedRange.values.slice(1);
        
        // Find template row
        const templateIdIndex = headers.findIndex(h => h === 'Template_ID');
        let targetRowIndex = -1;
        
        for (let i = 0; i < rows.length; i++) {
            if (rows[i][templateIdIndex] === templateId) {
                targetRowIndex = i + 1; // +1 for 0-based table rows (excluding header)
                break;
            }
        }
        
        if (targetRowIndex === -1) {
            return false;
        }
        
        // Delete row from table
        await graphClient
            .api(`/me/drive/items/${fileId}/workbook/tables/TemplatesTable/rows/itemAt(index=${targetRowIndex})`)
            .delete();
        
        return true;
        
    } catch (error) {
        console.error('❌ Delete template via Graph API error:', error);
        return false;
    }
}

// Toggle template status using Graph API
async function toggleTemplateStatusViaGraphAPI(graphClient, templateId) {
    try {
        // First get current status
        const templates = await getTemplatesViaGraphAPI(graphClient);
        if (!templates) {
            return { success: false };
        }
        
        const template = templates.find(t => t.Template_ID === templateId);
        if (!template) {
            return { success: false };
        }
        
        // Toggle status
        const newStatus = template.Active === 'Yes' ? 'No' : 'Yes';
        
        // Update using the update function
        const updateSuccess = await updateTemplateViaGraphAPI(graphClient, templateId, { Active: newStatus });
        
        return {
            success: updateSuccess,
            newStatus: newStatus
        };
        
    } catch (error) {
        console.error('❌ Toggle template status via Graph API error:', error);
        return { success: false };
    }
}

// Create Templates table if it doesn't exist
async function createTemplatesTable(graphClient, fileId, worksheetName, tableName) {
    try {
        console.log(`🗂️ Creating Templates table '${tableName}'...`);
        
        // Check if any tables already exist in the worksheet
        try {
            const existingTables = await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/tables`)
                .get();
            
            // If there's already a table with our name, just return
            const existingTargetTable = existingTables.value.find(t => t.name === tableName);
            if (existingTargetTable) {
                return;
            }
            
            // Use any existing table to avoid conflicts
            if (existingTables.value.length > 0) {
                const firstTable = existingTables.value[0];
                console.log(`🔄 Using existing table '${firstTable.name}' for templates`);
                return firstTable.name;
            }
        } catch (error) {
            // No existing tables found, will create new one
        }
        
        // Define template headers matching the Excel columns
        const headers = ['Template_ID', 'Template_Name', 'Template_Type', 'Subject', 'Body', 'Active', 'Attachments', 'Attachment_Count'];
        
        // Check if worksheet exists, create it if not
        try {
            await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')`)
                .get();
        } catch (worksheetError) {
            console.log(`📝 Creating worksheet '${worksheetName}'...`);
            await graphClient
                .api(`/me/drive/items/${fileId}/workbook/worksheets/add`)
                .post({
                    name: worksheetName
                });
        }
        
        // Write headers to first row
        await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/range(address='A1:${getExcelColumnLetter(headers.length - 1)}1')`)
            .patch({
                values: [headers]
            });
        
        // Create table from header row
        const tableRange = `A1:${getExcelColumnLetter(headers.length - 1)}1`;
        
        const tableRequest = {
            address: tableRange,
            hasHeaders: true,
            name: tableName
        };
        
        await graphClient
            .api(`/me/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/tables/add`)
            .post(tableRequest);
        
        console.log(`✅ Created Templates table '${tableName}'`);
        
    } catch (error) {
        console.error(`❌ Error creating Templates table:`, error.message);
        throw error;
    }
}



module.exports = router;