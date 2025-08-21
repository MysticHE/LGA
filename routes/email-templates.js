const express = require('express');
const XLSX = require('xlsx');
const { requireDelegatedAuth } = require('../middleware/delegatedGraphAuth');
const ExcelProcessor = require('../utils/excelProcessor');
const EmailContentProcessor = require('../utils/emailContentProcessor');
const router = express.Router();

// Initialize processors
const excelProcessor = new ExcelProcessor();
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

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.json({
                success: true,
                templates: [],
                message: 'No master file found'
            });
        }

        // Get templates from master file
        const templates = excelProcessor.getTemplates(masterWorkbook);

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

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Get templates and find specific one
        const templates = excelProcessor.getTemplates(masterWorkbook);
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

// Create new template
router.post('/', requireDelegatedAuth, async (req, res) => {
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

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download master file
        let masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            // Create new master file if it doesn't exist
            masterWorkbook = excelProcessor.createMasterFile();
        }

        // Add template to master file
        const templateId = excelProcessor.addTemplate(masterWorkbook, templateData);

        // Save updated master file
        const masterBuffer = excelProcessor.workbookToBuffer(masterWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

        console.log(`✅ Template created: ${templateId}`);

        res.json({
            success: true,
            message: 'Template created successfully',
            templateId: templateId,
            template: {
                Template_ID: templateId,
                ...templateData
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

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Update template in master file
        const templatesSheet = masterWorkbook.Sheets['Templates'];
        const templatesData = XLSX.utils.sheet_to_json(templatesSheet);
        
        // Find and update template
        let templateFound = false;
        for (let i = 0; i < templatesData.length; i++) {
            if (templatesData[i].Template_ID === templateId) {
                Object.assign(templatesData[i], updates);
                templateFound = true;
                break;
            }
        }

        if (!templateFound) {
            return res.status(404).json({
                success: false,
                message: 'Template not found'
            });
        }

        // Recreate templates sheet
        const newTemplatesSheet = XLSX.utils.json_to_sheet(templatesData);
        newTemplatesSheet['!cols'] = [
            {width: 20}, {width: 30}, {width: 20}, {width: 50}, {width: 80}, {width: 10}
        ];
        masterWorkbook.Sheets['Templates'] = newTemplatesSheet;

        // Save updated master file
        const masterBuffer = excelProcessor.workbookToBuffer(masterWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

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

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Remove template from master file
        const templatesSheet = masterWorkbook.Sheets['Templates'];
        let templatesData = XLSX.utils.sheet_to_json(templatesSheet);
        
        const originalLength = templatesData.length;
        templatesData = templatesData.filter(template => template.Template_ID !== templateId);

        if (templatesData.length === originalLength) {
            return res.status(404).json({
                success: false,
                message: 'Template not found'
            });
        }

        // Recreate templates sheet
        const newTemplatesSheet = XLSX.utils.json_to_sheet(templatesData);
        newTemplatesSheet['!cols'] = [
            {width: 20}, {width: 30}, {width: 20}, {width: 50}, {width: 80}, {width: 10}
        ];
        masterWorkbook.Sheets['Templates'] = newTemplatesSheet;

        // Save updated master file
        const masterBuffer = excelProcessor.workbookToBuffer(masterWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

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

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Get template
        const templates = excelProcessor.getTemplates(masterWorkbook);
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

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.status(404).json({
                success: false,
                message: 'Master file not found'
            });
        }

        // Toggle template status
        const templatesSheet = masterWorkbook.Sheets['Templates'];
        const templatesData = XLSX.utils.sheet_to_json(templatesSheet);
        
        let templateFound = false;
        let newStatus = 'No';
        
        for (let i = 0; i < templatesData.length; i++) {
            if (templatesData[i].Template_ID === templateId) {
                newStatus = templatesData[i].Active === 'Yes' ? 'No' : 'Yes';
                templatesData[i].Active = newStatus;
                templateFound = true;
                break;
            }
        }

        if (!templateFound) {
            return res.status(404).json({
                success: false,
                message: 'Template not found'
            });
        }

        // Recreate templates sheet
        const newTemplatesSheet = XLSX.utils.json_to_sheet(templatesData);
        newTemplatesSheet['!cols'] = [
            {width: 20}, {width: 30}, {width: 20}, {width: 50}, {width: 80}, {width: 10}
        ];
        masterWorkbook.Sheets['Templates'] = newTemplatesSheet;

        // Save updated master file
        const masterBuffer = excelProcessor.workbookToBuffer(masterWorkbook);
        await uploadToOneDrive(graphClient, masterBuffer, 'LGA-Master-Email-List.xlsx', '/LGA-Email-Automation');

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

        // Download master file
        const masterWorkbook = await downloadMasterFile(graphClient);
        
        if (!masterWorkbook) {
            return res.json({
                success: true,
                templates: [],
                message: 'No master file found'
            });
        }

        // Get templates and filter by type
        const allTemplates = excelProcessor.getTemplates(masterWorkbook);
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

// Helper function to download master file
async function downloadMasterFile(graphClient) {
    try {
        const masterFileName = 'LGA-Master-Email-List.xlsx';
        const masterFolderPath = '/LGA-Email-Automation';
        
        const files = await graphClient
            .api(`/me/drive/root:${masterFolderPath}:/children`)
            .filter(`name eq '${masterFileName}'`)
            .get();

        if (files.value.length === 0) {
            console.log('📋 Master file not found');
            return null;
        }

        const fileContent = await graphClient
            .api(`/me/drive/items/${files.value[0].id}/content`)
            .get();

        return excelProcessor.bufferToWorkbook(fileContent);
    } catch (error) {
        console.error('❌ Master file download error:', error);
        return null;
    }
}

// Helper function to upload file to OneDrive
async function uploadToOneDrive(client, fileBuffer, filename, folderPath) {
    try {
        // Create folder if it doesn't exist
        try {
            await client.api(`/me/drive/root:${folderPath}`).get();
        } catch (error) {
            if (error.code === 'itemNotFound') {
                const folderName = folderPath.split('/').pop();
                const parentPath = folderPath.substring(0, folderPath.lastIndexOf('/')) || '/';
                
                await client.api(`/me/drive/root:${parentPath}:/children`).post({
                    name: folderName,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'rename'
                });
            }
        }

        const uploadUrl = `/me/drive/root:${folderPath}/${filename}:/content`;
        const result = await client.api(uploadUrl).put(fileBuffer);
        
        console.log(`📤 Uploaded file: ${filename} to ${folderPath}`);
        
        return {
            id: result.id,
            name: result.name,
            webUrl: result.webUrl,
            size: result.size
        };
    } catch (error) {
        console.error('❌ OneDrive upload error:', error);
        throw error;
    }
}

module.exports = router;