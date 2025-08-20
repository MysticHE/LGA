const express = require('express');
const XLSX = require('xlsx');
const { requireDelegatedAuth, getDelegatedAuthProvider } = require('../middleware/delegatedGraphAuth');
const router = express.Router();

// Apply delegated Graph authentication middleware to protected routes
// Note: /test route is excluded from authentication requirement

/**
 * OneDrive Excel Integration
 * Upload and manage lead data in Microsoft 365 Excel files
 */

// Create Excel workbook in OneDrive
router.post('/onedrive/create-excel', requireDelegatedAuth, async (req, res) => {
    try {
        const { leads, filename, folderPath = '/LGA-Leads' } = req.body;

        if (!leads || !Array.isArray(leads) || leads.length === 0) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'Leads array is required and must not be empty'
            });
        }

        console.log(`📊 Creating Excel file for ${leads.length} leads in OneDrive...`);

        // Create Excel workbook
        const excelData = leads.map(lead => ({
            'Name': lead.name || '',
            'Title': lead.title || '',
            'Company Name': lead.organization_name || '',
            'Company Website': lead.organization_website_url || '',
            'Size': lead.estimated_num_employees || '',
            'Email': lead.email || '',
            'Email Verified': lead.email_verified || 'N',
            'LinkedIn URL': lead.linkedin_url || '',
            'Industry': lead.industry || '',
            'Location': lead.country || '',
            'Notes': lead.notes || '',
            'Conversion Status': lead.conversion_status || 'Pending',
            'Email Sent': '',
            'Email Status': 'Not Sent',
            'Sent Date': '',
            'Read Date': '',
            'Reply Date': '',
            'Last Updated': new Date().toISOString()
        }));

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(excelData);

        // Set column widths for better readability
        ws['!cols'] = [
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
            {width: 50}, // Notes
            {width: 18}, // Conversion Status
            {width: 12}, // Email Sent
            {width: 15}, // Email Status
            {width: 20}, // Sent Date
            {width: 20}, // Read Date
            {width: 20}, // Reply Date
            {width: 20}  // Last Updated
        ];

        XLSX.utils.book_append_sheet(wb, ws, 'Leads');

        // Convert to buffer
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        // Generate filename with timestamp if not provided
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:.]/g, '-');
        const finalFilename = filename || `singapore-leads-${timestamp}.xlsx`;

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Create folder if it doesn't exist
        await createOneDriveFolder(graphClient, folderPath);

        // Upload Excel file to OneDrive
        const uploadResult = await uploadToOneDrive(
            graphClient, 
            excelBuffer, 
            finalFilename, 
            folderPath
        );

        console.log(`✅ Excel file uploaded to OneDrive: ${finalFilename}`);

        res.json({
            success: true,
            filename: finalFilename,
            folderPath: folderPath,
            leadsCount: leads.length,
            oneDriveUrl: uploadResult.webUrl,
            fileId: uploadResult.id,
            metadata: {
                uploadedAt: new Date().toISOString(),
                size: excelBuffer.length,
                location: 'Microsoft OneDrive'
            }
        });

    } catch (error) {
        console.error('OneDrive Excel creation error:', error);
        res.status(500).json({
            error: 'OneDrive Excel Creation Error',
            message: 'Failed to create Excel file in OneDrive',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Update Excel file with email tracking data
router.post('/onedrive/update-excel-tracking', requireDelegatedAuth, async (req, res) => {
    try {
        const { fileId, leadEmail, trackingData } = req.body;

        if (!fileId || !leadEmail || !trackingData) {
            return res.status(400).json({
                error: 'Validation Error',
                message: 'File ID, lead email, and tracking data are required'
            });
        }

        console.log(`📊 Updating Excel tracking for ${leadEmail}...`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        // Download current Excel file
        const fileContent = await graphClient
            .api(`/me/drive/items/${fileId}/content`)
            .get();

        // Parse Excel content
        const workbook = XLSX.read(fileContent, { type: 'buffer' });
        const worksheet = workbook.Sheets['Leads'];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // Find and update the lead row
        let updated = false;
        for (let i = 0; i < data.length; i++) {
            if (data[i]['Email'] === leadEmail) {
                data[i]['Email Sent'] = trackingData.sent ? 'Yes' : 'No';
                data[i]['Email Status'] = trackingData.status || 'Sent';
                data[i]['Sent Date'] = trackingData.sentDate || '';
                data[i]['Read Date'] = trackingData.readDate || '';
                data[i]['Reply Date'] = trackingData.replyDate || '';
                data[i]['Last Updated'] = new Date().toISOString();
                updated = true;
                break;
            }
        }

        if (!updated) {
            return res.status(404).json({
                error: 'Lead Not Found',
                message: 'Lead email not found in Excel file'
            });
        }

        // Convert back to Excel
        const newWorksheet = XLSX.utils.json_to_sheet(data);
        // Preserve column widths
        newWorksheet['!cols'] = worksheet['!cols'];
        workbook.Sheets['Leads'] = newWorksheet;

        const updatedBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

        // Upload updated file back to OneDrive
        await graphClient
            .api(`/me/drive/items/${fileId}/content`)
            .put(updatedBuffer);

        console.log(`✅ Excel tracking updated for ${leadEmail}`);

        res.json({
            success: true,
            message: `Tracking data updated for ${leadEmail}`,
            updatedAt: new Date().toISOString()
        });

    } catch (error) {
        console.error('Excel tracking update error:', error);
        res.status(500).json({
            error: 'Excel Update Error',
            message: 'Failed to update Excel tracking data',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// List OneDrive files
router.get('/onedrive/files', requireDelegatedAuth, async (req, res) => {
    try {
        const { folderPath = '/LGA-Leads' } = req.query;

        console.log(`📂 Listing OneDrive files in ${folderPath}...`);

        // Get authenticated Graph client
        const graphClient = await req.delegatedAuth.getGraphClient(req.sessionId);

        const files = await graphClient
            .api(`/me/drive/root:${folderPath}:/children`)
            .filter("file ne null")
            .select('id,name,size,createdDateTime,lastModifiedDateTime,webUrl')
            .get();

        res.json({
            success: true,
            folderPath: folderPath,
            count: files.value.length,
            files: files.value.map(file => ({
                id: file.id,
                name: file.name,
                size: file.size,
                createdAt: file.createdDateTime,
                modifiedAt: file.lastModifiedDateTime,
                webUrl: file.webUrl
            }))
        });

    } catch (error) {
        console.error('OneDrive file listing error:', error);
        res.status(500).json({
            error: 'OneDrive Error',
            message: 'Failed to list OneDrive files',
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Helper function to create OneDrive folder
async function createOneDriveFolder(client, folderPath) {
    try {
        // Check if folder exists
        await client.api(`/me/drive/root:${folderPath}`).get();
        console.log(`📂 Folder ${folderPath} already exists`);
    } catch (error) {
        if (error.code === 'itemNotFound') {
            // Create folder
            const folderName = folderPath.split('/').pop();
            const parentPath = folderPath.substring(0, folderPath.lastIndexOf('/')) || '/';
            
            await client.api(`/me/drive/root:${parentPath}:/children`).post({
                name: folderName,
                folder: {},
                '@microsoft.graph.conflictBehavior': 'rename'
            });
            
            console.log(`📂 Created folder: ${folderPath}`);
        } else {
            throw error;
        }
    }
}

// Helper function to upload file to OneDrive
async function uploadToOneDrive(client, fileBuffer, filename, folderPath) {
    try {
        const uploadUrl = `/me/drive/root:${folderPath}/${filename}:/content`;
        
        const result = await client.api(uploadUrl).put(fileBuffer);
        
        return {
            id: result.id,
            name: result.name,
            webUrl: result.webUrl,
            size: result.size
        };
    } catch (error) {
        console.error('OneDrive upload error:', error);
        throw error;
    }
}

// Test Microsoft Graph connection - handles both authenticated and unauthenticated requests
router.get('/test', async (req, res) => {
    try {
        const sessionId = req.headers['x-session-id'] || req.query.sessionId;
        
        if (!sessionId) {
            return res.json({
                success: false,
                message: 'Authentication required for Microsoft Graph integration',
                authRequired: true,
                loginUrl: '/auth/login'
            });
        }

        const authProvider = getDelegatedAuthProvider();
        
        if (!authProvider.isAuthenticated(sessionId)) {
            return res.json({
                success: false,
                message: 'Session not authenticated',
                authRequired: true,
                loginUrl: '/auth/login'
            });
        }

        const testResult = await authProvider.testConnection(sessionId);
        
        if (testResult.success) {
            res.json({
                success: true,
                message: 'Microsoft Graph connection successful',
                user: testResult.user,
                email: testResult.email,
                accessType: 'Delegated',
                permissions: 'User.Read, Files.ReadWrite.All, Mail.Send, Mail.ReadWrite'
            });
        } else {
            res.status(401).json({
                success: false,
                message: 'Microsoft Graph connection failed',
                error: testResult.error,
                authRequired: true
            });
        }
    } catch (error) {
        console.error('Microsoft Graph test error:', error);
        res.status(500).json({
            success: false,
            message: 'Microsoft Graph test failed',
            error: error.message
        });
    }
});

module.exports = router;