// ADVANCED EXCEL UPLOAD FIX FOR MICROSOFT GRAPH API
// This implements the official Microsoft recommendations for reliable Excel file uploads

const axios = require('axios');
const XLSX = require('xlsx');

/**
 * Advanced Excel file upload using Microsoft Graph API Upload Session
 * Based on official Microsoft documentation and best practices
 * @param {Object} client - Microsoft Graph client
 * @param {Buffer} fileBuffer - Excel file buffer
 * @param {string} filename - Target filename
 * @param {string} folderPath - OneDrive folder path
 * @returns {Object} Upload result
 */
async function advancedExcelUpload(client, fileBuffer, filename, folderPath) {
    console.log(`🚀 ADVANCED EXCEL UPLOAD: ${filename} (${fileBuffer.length} bytes)`);
    
    // Step 1: Validate Excel file integrity
    if (!validateExcelFile(fileBuffer)) {
        throw new Error('Invalid Excel file format');
    }
    
    // Step 2: Create upload session (Microsoft's official recommendation)
    const uploadSession = await createUploadSession(client, filename, folderPath);
    console.log(`📋 Upload session created: ${uploadSession.uploadUrl}`);
    
    // Step 3: Upload using raw HTTP request (avoids Graph SDK overhead)
    const uploadResult = await uploadFileToSession(uploadSession.uploadUrl, fileBuffer);
    console.log(`✅ File uploaded successfully`);
    
    // Step 4: Verify upload integrity
    await verifyUploadedExcelFile(client, fileBuffer, filename, folderPath);
    console.log(`✅ Upload verification passed`);
    
    return uploadResult;
}

/**
 * Validate Excel file buffer integrity
 */
function validateExcelFile(fileBuffer) {
    if (!Buffer.isBuffer(fileBuffer)) {
        console.error('❌ Invalid input: not a Buffer');
        return false;
    }
    
    if (fileBuffer.length < 100) {
        console.error('❌ Invalid input: file too small');
        return false;
    }
    
    // Check for Excel signature (ZIP file format)
    const signature = fileBuffer.slice(0, 2).toString('hex');
    if (signature !== '504b') {
        console.error(`❌ Invalid Excel signature: ${signature} (expected: 504b)`);
        return false;
    }
    
    // Try to read as Excel workbook
    try {
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            console.error('❌ No sheets found in workbook');
            return false;
        }
        console.log(`✅ Valid Excel file with sheets: [${workbook.SheetNames.join(', ')}]`);
        return true;
    } catch (error) {
        console.error(`❌ Excel validation failed: ${error.message}`);
        return false;
    }
}

/**
 * Create upload session using Microsoft Graph API
 */
async function createUploadSession(client, filename, folderPath) {
    try {
        const sessionUrl = `/me/drive/root:${folderPath}/${filename}:/createUploadSession`;
        
        const uploadSession = await client.api(sessionUrl).post({
            item: {
                '@microsoft.graph.conflictBehavior': 'replace',
                name: filename
            }
        });
        
        if (!uploadSession.uploadUrl) {
            throw new Error('No upload URL in session response');
        }
        
        return uploadSession;
    } catch (error) {
        console.error(`❌ Failed to create upload session: ${error.message}`);
        throw error;
    }
}

/**
 * Upload file to session using direct HTTP request
 * This bypasses the Graph SDK to avoid any encoding issues
 */
async function uploadFileToSession(uploadUrl, fileBuffer) {
    try {
        console.log(`📤 Uploading ${fileBuffer.length} bytes to session...`);
        
        const response = await axios.put(uploadUrl, fileBuffer, {
            headers: {
                'Content-Length': fileBuffer.length.toString(),
                'Content-Range': `bytes 0-${fileBuffer.length - 1}/${fileBuffer.length}`,
                // NOTE: No Content-Type header - let Microsoft Graph handle it
            },
            maxContentLength: Infinity,
            maxBodyLength: Infinity,
            timeout: 60000 // 60 second timeout
        });
        
        if (response.status !== 200 && response.status !== 201) {
            throw new Error(`Upload failed: HTTP ${response.status} - ${response.statusText}`);
        }
        
        console.log(`✅ Upload completed: ${response.status} - ${response.statusText}`);
        return response.data;
        
    } catch (error) {
        if (error.response) {
            console.error(`❌ Upload HTTP error: ${error.response.status} - ${error.response.statusText}`);
            console.error(`❌ Response data:`, error.response.data);
        }
        throw new Error(`Session upload failed: ${error.message}`);
    }
}

/**
 * Verify uploaded Excel file integrity
 */
async function verifyUploadedExcelFile(client, originalBuffer, filename, folderPath, maxRetries = 3) {
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            console.log(`🔍 Verification attempt ${attempt}/${maxRetries}...`);
            
            // Wait for OneDrive processing (increases with each attempt)
            const waitTime = attempt * 2000; // 2s, 4s, 6s
            await new Promise(resolve => setTimeout(resolve, waitTime));
            
            // Download the uploaded file
            const downloadedBuffer = await client.api(`/me/drive/root:${folderPath}/${filename}:/content`).get();
            const properBuffer = Buffer.isBuffer(downloadedBuffer) ? downloadedBuffer : Buffer.from(downloadedBuffer);
            
            console.log(`📥 Downloaded ${properBuffer.length} bytes (original: ${originalBuffer.length} bytes)`);
            
            // Size check
            if (properBuffer.length !== originalBuffer.length) {
                throw new Error(`Size mismatch: downloaded ${properBuffer.length} bytes, expected ${originalBuffer.length} bytes`);
            }
            
            // Structure verification
            const workbook = XLSX.read(properBuffer, { type: 'buffer' });
            console.log(`📊 Verified sheets: [${workbook.SheetNames.join(', ')}]`);
            
            // Check required sheets
            const requiredSheets = ['Leads', 'Templates', 'Campaign_History'];
            const missingSheets = requiredSheets.filter(sheet => !workbook.SheetNames.includes(sheet));
            
            if (missingSheets.length > 0) {
                throw new Error(`Missing required sheets: ${missingSheets.join(', ')}`);
            }
            
            // Check Leads sheet data
            const leadsSheet = workbook.Sheets['Leads'];
            const leadsData = XLSX.utils.sheet_to_json(leadsSheet);
            console.log(`✅ Verification successful: Leads sheet has ${leadsData.length} rows`);
            
            return true;
            
        } catch (error) {
            console.error(`❌ Verification attempt ${attempt} failed: ${error.message}`);
            
            if (attempt === maxRetries) {
                throw new Error(`Upload verification failed after ${maxRetries} attempts: ${error.message}`);
            }
        }
    }
}

module.exports = {
    advancedExcelUpload,
    validateExcelFile,
    createUploadSession,
    uploadFileToSession,
    verifyUploadedExcelFile
};