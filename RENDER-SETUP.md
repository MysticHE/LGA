# Render Deployment Setup Guide

## Microsoft Graph API Integration for Lead Generation Platform

This guide covers the complete setup process for deploying your Lead Generation platform with Microsoft Graph API integration on Render.

## Table of Contents
1. [Azure App Registration Setup](#azure-app-registration-setup)
2. [Render Environment Variables](#render-environment-variables)
3. [Webhook Configuration](#webhook-configuration)
4. [Testing the Integration](#testing-the-integration)
5. [Troubleshooting](#troubleshooting)

---

## Azure App Registration Setup

### 1. Create Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Fill in the details:
   - **Name**: `LGA-Microsoft-Graph-Integration`
   - **Supported account types**: `Accounts in this organizational directory only`
   - **Redirect URI**: Leave blank for now (we'll add it later)
5. Click **Register**

### 2. Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions** (not Delegated)
5. Add the following permissions:
   ```
   Files.ReadWrite.All        # OneDrive file access
   Mail.Send                  # Send emails
   Mail.ReadWrite.All         # Read email status for tracking
   User.Read.All              # Read user profiles
   ```
6. Click **Add permissions**
7. Click **Grant admin consent** (you need admin privileges)

### 3. Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add description: `LGA-Render-Integration`
4. Set expiry: `24 months` (or as per your organization policy)
5. Click **Add**
6. **IMPORTANT**: Copy the secret value immediately (you won't see it again)

### 4. Note Down Required Values

From your Azure app registration, collect these values:
- **Application (client) ID**: Found on Overview page
- **Directory (tenant) ID**: Found on Overview page  
- **Client secret**: The value you copied in step 3

---

## Render Environment Variables

### 1. Access Render Dashboard

1. Go to [Render Dashboard](https://dashboard.render.com)
2. Navigate to your deployed service
3. Go to **Environment** tab

### 2. Add Required Environment Variables

Add the following environment variables in Render:

#### Core Application Variables (existing)
```
APIFY_API_TOKEN=your_apify_token_here
OPENAI_API_KEY=your_openai_key_here
PORT=3000
NODE_ENV=production
MAX_REQUESTS_PER_MINUTE=10
```

#### Microsoft Graph Integration Variables (new)
```
AZURE_TENANT_ID=your_tenant_id_from_step_4
AZURE_CLIENT_ID=your_client_id_from_step_4
AZURE_CLIENT_SECRET=your_client_secret_from_step_4
RENDER_EXTERNAL_URL=https://your-app-name.onrender.com
```

**Note**: Replace `your-app-name` with your actual Render service name.

### 3. Example Configuration

```bash
# Core Variables
APIFY_API_TOKEN=apify_api_1234567890abcdef
OPENAI_API_KEY=sk-1234567890abcdef
PORT=3000
NODE_ENV=production

# Microsoft Graph Variables  
AZURE_TENANT_ID=12345678-1234-1234-1234-123456789012
AZURE_CLIENT_ID=87654321-4321-4321-4321-210987654321
AZURE_CLIENT_SECRET=ABC123def456GHI789jkl012MNO345pqr678STU901
RENDER_EXTERNAL_URL=https://lga-platform.onrender.com
```

---

## Webhook Configuration

### 1. Update Azure Redirect URIs

1. In Azure Portal, go to your app registration
2. Navigate to **Authentication**
3. Under **Platform configurations**, click **Add a platform**
4. Select **Web**
5. Add redirect URI: `https://your-app-name.onrender.com/api/email/webhook/notifications`
6. Click **Configure**

### 2. Webhook Endpoints

The system will automatically set up these webhook endpoints:

```
GET  /api/email/webhook/notifications     # Webhook validation
POST /api/email/webhook/notifications     # Receive notifications  
POST /api/email/webhook/subscribe         # Create subscription
```

### 3. Test Webhook Subscription

After deployment, test the webhook by calling:
```bash
curl -X POST https://your-app-name.onrender.com/api/email/webhook/subscribe \
  -H "Content-Type: application/json"
```

---

## Testing the Integration

### 1. Test Microsoft Graph Connection

```bash
curl https://your-app-name.onrender.com/api/microsoft-graph/test
```

Expected response:
```json
{
  "success": true,
  "message": "Microsoft Graph connection successful",
  "user": "Your User Name",
  "oneDrive": {
    "name": "OneDrive",
    "owner": "Your Name",
    "quota": {...}
  }
}
```

### 2. Test OneDrive Excel Creation

```bash
curl -X POST https://your-app-name.onrender.com/api/microsoft-graph/onedrive/create-excel \
  -H "Content-Type: application/json" \
  -d '{
    "leads": [
      {
        "name": "John Doe",
        "email": "john@example.com", 
        "organization_name": "Test Company"
      }
    ],
    "filename": "test-leads.xlsx"
  }'
```

### 3. Test Email Campaign

```bash
curl -X POST https://your-app-name.onrender.com/api/email/send-campaign \
  -H "Content-Type: application/json" \
  -d '{
    "leads": [
      {
        "name": "John Doe",
        "email": "john@example.com",
        "organization_name": "Test Company"
      }
    ],
    "subject": "Test Email",
    "emailTemplate": "<h1>Hello {name}!</h1><p>This is a test email for {company}.</p>",
    "trackReads": true
  }'
```

---

## Frontend Integration

The platform now includes these new features:

### 1. OneDrive Save Option
- Results panel includes "Save to OneDrive" button
- Automatically creates Excel file with tracking columns
- Provides direct link to OneDrive file

### 2. Email Campaign Interface
- Send bulk emails with personalized content
- Real-time tracking dashboard
- Status indicators: Sent, Read, Replied

### 3. Enhanced Workflow
```
Generate Leads → AI Content → Save to OneDrive → Send Email Campaign → Track Results
```

---

## API Endpoints Summary

### Microsoft Graph Routes
```
GET  /api/microsoft-graph/test                           # Test connection
POST /api/microsoft-graph/onedrive/create-excel          # Create Excel in OneDrive
POST /api/microsoft-graph/onedrive/update-excel-tracking # Update tracking data
GET  /api/microsoft-graph/onedrive/files                 # List OneDrive files
```

### Email Automation Routes
```
POST /api/email/send-campaign              # Send email campaign
GET  /api/email/tracking/:campaignId       # Get campaign tracking
POST /api/email/webhook/notifications      # Webhook endpoint
GET  /api/email/webhook/notifications       # Webhook validation
POST /api/email/webhook/subscribe          # Create subscription
GET  /api/email/track-read                 # Pixel tracking endpoint
```

---

## Troubleshooting

### Common Issues

#### 1. "Microsoft Graph Authentication Error"
- **Cause**: Missing or incorrect Azure credentials
- **Solution**: Verify environment variables in Render dashboard
- **Check**: Ensure client secret hasn't expired

#### 2. "Failed to create webhook subscription"  
- **Cause**: Incorrect RENDER_EXTERNAL_URL or missing permissions
- **Solution**: Verify URL is correct and app has Mail.ReadWrite.All permission

#### 3. "OneDrive access denied"
- **Cause**: Missing Files.ReadWrite.All permission
- **Solution**: Add permission in Azure and grant admin consent

#### 4. "Email sending failed"
- **Cause**: Missing Mail.Send permission or authentication issue
- **Solution**: Verify permissions and test connection endpoint

### Debug Steps

1. **Check Render logs** for detailed error messages
2. **Test connection endpoint** to verify Azure integration
3. **Verify environment variables** are correctly set
4. **Check Azure app permissions** and admin consent status

### Support Logs

Enable detailed logging by setting:
```
NODE_ENV=development  # Temporarily for debugging
```

---

## Security Considerations

1. **Client Secret Management**: Store only in Render environment variables
2. **Webhook Validation**: System validates webhook tokens automatically  
3. **Rate Limiting**: Built-in protection against API abuse
4. **CORS Configuration**: Restricted to your domain in production
5. **Data Encryption**: All communications use HTTPS

---

## Cost Optimization

1. **Microsoft Graph API**: Uses efficient batching to minimize API calls
2. **Webhook Subscriptions**: Automatically renewed to maintain tracking
3. **OneDrive Storage**: Files organized in dedicated folders
4. **Rate Limiting**: Respects Microsoft Graph API limits

---

## Next Steps

After successful setup:

1. **Test with small lead batch** (5-10 leads)
2. **Verify Excel file creation** in OneDrive
3. **Send test email campaign** to your own email
4. **Check email tracking** functionality
5. **Scale up to production volumes**

## Support

For additional help:
- Check Render deployment logs
- Verify Azure app registration settings
- Test individual API endpoints
- Monitor webhook subscription status

The system provides comprehensive logging to help diagnose any integration issues.