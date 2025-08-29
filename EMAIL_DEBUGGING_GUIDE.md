# Email Automation Debugging Guide

## 🔍 Issue Analysis: "Failed to send emails: The requested resource was not found"

### Root Cause Identified ✅

The error occurs because **Azure/Microsoft Graph credentials are missing** from your Render environment variables, causing the authentication middleware to fail.

### Error Chain:
1. User clicks "Send New Emails" button
2. Frontend calls `/api/email-scheduler/campaigns/start`
3. Server attempts to initialize Microsoft Graph authentication
4. `requireDelegatedAuth` middleware fails due to missing Azure credentials
5. Returns generic "resource not found" error to user

## 🛠️ Solution Implemented

### 1. Enhanced Error Detection
- Server now checks Azure credentials on startup
- Provides clear warnings when credentials are missing
- Gracefully disables email features if configuration is incomplete

### 2. Improved Error Messages
- **Frontend**: Detailed error messages with troubleshooting steps
- **Backend**: Comprehensive logging with specific missing credentials
- **Middleware**: Enhanced authentication debugging with session tracking

### 3. Better User Experience
- Clear service unavailable messages when Azure is not configured
- Step-by-step troubleshooting guidance
- Debug information for developers

## 🔧 Render Environment Variables Required

Your Render service needs these environment variables set:

```
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-app-registration-client-id
AZURE_CLIENT_SECRET=your-app-registration-secret
```

## 📋 To Fix the Email Automation:

### Step 1: Check Current Environment Variables
1. Go to your Render dashboard
2. Select your service
3. Go to "Environment" tab
4. Verify the three Azure variables are present and have values

### Step 2: Create Azure App Registration (if not done)
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to "Azure Active Directory" > "App registrations"
3. Click "New registration"
4. Set name: "LGA Email Automation"
5. Set redirect URI: `https://your-render-app.onrender.com/auth/callback`

### Step 3: Configure API Permissions
Required Microsoft Graph permissions:
- `User.Read` (Delegated)
- `Files.ReadWrite.All` (Delegated) 
- `Mail.Send` (Delegated)
- `Mail.ReadWrite` (Delegated)

### Step 4: Generate Client Secret
1. Go to "Certificates & secrets"
2. Click "New client secret" 
3. Copy the secret value immediately (you can't see it again)

### Step 5: Update Render Environment Variables
Add the three values to your Render environment:
- `AZURE_TENANT_ID`: From Azure AD overview page
- `AZURE_CLIENT_ID`: From app registration overview
- `AZURE_CLIENT_SECRET`: The secret you just generated

### Step 6: Restart Your Render Service
After adding the environment variables, restart your service for changes to take effect.

## 🚀 Testing After Fix

1. Restart your Render service
2. Check logs for: `✅ Email automation services initialized`
3. Go to email automation page
4. Try uploading a file or sending emails
5. Should see proper authentication flow instead of "resource not found"

## 📊 Debugging Features Added

### Server Logs
- Clear Azure credential validation on startup
- Detailed authentication middleware logging
- Microsoft Graph API call tracing

### Frontend Error Handling  
- Service unavailable detection with admin guidance
- Authentication debugging information
- Session management troubleshooting

### API Response Enhancement
- Structured error responses with troubleshooting steps
- Missing credential identification
- Authentication status tracking

## ❓ Common Issues & Solutions

### "Service Unavailable" Error
**Cause**: Missing Azure credentials  
**Fix**: Add the three Azure environment variables to Render

### "Authentication Required" Error  
**Cause**: User not logged into Microsoft 365  
**Fix**: Click "Connect to Microsoft 365" and complete login flow

### "Session not found" Error
**Cause**: Session expired or invalid  
**Fix**: Refresh page and re-authenticate

### Server Won't Start
**Cause**: Invalid Azure credential format  
**Fix**: Verify credentials are correct, no extra spaces/characters

---

## 🎯 Summary

The "Send New Emails" button was failing because the application couldn't authenticate with Microsoft Graph due to missing Azure credentials. With the enhanced error handling and debugging features now in place, you'll get clear guidance on exactly what's missing and how to fix it.

Once you add the Azure environment variables to Render, the email automation will work properly! 🚀