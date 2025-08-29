# Azure Microsoft Graph Integration Debug Checklist

## ✅ Environment Variables (Confirmed from your screenshot)
- [x] `AZURE_TENANT_ID`: 496f1a0a-6a4a-4436-b4b3-fdb75d236254
- [x] `AZURE_CLIENT_ID`: d6042b07-f7dc-4706-bf6f-847a7bd1530d  
- [x] `AZURE_CLIENT_SECRET`: 6eL8Q-oD2j7ZrPokejI0pwuxs.MtEXddsxzxwc5n
- [x] `RENDER_EXTERNAL_URL`: https://lga-4eou.onrender.com

## 🔍 Azure App Registration Checklist

### 1. App Registration Basic Setup
- [ ] App is created in Azure AD
- [ ] App supports "Accounts in this organizational directory only"
- [ ] Client secret is not expired

### 2. API Permissions (Required)
Check these permissions are granted in Azure Portal > App registrations > Your app > API permissions:

**Microsoft Graph Application Permissions:**
- [ ] `Files.ReadWrite.All` - OneDrive file access
- [ ] `Mail.Send` - Send emails  
- [ ] `Mail.ReadWrite.All` - Read email status for tracking
- [ ] `User.Read.All` - Read user profiles

**Important:** These must be **Application permissions** (not Delegated permissions)

### 3. Admin Consent
- [ ] Click "Grant admin consent for [Your Organization]" button
- [ ] All permissions show green checkmarks with "Granted for [Your Organization]"

### 4. Authentication Settings
- [ ] Under "Authentication" > "Platform configurations"
- [ ] Add Web platform with redirect URI: `https://lga-4eou.onrender.com/api/email/webhook/notifications`

## 🔧 Common Issues & Solutions

### Issue 1: "insufficient_privileges" error
**Cause:** Missing admin consent or wrong permission type
**Solution:** 
1. Go to API permissions
2. Ensure all permissions are **Application** type (not Delegated)
3. Click "Grant admin consent"

### Issue 2: "invalid_client" error  
**Cause:** Wrong Client ID or expired Client Secret
**Solution:**
1. Verify AZURE_CLIENT_ID matches Application (client) ID exactly
2. Create new client secret if expired
3. Update AZURE_CLIENT_SECRET in Render

### Issue 3: "invalid_tenant" error
**Cause:** Wrong Tenant ID
**Solution:** Verify AZURE_TENANT_ID matches Directory (tenant) ID exactly

### Issue 4: Connection timeout
**Cause:** Network or authentication provider issues
**Solution:** Check Render logs for detailed error messages

## 🧪 Testing Commands

### Test from Render Service:
```bash
curl https://lga-4eou.onrender.com/api/microsoft-graph/test
```

### Expected Success Response:
```json
{
  "success": true,
  "message": "Microsoft Graph connection successful",
  "user": "Your Display Name",
  "oneDrive": {
    "name": "OneDrive",
    "owner": "Your Name",
    "quota": {...}
  }
}
```

### Expected Error Response:
```json
{
  "success": false,
  "error": "Microsoft Graph Authentication Error",
  "message": "Failed to authenticate with Microsoft Graph",
  "details": "Specific error message here"
}
```

## 📋 Debug Steps

1. **Check Render Logs:** Look for startup errors
2. **Test Endpoint:** Run curl command above
3. **Verify Azure Setup:** Go through checklist above
4. **Check Permissions:** Ensure admin consent granted
5. **Verify Client Secret:** Make sure it's not expired

## 🆘 If Still Failing

Share the exact error message from:
1. Render service logs
2. curl test result
3. Browser network tab when testing connection