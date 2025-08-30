# Email Tracking Debug Test

## Issue
Email `hui.en94@hotmail.com` tracking pixel hit but Excel update failed with "Lead not found" error.

## Debugging Steps

### 1. Test Manual Update (This will trigger detailed debug logs)
```bash
curl -X POST "https://your-app.onrender.com/api/email/test-read-update" \
-H "Content-Type: application/json" \
-H "X-Session-Id: YOUR_SESSION_ID" \
-d '{"email": "hui.en94@hotmail.com", "testType": "read"}'
```

### 2. Check Diagnostic Data
```bash
curl "https://your-app.onrender.com/api/email/diagnostic/hui.en94@hotmail.com" \
-H "X-Session-Id: YOUR_SESSION_ID"
```

### 3. View All Tracking Data
```bash
curl "https://your-app.onrender.com/api/email/diagnostic" \
-H "X-Session-Id: YOUR_SESSION_ID"
```

## Expected Debug Output

The enhanced logging will now show:
- **📊 Available sheets**: ["Leads", "Templates"] 
- **📧 Column names**: All email-related columns in the file
- **🔍 First 5 emails**: Actual email addresses in the Excel file
- **📋 Total leads**: Count of leads in the file

## Possible Root Causes

1. **Session Mismatch**: Tracking uses different session than where email was sent
2. **File Caching**: Old version of Excel file being downloaded  
3. **Column Name Issue**: Email stored in different column than expected
4. **Data Format Issue**: Email has extra spaces, encoding, or formatting

## Next Steps

After running the test above, check the server logs for the detailed debug output to identify the exact cause.