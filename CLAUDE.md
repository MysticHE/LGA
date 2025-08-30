# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Lead Generation Automation (LGA) platform that replicates n8n workflows with Apollo.io scraping and AI-powered outreach generation. The system integrates with Microsoft Graph API for email automation and OneDrive file management.

**Core Features:**
- Apollo.io lead scraping via Apify API
- AI-powered outreach content generation using OpenAI
- Microsoft Graph integration for email automation and OneDrive storage
- Excel file management with table append functionality
- Real-time progress tracking and background job processing

## Development Commands

### Essential Commands
```bash
# Install dependencies
npm install

# Development mode (auto-restart on changes)
npm run dev

# Production mode
npm start

# Install dependencies (alternative)
npm run install-deps

# Build (no-op - Node.js app)
npm run build
```

### Environment Setup
```bash
# Copy environment template (if exists)
copy .env.example .env

# Required environment variables:
APIFY_API_TOKEN=your_apify_token
OPENAI_API_KEY=your_openai_key
AZURE_TENANT_ID=your_tenant_id
AZURE_CLIENT_ID=your_client_id  
AZURE_CLIENT_SECRET=your_client_secret
RENDER_EXTERNAL_URL=https://your-app.onrender.com
PORT=3000
NODE_ENV=production
```

## Architecture Overview

### High-Level Architecture
```
Frontend (lead-generator.html) 
    ↓ REST API calls
Express.js Server (server.js)
    ↓ Route handlers
Business Logic (routes/*)
    ↓ External APIs
Apollo/Apify ← → OpenAI ← → Microsoft Graph
```

### Core Components

**Server Entry Point:** `server.js`
- Express.js application with security middleware (Helmet, CORS)
- Rate limiting and error handling
- Conditional Azure services initialization
- Serves static files and API routes

**Route Architecture:**
- `routes/apollo.js` - Apollo.io URL generation and Apify lead scraping
- `routes/leads.js` - Lead processing, OpenAI integration, and Excel export
- `routes/microsoft-graph.js` - OneDrive file management and Excel table operations
- `routes/email-automation.js` - Email campaign management and tracking
- `routes/email-tracking.js` - Email read receipts and webhook handling
- `routes/auth.js` - Microsoft Graph authentication flows

**Middleware:**
- `middleware/rateLimiter.js` - API rate limiting protection
- `middleware/delegatedGraphAuth.js` - Microsoft Graph delegated authentication

**Utilities:**
- `utils/excelProcessor.js` - Excel file creation and manipulation
- `utils/emailContentProcessor.js` - Email template processing and personalization
- `utils/contentAnalyzer.js` - Content analysis and optimization
- `utils/pdfContentProcessor.js` - PDF content extraction

**Background Jobs:**
- `jobs/emailScheduler.js` - Scheduled email campaigns and tracking updates

### Data Flow Architecture

**Lead Generation Workflow:**
1. Form submission → Apollo URL generation (`routes/apollo.js`)
2. Apify scraper → Lead data extraction (`routes/apollo.js`)
3. OpenAI processing → Personalized outreach generation (`routes/leads.js`)
4. Excel creation → OneDrive storage with table append (`routes/microsoft-graph.js`)
5. Optional: Email campaign execution (`routes/email-automation.js`)

**Email Automation Workflow:**
1. Campaign creation → Bulk email processing (`routes/email-automation.js`)
2. Email tracking → Webhook notifications (`routes/email-tracking.js`)
3. Status updates → Excel table updates (`routes/microsoft-graph.js`)
4. Scheduled tasks → Background job processing (`jobs/emailScheduler.js`)

### Key Integration Points

**Microsoft Graph Integration:**
- **Authentication:** Delegated permissions via MSAL (not client credentials)
- **OneDrive:** Excel file creation with automatic table management
- **Outlook:** Email sending and read tracking via webhooks
- **Excel Tables:** Append-only operations to preserve existing data

**External API Dependencies:**
- **Apify:** Apollo.io lead scraping (requires APIFY_API_TOKEN)
- **OpenAI:** Content generation (requires OPENAI_API_KEY) 
- **Microsoft Graph:** File/email operations (requires Azure app registration)

## Testing and Debugging

### Health Check Endpoints
```bash
# Server health
GET /health

# API connection tests
GET /api/apollo/test
GET /api/leads/test  
GET /api/microsoft-graph/test
```

### Common Development Tasks

**Adding New Routes:**
- Create route file in `routes/` directory
- Follow existing pattern with proper error handling
- Add route import and mounting in `server.js`
- Include rate limiting if needed

**Excel Table Operations:**
- All Excel operations use table append (never replace)
- Automatic table creation for new files
- Use `routes/microsoft-graph.js` endpoints for file operations

**Email Campaign Features:**
- Use webhook-based tracking for read receipts
- Background job processing for large campaigns
- Status updates automatically sync to Excel tables

### Production Considerations

**Security:**
- All API keys stored in environment variables
- Helmet.js security headers configured
- CORS restricted to specific origins
- Rate limiting enabled (10 requests/minute default)

**Performance:**
- Background job processing for large operations
- Excel table operations for efficiency
- Caching for content analysis results

**Monitoring:**
- Comprehensive error logging throughout
- Health check endpoints for service monitoring
- Webhook validation for email tracking reliability

## Recent Fixes and Improvements

### Email Subject/Body Parsing (Fixed)
**Issue:** Email subjects were not being extracted correctly from AI-generated content, causing fallback subjects like "Connecting with {Company}" to be used instead of the actual AI-generated subject lines.

**Root Cause:** Regex patterns in `utils/emailContentProcessor.js` were too rigid and couldn't handle variations in AI-generated content formatting.

**Solution Implemented:**
- **Robust Regex Patterns:** Updated to handle extra spaces, BOM characters, and numbered prefixes
- **Flexible Subject Extraction:** Handles `Subject Line:`, `1. Subject Line:`, and spacing variations
- **Enhanced Body Cleaning:** Removes numbered email body labels like `2. Email Body:`
- **Proper Fallback Logic:** Only uses fallback when genuinely no subject found

**Key Changes:**
```javascript
// Improved regex in parseEmailContent()
const subjectMatch = aiContent.match(/(?:Subject\s*Line:|^\s*\d+\.\s*Subject\s*Line:)\s*(.+)/im);

// Enhanced body cleaning
body = body.replace(/^(?:\d+\.\s*)?Email\s*Body:\s*/im, '').trim();

// BOM character handling
aiContent = aiContent.replace(/^\uFEFF/, '').trim();
```

### Excel Column Address Calculation (Fixed)
**Issue:** Graph API errors when updating Excel columns beyond Z (AA, AB, etc.) due to incorrect cell address calculation.

**Root Cause:** Using `String.fromCharCode(65 + colIndex)` which only works for columns A-Z (0-25).

**Solution:** Added `getExcelColumnLetter()` helper function to properly handle multi-letter column names (AA, AB, AC, etc.).

### Excel Structure Optimization (Completed)
**Changes Made:**
- Removed 4 unnecessary columns: `Email_Choice`, `Email_Content_Sent`, `Auto_Send_Enabled`, `Max_Emails`
- Reduced Excel from 28 to 24 columns for cleaner structure
- Maintained all essential tracking and automation functionality

## Important Notes

- **No Test Framework:** This project has no automated tests. Manual testing required.
- **Environment Dependencies:** Azure credentials required for full functionality - system gracefully degrades without them
- **Excel Integration:** Uses table append operations to prevent data loss
- **Authentication Flow:** Uses delegated permissions (user auth) not application permissions
- **Rate Limiting:** Built-in protection - adjust MAX_REQUESTS_PER_MINUTE if needed
- **Email Content Processing:** AI-generated content is automatically parsed to extract subjects and clean email bodies
- **Excel Column Support:** Supports unlimited Excel columns (A-Z, AA-AB, etc.) with proper Graph API integration