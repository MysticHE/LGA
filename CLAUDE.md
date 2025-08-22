# Lead Generation Automation (LGA) - Claude Context

## Project Overview

A complete lead generation tool that automates Apollo.io scraping with AI-powered outreach generation and Microsoft Graph API integration. Features background job processing, OneDrive Excel storage, and comprehensive email automation with scheduling and tracking capabilities.

**NEW: Email Automation System v3.0** - Complete separation of scraping and email automation with master Excel file management, intelligent scheduling, and multi-template support.

## Core Architecture

### Background Job System
- **Purpose**: Avoid 5-minute infrastructure timeouts for large scraping operations
- **Flow**: Start job → Poll status → Retrieve results when complete
- **Timeout**: No time limits (removed 30-minute restriction for unlimited datasets)

### Apollo Integration (Apify-based)
- **Sync API Issue**: Apify's `/run-sync-get-dataset-items` has 300-second (5-minute) hard timeout
- **Solution**: Use async API (`/runs` endpoint) with polling mechanism
- **Dataset Retrieval**: Multi-method fallback system for robust data access

### Microsoft Graph Integration (v2.1 - Delegated Authentication)
- **OneDrive Storage**: Automatic Excel file creation with tracking columns using user's Microsoft 365 account
- **Email Automation**: Bulk email campaigns through Microsoft Graph Mail API with user authentication
- **Real-time Tracking**: Webhook notifications for email read receipts and replies
- **Authentication**: MSAL-based delegated authentication flow with popup login experience
- **User Experience**: Familiar Microsoft 365 login with session management and token refresh

### Email Automation System (v3.0 - NEW)
- **Master Excel Management**: Single `LGA-Master-Email-List.xlsx` file per user with intelligent merging
- **Dual Content System**: AI-generated personalized emails AND fixed templates with variables
- **Smart Scheduling**: Immediate and scheduled campaigns with follow-up sequences
- **Background Automation**: Cron-based email scheduler running hourly across all user sessions
- **Duplicate Prevention**: Cross-campaign email deduplication and intelligent filtering
- **Template Management**: Full CRUD operations for email templates with validation and preview

## Environment Setup

### Required Environment Variables
```bash
# Core API Keys
APIFY_API_TOKEN=your_apify_token_here
OPENAI_API_KEY=your_openai_key_here

# Microsoft Graph Integration (v2.0)
AZURE_TENANT_ID=your_tenant_id_here
AZURE_CLIENT_ID=your_client_id_here
AZURE_CLIENT_SECRET=your_client_secret_here
RENDER_EXTERNAL_URL=https://your-app-name.onrender.com

# Server Configuration
PORT=3000
NODE_ENV=development

# Rate Limiting
MAX_REQUESTS_PER_MINUTE=10
```

### Dependencies
```json
{
  "express": "^4.18.0",
  "axios": "^1.5.0",
  "openai": "^4.0.0",
  "xlsx": "^0.18.0",
  "rate-limiter-flexible": "^2.4.0",
  "multer": "^1.4.5",
  "pdf-parse": "^1.1.1",
  "@microsoft/microsoft-graph-client": "^3.0.0",
  "@azure/identity": "^4.0.0",
  "@azure/msal-node": "^2.0.0",
  "node-cron": "^3.0.0",
  "cors": "^2.8.5",
  "helmet": "^7.0.0"
}
```

### Azure App Registration Requirements (Delegated Permissions)
```
✅ Files.ReadWrite.All        # OneDrive file access
✅ Mail.Send                  # Send emails
✅ Mail.ReadWrite             # Read email status for tracking  
✅ User.Read                  # Read user profile
✅ offline_access             # Refresh tokens
```

## Key Features

### 1. Enhanced AI Email Generation with Product Materials
- **PDF Upload**: Support for multiple PDF files up to 10MB each (drag & drop interface)
- **Content Integration**: Automatically extracts text from insurance product materials
- **Enhanced Prompts**: Professional email generation using uploaded product information
- **Smart Context**: AI researches company names and generates industry-specific content
- **Token Management**: Intelligent content truncation to stay within API limits (~3K characters)
- **Professional Format**: Email-only output (removed ice breaker) with subject lines and structured content

### 2. Unlimited Lead Scraping
- **Default**: 0 (unlimited) - removed 2000 lead limit
- **Frontend**: Input allows 0-10,000 leads
- **Backend**: Handles `recordLimit = 0` for unlimited processing

### 3. Microsoft Graph Integration (NEW v2.0 - Delegated Authentication)
- **User Authentication**: MSAL-based popup login with familiar Microsoft 365 experience
- **OneDrive Excel Storage**: Save leads to user's Microsoft 365 Excel files with tracking columns
- **Email Automation**: Send personalized campaigns through user's Microsoft Graph Mail API access
- **Real-time Email Tracking**: Track email opens, reads, and replies with webhook notifications
- **Excel Sync**: Real-time updates to user's OneDrive Excel files when emails are read/replied
- **Session Management**: Secure token storage with automatic refresh capabilities
- **Campaign Analytics**: Detailed tracking and reporting for email campaigns

### 4. Async Job Processing
- **Job Creation**: Immediate response with job ID
- **Status Polling**: Real-time progress tracking
- **Result Retrieval**: Separate endpoint for large datasets
- **Extended Workflow**: Now supports up to 6 steps including OneDrive save and email campaigns

### 5. Error Handling
- **Apify Timeouts**: Automatic retry with exponential backoff
- **Dataset 404s**: Multi-method retrieval fallback
- **Rate Limiting**: Exempted polling endpoints to prevent 429 errors
- **Network Issues**: Retry logic with graceful degradation
- **Microsoft Graph Fallbacks**: Graceful handling when Graph API is unavailable

## API Endpoints

### Core Workflow (Extended v2.1)
```
POST /api/leads/start-workflow-job
├── POST /api/apollo/generate-url
├── POST /api/apollo/start-scrape-job
├── GET /api/apollo/job-status/{jobId}
├── GET /api/apollo/job-result/{jobId}
├── POST /api/leads/generate-outreach
├── POST /api/microsoft-graph/onedrive/create-excel    # OneDrive save
└── POST /api/email/send-campaign                      # Email automation
```

### Email Automation System (NEW v3.0)
```
/api/email-automation/
├── /master-list
│   ├── POST /upload - Upload and merge Excel files with duplicate detection
│   ├── GET /data - Get master list data with filtering and pagination
│   ├── GET /stats - Get dashboard statistics and analytics
│   ├── PUT /lead/:email - Update individual lead information
│   ├── GET /due-today - Get leads due for email today
│   └── GET /export - Export master list to Excel
├── /templates
│   ├── GET / - List all email templates
│   ├── POST / - Create new email template
│   ├── GET /:templateId - Get specific template
│   ├── PUT /:templateId - Update existing template
│   ├── DELETE /:templateId - Delete template
│   ├── POST /:templateId/preview - Preview template with sample data
│   ├── PATCH /:templateId/toggle - Toggle template active status
│   ├── GET /type/:templateType - Get templates by type
│   └── POST /validate - Validate template content
├── /campaigns
│   ├── POST /start - Start immediate or scheduled email campaign
│   ├── GET /:campaignId - Get campaign status and statistics
│   ├── GET / - List all campaigns with filtering
│   ├── POST /:campaignId/pause - Pause active campaign
│   ├── POST /:campaignId/resume - Resume paused campaign
│   └── POST /process-scheduled - Process scheduled campaigns (background job)
└── POST /send-email/:email - Send email to specific lead
```

### Job Management
```
GET /api/leads/job-status/{jobId}     # Poll job progress
GET /api/leads/job-result/{jobId}     # Get completed results
GET /api/leads/jobs                   # List all active jobs
```

### Microsoft Graph Integration (NEW v2.0 - Delegated Auth)
```
GET  /api/microsoft-graph/test                         # Test connection (requires session)
POST /api/microsoft-graph/onedrive/create-excel        # Create Excel in user's OneDrive
POST /api/microsoft-graph/onedrive/update-excel-tracking # Update tracking data
GET  /api/microsoft-graph/onedrive/files               # List user's OneDrive files
POST /api/email/send-campaign                          # Send email campaign
GET  /api/email/tracking/:campaignId                   # Get campaign tracking
POST /api/email/webhook/notifications                  # Webhook endpoint
GET  /api/email/webhook/notifications                  # Webhook validation
POST /api/email/webhook/subscribe                      # Create subscription
GET  /api/email/track-read                             # Pixel tracking
```

### Authentication Routes (NEW v2.0)
```
GET  /auth/login                                       # Initiate Microsoft 365 login
GET  /auth/callback                                    # OAuth2 callback handler
GET  /auth/status                                      # Check authentication status
POST /auth/logout                                      # Logout user
GET  /auth/test-graph                                  # Test authenticated Graph connection
```

### PDF Materials Management
```
POST /api/leads/upload-materials      # Upload PDF files (multipart/form-data)
GET  /api/leads/materials             # List uploaded materials
DELETE /api/leads/materials/{id}      # Delete specific material
```

## Technical Implementation

### Apollo URL Generation
- **Location**: Singapore (hardcoded)
- **Filters**: Email verified, company domain exists
- **Parameters**: Job titles, company sizes
- **Output**: Apollo.io search URL with encoded parameters

### Advanced PDF Content Optimization Process (v2.0)
1. **PDF Text Extraction**: Raw text extraction using pdf-parse library
2. **Content Processor**: Multi-stage intelligent processing engine
   - Text cleaning (remove headers, footers, artifacts)
   - Content segmentation by type (products, benefits, coverage)
   - Industry-specific scoring and relevance ranking
   - Intelligent content selection up to 3,500 characters
3. **Caching System**: Industry/role-specific caching for performance
4. **Final Email Generation**: Optimized content fed to GPT-4o-mini for email creation

### Content Processing Architecture
**Key Components:**
- **PDFContentProcessor**: Text cleaning, segmentation, and intelligent selection
- **ContentCache**: Industry/role-specific caching with LRU eviction
- **Configuration System**: Environment-specific processing parameters

**Processing Flow:**
```
PDF Upload → Content Processor → Industry Scoring → Cache Storage → Email AI Prompt
```

**Optimization Results:**
- **Content Quality**: 87% compression with preserved product specifics
- **Performance**: 1-2s processing (cached), 70%+ cache hit rate expected
- **Token Efficiency**: 3,500 chars of highly relevant content vs 13K+ raw
- **Product Preservation**: Specific product names and details maintained

### Current Content Processing Parameters
```
Character Limit: 3,500 (increased from 2,500 for better coverage)
Quality Threshold: Score ≥1 (lowered from 3 for more inclusion)
Section Scoring: Products=12, Coverage=11, Benefits=10, Business=8
Industry Optimization: Enabled with role-specific content filtering
AI Summarization: Disabled (preserves specific product details)
Caching: Enabled with 24-hour TTL and compression ratio tracking
```

### Scraping Process
1. **Generate Apollo URL** with job titles and company size filters
2. **Start Apify scraper** with async API call
3. **Poll Apify run status** every 5 seconds until completion
4. **Retrieve dataset** using multi-method fallback approach
5. **Transform data** to standardized lead format
6. **Generate AI outreach** (optional, with or without product materials)

### Data Flow (Extended v2.0)
```
Frontend Form → PDF Upload (Optional) → Apollo URL → Apify Scraper → 
Raw Leads → Enhanced AI Outreach → OneDrive Save (Optional) → Email Campaign (Optional) → 
Final Results → Excel Export → Real-time Email Tracking
```

### Microsoft Graph Workflow (NEW v2.0 - Delegated Auth)
```
1. User Authentication → 2. Generate Leads → 3. Create OneDrive Excel (User's Account) → 
4. Send Email Campaign (User's Account) → 5. Track Email Opens → 6. Update Excel File → 7. Campaign Analytics
```

### Authentication Flow (NEW)
```
1. User checks OneDrive/Email options → 2. System detects auth required → 3. Popup opens with Microsoft 365 login → 
4. User authenticates → 5. Session tokens stored → 6. Operations proceed with user's account
```

## Common Issues & Solutions

### 1. 5-Minute Timeout (SOLVED)
- **Problem**: Apify sync API timeout after 300 seconds
- **Solution**: Switched to async API with polling
- **Implementation**: Background jobs with status tracking

### 2. Dataset Retrieval 404 (SOLVED)
- **Problem**: Completed runs returning 404 when accessing dataset
- **Solution**: Multi-method retrieval with fallbacks
- **Methods**: Direct dataset → Run-based → Alternative format

### 3. Rate Limiting 429 (SOLVED)
- **Problem**: Frontend polling hitting rate limits
- **Solution**: Exempt polling endpoints from rate limiting
- **Implementation**: Middleware bypass for `/job-status/` and `/job-result/`

### 4. Large Dataset Handling
- **Chunking**: Process leads in configurable chunks (default: 100)
- **Memory Management**: Session-based retrieval for large datasets
- **Progress Tracking**: Real-time progress updates during processing

## Frontend Integration

### User Experience (Enhanced v2.0 - Delegated Auth)
- **Terminology**: "Web scraping" instead of "Apollo" for user-facing messages
- **PDF Upload**: Drag & drop interface with real-time file management
- **Product Materials**: Toggle to enable/disable enhanced AI emails
- **Microsoft 365 Integration**: OneDrive save and email campaign options with user authentication
- **Authentication**: Familiar Microsoft 365 popup login experience
- **Email Campaign Builder**: Template editor with personalization variables
- **Progress**: Real-time status updates with elapsed time display (now up to 6 steps)
- **Error Handling**: User-friendly error messages with retry suggestions
- **Session Management**: Persistent authentication across browser sessions
- **Results**: Downloadable Excel export with OneDrive links and tracking status

### PDF Upload Features
- **Drag & Drop**: Intuitive file upload interface
- **File Management**: Real-time list with file details (size, pages, date)
- **Auto-cleanup**: Materials deleted after 24 hours
- **Validation**: PDF-only uploads with size limits (10MB per file)
- **Smart Integration**: Checkbox to enable/disable material usage

### Polling Strategy
- **Interval**: 3 seconds with exponential backoff on rate limits
- **Timeout**: No frontend timeout (matches unlimited backend)
- **Error Recovery**: Automatic retry with user notification

## File Structure (Updated v4.0 - Advanced PDF Processing System)
```
├── routes/
│   ├── apollo.js                # Apify integration & async job management
│   ├── leads.js                 # Background job processing & advanced PDF optimization
│   ├── microsoft-graph.js       # OneDrive Excel integration (delegated auth)
│   ├── email-automation.js      # Email automation master list management
│   ├── email-templates.js       # Template CRUD operations
│   ├── email-scheduler.js       # Campaign management & scheduling
│   ├── auth.js                  # Microsoft 365 authentication routes
│   └── index.js                 # Main router
├── middleware/
│   ├── rateLimiter.js           # Rate limiting with polling exemptions
│   ├── graphAuth.js             # Original: Application auth (deprecated)
│   └── delegatedGraphAuth.js    # MSAL delegated authentication
├── utils/
│   ├── pdfContentProcessor.js   # NEW: Advanced PDF content processing engine
│   ├── contentAnalyzer.js       # NEW: AI-powered content analysis (deprecated - unused)
│   ├── contentCache.js          # NEW: Industry/role-specific caching system
│   ├── excelProcessor.js        # Excel file processing & master file management
│   └── emailContentProcessor.js # Email content processing & template handling
├── config/
│   └── contentConfig.js         # NEW: Content processing configuration & settings
├── jobs/
│   └── emailScheduler.js        # Background email automation scheduler
├── lead-generator.html          # Lead scraping interface (with navigation)
├── email-automation.html        # NEW: Complete email automation interface
├── server.js                    # Express server setup (updated with email automation routes)
├── RENDER-SETUP.md              # Deployment guide (updated for delegated auth)
├── IMPLEMENTATION-SUMMARY.md    # Complete feature documentation
├── AZURE-DEBUG-CHECKLIST.md     # Azure troubleshooting guide
└── CLAUDE.md                    # This file (project context)
```

## Development Workflow

### Testing (Extended v2.0)
```bash
# Core API Testing
GET /api/leads/test                  # Test OpenAI integration
GET /api/apollo/test                 # Test Apollo integration
GET /api/leads/jobs                  # List active jobs

# Microsoft Graph Testing (NEW - Delegated Auth)
GET /auth/login                       # Initiate Microsoft 365 authentication
GET /auth/status?sessionId={id}       # Check authentication status
GET /api/microsoft-graph/test         # Test connection (requires X-Session-Id header)
GET /api/microsoft-graph/onedrive/files # List user's OneDrive files

# PDF Materials Testing
GET /api/leads/materials
POST /api/leads/upload-materials (with PDF files)
DELETE /api/leads/materials/{materialId}

# Email Campaign Testing (NEW)
GET /api/email/tracking/{campaignId}  # Get campaign status
POST /api/email/webhook/subscribe     # Create webhook subscription
```

### Debugging
- **Console Logs**: Comprehensive logging with job IDs
- **Error Tracking**: Structured error responses with details
- **Job Storage**: In-memory job tracking (use Redis in production)

## Production Considerations

### Scaling
- **Job Storage**: Replace in-memory Map with Redis
- **Rate Limiting**: Adjust based on Apify plan limits
- **Monitoring**: Implement comprehensive logging and metrics
- **Error Handling**: Add dead letter queues for failed jobs

### Security
- **API Keys**: Secure environment variable storage
- **Rate Limiting**: Protect against abuse
- **Input Validation**: Sanitize all user inputs
- **CORS**: Configure appropriate CORS policies

## Recent Changes

### v3.0 - Complete Email Automation System (MAJOR UPDATE)
- **NEW: Separate Email Automation Interface**: Complete standalone `email-automation.html` with navigation
- **NEW: Master Excel File System**: Single `LGA-Master-Email-List.xlsx` per user with intelligent merging
- **NEW: Dual Content System**: Support for both AI-generated emails AND template-based emails with variables
- **NEW: Smart Duplicate Detection**: Cross-campaign email deduplication and upload merge capabilities
- **NEW: Template Management**: Full CRUD system for email templates with validation and preview
- **NEW: Campaign Scheduler**: Immediate and scheduled campaigns with follow-up automation
- **NEW: Background Email Automation**: Cron-based scheduler running hourly across all user sessions
- **NEW: Enhanced Excel Structure**: Extended columns for email automation while maintaining scraping compatibility
- **NEW: API Routes**: Complete email automation API with `/master-list`, `/templates`, and `/campaigns` endpoints
- **NEW: Utility Classes**: `ExcelProcessor` and `EmailContentProcessor` for robust file and content handling
- **NEW: Background Jobs**: `jobs/emailScheduler.js` for automated email sending with session management
- **Enhanced: Navigation**: Seamless switching between lead scraping and email automation
- **Enhanced: User Experience**: Upload Excel → Merge with master → Create campaigns → Automated follow-ups

### v4.0 - Advanced PDF Content Optimization System (MAJOR UPDATE)
- **NEW: PDFContentProcessor**: Multi-stage intelligent content processing engine with text cleaning and segmentation
- **NEW: ContentCache**: Industry/role-specific caching system with LRU eviction and 24-hour TTL
- **NEW: Configuration System**: Environment-specific processing parameters with feature flags
- **Enhanced Content Processing**: 87% compression while preserving specific product details
- **Improved Token Efficiency**: 3,500 characters of highly relevant content vs 13K+ raw PDF text
- **Industry-Specific Optimization**: Content scoring and filtering based on lead's industry and role
- **Performance Improvements**: 70%+ cache hit rate expected, 1-2s processing for cached content
- **Hybrid Approach**: Increased character limit (2,500→3,500) and lowered quality threshold for better coverage
- **Fixed Cache Metadata**: Proper compression ratio tracking and display in logs
- **Product Preservation**: Removed AI Analyzer to maintain specific product names and details from PDFs
- **New API Endpoints**: `/api/leads/content-processing-info` and `/api/leads/clear-content-cache`
- **Dependencies**: No new dependencies required - uses existing pdf-parse library

### v2.1 - Delegated Authentication Flow (CRITICAL UPDATE)
- **FIXED: Authentication Flow**: Switched from Application to Delegated authentication for proper Microsoft Graph access
- **NEW: MSAL Integration**: Complete MSAL-based authentication with popup login experience
- **NEW: Authentication Routes**: `/auth/login`, `/auth/callback`, `/auth/status`, `/auth/logout` endpoints
- **NEW: Session Management**: Secure token storage with automatic refresh capabilities
- **NEW: Frontend Popup Auth**: Seamless Microsoft 365 login popup with session persistence
- **Updated: Azure Permissions**: Changed to Delegated permissions (Files.ReadWrite.All, Mail.Send, Mail.ReadWrite, User.Read, offline_access)
- **Enhanced: User Experience**: Uses user's own Microsoft 365 account for all operations
- **Fixed: `/me` Endpoint Issues**: All OneDrive and email operations now work with delegated tokens
- **NEW: middleware/delegatedGraphAuth.js**: Complete MSAL authentication provider
- **NEW: routes/auth.js**: OAuth2 flow handling with callback management

### v2.0 - Microsoft Graph API Integration (MAJOR UPDATE)
- **NEW: OneDrive Excel Integration**: Automatically save leads to Microsoft 365 Excel files with tracking columns
- **NEW: Email Automation System**: Send bulk personalized email campaigns through Microsoft Graph Mail API
- **NEW: Real-time Email Tracking**: Webhook notifications for email reads, replies with Excel sync
- **NEW: Campaign Analytics**: Detailed tracking dashboard for email campaign performance
- **Enhanced Background Jobs**: Extended workflow from 4 to up to 6 steps with MS365 integration
- **Updated Frontend**: Microsoft 365 integration section with email campaign builder
- **New Dependencies**: @microsoft/microsoft-graph-client, @azure/identity, @azure/msal-node, node-cron
- **Comprehensive Documentation**: RENDER-SETUP.md and IMPLEMENTATION-SUMMARY.md guides
- **Webhook System**: Real-time notifications with pixel tracking and subscription management

### v1.3.2 - Enhanced Filtering Display & Results Dashboard
- **NEW: Filtered Leads Counter**: Results now display how many leads were filtered out by exclusion criteria
- **Enhanced Applied Filters Display**: Shows excluded email domains and industries in results panel
- **Visual Filter Feedback**: New "Filtered Out" stat card with orange accent when filters are applied
- **Improved Results Layout**: Dynamic stat cards that only show relevant information
- **Better User Transparency**: Complete visibility into filtering process and results impact

### v1.3.1 - Critical Bug Fixes for Exclusion Filters & PDF Integration
- **FIXED: Exclusion Filters Not Working**: Added missing `excludeEmailDomains` and `excludeIndustries` parameters to frontend API call
- **FIXED: PDF Materials Not Being Used**: Added missing `useProductMaterials` parameter to frontend API call
- **Enhanced Email Filtering**: Both email domain and industry exclusion filters now work correctly during lead processing
- **Verified PDF Integration**: Confirmed PDF content is properly extracted and passed to OpenAI for enhanced email generation
- **Improved Debugging**: Console logs now show filtered leads and material usage status

### v1.3 - Enhanced AI with Product Materials
- **PDF Upload System**: Drag & drop interface for SME insurance materials
- **Enhanced AI Prompts**: Professional email-only format with company research
- **Content Integration**: Smart truncation and token management
- **Material Management**: In-memory storage with auto-cleanup after 24 hours
- **Improved UX**: Toggle to enable/disable enhanced AI features
- **Dependencies**: Added multer and pdf-parse for file processing

### v1.2 - Unlimited Leads & Terminology
- Removed 2000 lead limit, now defaults to unlimited (0)
- Replaced "Apollo" with "Web" in user-facing messages
- Removed 30-minute polling timeout for very large datasets

### v1.1 - Async Job System
- Implemented background job processing
- Added real-time progress tracking
- Multi-method dataset retrieval with fallbacks

### v1.0 - Initial Implementation
- Basic Apollo URL generation
- Synchronous scraping (limited to 5 minutes)
- OpenAI outreach generation