# Lead Generation Automation (LGA) - Claude Context

## Project Overview

A complete lead generation tool that automates Apollo.io scraping with AI-powered outreach generation and Microsoft Graph API integration. Features background job processing, OneDrive Excel storage, and email automation with tracking capabilities.

## Core Architecture

### Background Job System
- **Purpose**: Avoid 5-minute infrastructure timeouts for large scraping operations
- **Flow**: Start job → Poll status → Retrieve results when complete
- **Timeout**: No time limits (removed 30-minute restriction for unlimited datasets)

### Apollo Integration (Apify-based)
- **Sync API Issue**: Apify's `/run-sync-get-dataset-items` has 300-second (5-minute) hard timeout
- **Solution**: Use async API (`/runs` endpoint) with polling mechanism
- **Dataset Retrieval**: Multi-method fallback system for robust data access

### Microsoft Graph Integration (v2.0)
- **OneDrive Storage**: Automatic Excel file creation with tracking columns
- **Email Automation**: Bulk email campaigns through Microsoft Graph Mail API
- **Real-time Tracking**: Webhook notifications for email read receipts and replies
- **Authentication**: Client Secret Credential flow for server-side operations

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
  "node-cron": "^3.0.0"
}
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

### 3. Microsoft Graph Integration (NEW v2.0)
- **OneDrive Excel Storage**: Automatically save leads to Microsoft 365 Excel files with tracking columns
- **Email Automation**: Send bulk personalized email campaigns through Microsoft Graph Mail API
- **Real-time Email Tracking**: Track email opens, reads, and replies with webhook notifications
- **Excel Sync**: Real-time updates to OneDrive Excel files when emails are read/replied
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

### Core Workflow (Extended v2.0)
```
POST /api/leads/start-workflow-job
├── POST /api/apollo/generate-url
├── POST /api/apollo/start-scrape-job
├── GET /api/apollo/job-status/{jobId}
├── GET /api/apollo/job-result/{jobId}
├── POST /api/leads/generate-outreach
├── POST /api/microsoft-graph/onedrive/create-excel    # NEW: OneDrive save
└── POST /api/email/send-campaign                      # NEW: Email automation
```

### Job Management
```
GET /api/leads/job-status/{jobId}     # Poll job progress
GET /api/leads/job-result/{jobId}     # Get completed results
GET /api/leads/jobs                   # List all active jobs
```

### Microsoft Graph Integration (NEW v2.0)
```
GET  /api/microsoft-graph/test                         # Test connection
POST /api/microsoft-graph/onedrive/create-excel        # Create Excel in OneDrive
POST /api/microsoft-graph/onedrive/update-excel-tracking # Update tracking data
GET  /api/microsoft-graph/onedrive/files               # List OneDrive files
POST /api/email/send-campaign                          # Send email campaign
GET  /api/email/tracking/:campaignId                   # Get campaign tracking
POST /api/email/webhook/notifications                  # Webhook endpoint
GET  /api/email/webhook/notifications                  # Webhook validation
POST /api/email/webhook/subscribe                      # Create subscription
GET  /api/email/track-read                             # Pixel tracking
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

### Enhanced AI Email Generation Process
1. **PDF Processing**: Extract text content from uploaded insurance materials
2. **Content Integration**: Combine product materials with lead information (limited to 3K chars)
3. **Company Research**: AI analyzes company name for business context
4. **Professional Email**: Generate subject line and structured email content using GPT-4o-mini
5. **Token Optimization**: Smart truncation to fit within API limits (500 max tokens, 0.7 temperature)

### Current AI Prompt Structure
```
Professional SME Insurance Email Generator

[IF PDF MATERIALS UPLOADED - up to 3K characters:]
PRODUCT MATERIALS & SERVICES:
[Combined PDF content from all uploaded files]

PROSPECT RESEARCH & LEAD INFO:
- Company research for business model and insurance needs
- Lead details: name, title, company, industry, location, LinkedIn

TASK: Generate professional email with:
- Subject Line: 5-8 personalized words
- Email Body: 150-200 words with opening, value proposition, business case, social proof, CTA

GUIDELINES: Professional tone, specific product references, business value focus, no jargon
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

### Microsoft Graph Workflow (NEW)
```
1. Generate Leads → 2. Create OneDrive Excel → 3. Send Email Campaign → 
4. Track Email Opens → 5. Update Excel File → 6. Campaign Analytics
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

### User Experience (Enhanced v2.0)
- **Terminology**: "Web scraping" instead of "Apollo" for user-facing messages
- **PDF Upload**: Drag & drop interface with real-time file management
- **Product Materials**: Toggle to enable/disable enhanced AI emails
- **Microsoft 365 Integration**: OneDrive save and email campaign options
- **Email Campaign Builder**: Template editor with personalization variables
- **Progress**: Real-time status updates with elapsed time display (now up to 6 steps)
- **Error Handling**: User-friendly error messages with retry suggestions
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

## File Structure (Updated v2.0)
```
├── routes/
│   ├── apollo.js            # Apify integration & async job management
│   ├── leads.js             # Background job processing & OpenAI integration
│   ├── microsoft-graph.js   # NEW: OneDrive Excel integration
│   ├── email-automation.js  # NEW: Email campaigns & tracking
│   └── index.js             # Main router
├── middleware/
│   ├── rateLimiter.js       # Rate limiting with polling exemptions
│   └── graphAuth.js         # NEW: Microsoft Graph authentication
├── lead-generator.html      # Frontend interface (enhanced with MS365)
├── server.js                # Express server setup (updated CSP)
├── RENDER-SETUP.md          # NEW: Deployment guide for Azure & Render
├── IMPLEMENTATION-SUMMARY.md # NEW: Complete feature documentation
└── CLAUDE.md                # This file (project context)
```

## Development Workflow

### Testing (Extended v2.0)
```bash
# Core API Testing
GET /api/leads/test                  # Test OpenAI integration
GET /api/apollo/test                 # Test Apollo integration
GET /api/leads/jobs                  # List active jobs

# Microsoft Graph Testing (NEW)
GET /api/microsoft-graph/test        # Test Microsoft Graph connection
GET /api/microsoft-graph/onedrive/files # List OneDrive files

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

### v2.0 - Microsoft Graph API Integration (MAJOR UPDATE)
- **NEW: OneDrive Excel Integration**: Automatically save leads to Microsoft 365 Excel files with tracking columns
- **NEW: Email Automation System**: Send bulk personalized email campaigns through Microsoft Graph Mail API
- **NEW: Real-time Email Tracking**: Webhook notifications for email reads, replies with Excel sync
- **NEW: Campaign Analytics**: Detailed tracking dashboard for email campaign performance
- **Enhanced Background Jobs**: Extended workflow from 4 to up to 6 steps with MS365 integration
- **Updated Frontend**: Microsoft 365 integration section with email campaign builder
- **New Dependencies**: @microsoft/microsoft-graph-client, @azure/identity, @azure/msal-node, node-cron
- **Comprehensive Documentation**: RENDER-SETUP.md and IMPLEMENTATION-SUMMARY.md guides
- **Authentication Middleware**: Client Secret Credential flow for server-side operations
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