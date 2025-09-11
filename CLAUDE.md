# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

### Development Commands
- `npm run dev` - Start development server with nodemon auto-restart
- `npm start` - Start production server 
- `node server.js` - Direct server start (same as npm start)

### Testing API Endpoints
- Visit `http://localhost:3000/api/apollo/test` - Test Apify connection
- Visit `http://localhost:3000/api/leads/test` - Test OpenAI connection  
- Visit `http://localhost:3000/health` - Server health check

### File Processing
- No build step required (server serves static HTML directly)
- Excel files are processed in-memory using XLSX library
- PDF files are processed using pdf-parse for content extraction

## Architecture Overview

This is a lead generation automation tool that combines Apollo.io scraping with AI-powered email outreach and Microsoft Graph integration for email campaigns.

### Core Components

**Server Architecture (server.js)**
- Express.js server with security middleware (Helmet, CORS)
- Process singleton pattern prevents multiple server instances
- Modular route-based architecture with specialized endpoints

**Route Modules**
- `routes/apollo.js` - Apollo.io/Apify integration for lead scraping
- `routes/leads.js` - Lead processing and OpenAI content generation  
- `routes/email-automation.js` - Email campaign management with Excel integration
- `routes/microsoft-graph.js` - Microsoft Graph API integration
- `routes/email-*.js` - Various email-related services (templates, tracking, scheduling, bounce detection)
- `routes/auth.js` - Authentication handling
- `routes/campaign-status.js` - Campaign monitoring

**Utility Modules (utils/)**
- `excelProcessor.js` - Excel file manipulation and processing
- `excelGraphAPI.js` - Direct Excel updates via Microsoft Graph API
- `emailContentProcessor.js` - Email content generation and processing  
- `campaignLockManager.js` - Prevents concurrent campaign conflicts
- `campaignTokenManager.js` - Manages campaign authentication tokens
- `excelDuplicateChecker.js` - Prevents duplicate lead processing
- `processSingleton.js` - Ensures single server instance

### Data Flow Architecture

1. **Lead Generation**: Form submission → Apollo URL generation → Apify scraper → Lead extraction
2. **Content Processing**: Lead data → OpenAI API → AI-generated outreach content
3. **Email Campaigns**: Excel upload → Lead processing → Microsoft Graph → Email sending → Tracking updates
4. **Tracking**: Read receipts (pixel tracking) + Reply monitoring (cron job) → Excel updates via Graph API

### Key Integration Points

**External APIs**
- Apify API for Apollo.io lead scraping
- OpenAI API for AI content generation
- Microsoft Graph API for email sending and Excel integration

**Authentication**
- Delegated authentication for Microsoft Graph (in `middleware/delegatedGraphAuth.js`)
- API key-based authentication for external services

**File Processing**
- Excel files processed with XLSX library and updated via Graph API
- PDF content extraction for reference material processing
- Multer for file upload handling with 10MB limits

## Environment Configuration

Required environment variables (see `.env.example`):
- `APIFY_API_TOKEN` - Apollo.io scraping via Apify
- `OPENAI_API_KEY` - AI content generation  
- `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` - Microsoft Graph integration
- `RENDER_EXTERNAL_URL` - Deployment URL for tracking callbacks

## Important Implementation Notes

### Process Management
- Server uses singleton pattern to prevent multiple instances
- Campaign lock manager prevents concurrent email sending conflicts  
- Background job processing for large datasets to avoid timeouts

### Excel Integration Pattern
- Direct Graph API updates for real-time Excel modifications
- Duplicate checking before processing leads
- Batch processing with error recovery for large datasets

### Email Campaign Architecture
- Template-based email generation with AI personalization
- Tracking via pixel images and Microsoft Graph inbox monitoring
- Rate limiting and bounce detection for deliverability

### Error Handling
- Graceful degradation when optional services are unavailable
- Comprehensive API testing endpoints for troubleshooting
- Detailed error logging for campaign and processing issues