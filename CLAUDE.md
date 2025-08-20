# Lead Generation Automation (LGA) - Claude Context

## Project Overview

A complete lead generation tool that automates Apollo.io scraping with AI-powered outreach generation. Features background job processing to handle large datasets without timeout limitations.

## Core Architecture

### Background Job System
- **Purpose**: Avoid 5-minute infrastructure timeouts for large scraping operations
- **Flow**: Start job → Poll status → Retrieve results when complete
- **Timeout**: No time limits (removed 30-minute restriction for unlimited datasets)

### Apollo Integration (Apify-based)
- **Sync API Issue**: Apify's `/run-sync-get-dataset-items` has 300-second (5-minute) hard timeout
- **Solution**: Use async API (`/runs` endpoint) with polling mechanism
- **Dataset Retrieval**: Multi-method fallback system for robust data access

## Environment Setup

### Required Environment Variables
```bash
# Core API Keys
APIFY_API_TOKEN=your_apify_token_here
OPENAI_API_KEY=your_openai_key_here

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
  "rate-limiter-flexible": "^2.4.0"
}
```

## Key Features

### 1. Unlimited Lead Scraping
- **Default**: 0 (unlimited) - removed 2000 lead limit
- **Frontend**: Input allows 0-10,000 leads
- **Backend**: Handles `recordLimit = 0` for unlimited processing

### 2. Async Job Processing
- **Job Creation**: Immediate response with job ID
- **Status Polling**: Real-time progress tracking
- **Result Retrieval**: Separate endpoint for large datasets

### 3. Error Handling
- **Apify Timeouts**: Automatic retry with exponential backoff
- **Dataset 404s**: Multi-method retrieval fallback
- **Rate Limiting**: Exempted polling endpoints to prevent 429 errors
- **Network Issues**: Retry logic with graceful degradation

## API Endpoints

### Core Workflow
```
POST /api/leads/start-workflow-job
├── POST /api/apollo/generate-url
├── POST /api/apollo/start-scrape-job
├── GET /api/apollo/job-status/{jobId}
├── GET /api/apollo/job-result/{jobId}
└── POST /api/leads/generate-outreach
```

### Job Management
```
GET /api/leads/job-status/{jobId}     # Poll job progress
GET /api/leads/job-result/{jobId}     # Get completed results
GET /api/leads/jobs                   # List all active jobs
```

## Technical Implementation

### Apollo URL Generation
- **Location**: Singapore (hardcoded)
- **Filters**: Email verified, company domain exists
- **Parameters**: Job titles, company sizes
- **Output**: Apollo.io search URL with encoded parameters

### Scraping Process
1. **Generate Apollo URL** with job titles and company size filters
2. **Start Apify scraper** with async API call
3. **Poll Apify run status** every 5 seconds until completion
4. **Retrieve dataset** using multi-method fallback approach
5. **Transform data** to standardized lead format
6. **Generate AI outreach** (optional, requires OpenAI API)

### Data Flow
```
Frontend Form → Apollo URL → Apify Scraper → Raw Leads → 
Transform → AI Outreach → Final Results → Excel Export
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

### User Experience
- **Terminology**: "Web scraping" instead of "Apollo" for user-facing messages
- **Progress**: Real-time status updates with elapsed time display
- **Error Handling**: User-friendly error messages with retry suggestions
- **Results**: Downloadable Excel export with structured data

### Polling Strategy
- **Interval**: 3 seconds with exponential backoff on rate limits
- **Timeout**: No frontend timeout (matches unlimited backend)
- **Error Recovery**: Automatic retry with user notification

## File Structure
```
├── routes/
│   ├── apollo.js          # Apify integration & async job management
│   ├── leads.js           # Background job processing & OpenAI integration
│   └── index.js           # Main router
├── middleware/
│   └── rateLimiter.js     # Rate limiting with polling exemptions
├── lead-generator.html    # Frontend interface
└── server.js              # Express server setup
```

## Development Workflow

### Testing
```bash
# Test OpenAI integration
GET /api/leads/test

# Test Apollo integration  
GET /api/apollo/test

# List active jobs
GET /api/leads/jobs
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