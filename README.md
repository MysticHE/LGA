# Lead Generation Automation

A complete lead generation tool that replicates your n8n workflow with Apollo.io scraping and AI-powered outreach generation.

## 🚀 Features

- **Interactive Form**: Job titles and company size selection
- **Apollo URL Generation**: Creates targeted Apollo.io search URLs
- **Lead Scraping**: Automated lead extraction via Apify Apollo scraper
- **AI Outreach**: Personalized outreach content generation using OpenAI
- **Excel Export**: Download leads with outreach content in Excel format
- **Progress Tracking**: Real-time workflow progress monitoring
- **Rate Limiting**: Built-in API protection and error handling

## 📋 Prerequisites

1. **Node.js** (v16 or higher)
2. **Apify API Token** - For Apollo.io lead scraping
3. **OpenAI API Key** - For AI outreach generation

## 🛠️ Installation

1. **Clone/Download the project**
```bash
cd LGA
```

2. **Install dependencies**
```bash
npm install
```

3. **Configure environment variables**
```bash
# Copy the example file
copy .env.example .env

# Edit .env with your API keys
APIFY_API_TOKEN=apify_api_YOUR_TOKEN_HERE
OPENAI_API_KEY=sk-YOUR_OPENAI_KEY_HERE
```

4. **Start the server**
```bash
# Development mode (auto-restart)
npm run dev

# Production mode
npm start
```

5. **Access the application**
   - Open http://localhost:3000 in your browser

## 🔧 API Keys Setup

### Apify API Token
1. Sign up at https://apify.com
2. Go to Settings → Integrations → API tokens
3. Create a new token and copy it to `.env`

### OpenAI API Key
1. Sign up at https://platform.openai.com
2. Go to API keys section
3. Create a new secret key and copy it to `.env`

## 📊 Usage

### Quick Start
1. Select job titles (Owner, Founder, C-suite, etc.)
2. Choose company sizes (1-10, 11-20, etc.)
3. Configure options:
   - **Generate AI Outreach**: Enable/disable personalized content
   - **Max Leads**: Set limit (10-500)

### Two Workflow Options

#### 1. Generate URL Only
- Creates Apollo.io search URL
- Copy/paste for manual use
- No API calls required

#### 2. Complete Workflow
- Generates Apollo URL
- Scrapes leads automatically
- Creates AI outreach content
- Downloads Excel file with all data

## 📁 Project Structure

```
LGA/
├── server.js              # Main Express server
├── package.json           # Dependencies & scripts
├── .env                   # Environment variables (create from .env.example)
├── lead-generator.html    # Frontend application
├── routes/
│   ├── apollo.js         # Apollo/Apify integration
│   └── leads.js          # Lead processing & OpenAI
├── middleware/
│   └── rateLimiter.js    # API rate limiting
└── README.md             # This file
```

## 🌐 API Endpoints

### Apollo Integration
- `POST /api/apollo/generate-url` - Generate Apollo search URL
- `POST /api/apollo/scrape-leads` - Scrape leads from Apollo (handles large datasets with session-based chunking)
- `POST /api/apollo/get-leads-chunk` - Get leads in chunks for large datasets
- `GET /api/apollo/test` - Test Apify connection

### Lead Processing
- `POST /api/leads/generate-outreach` - Generate AI outreach content
- `POST /api/leads/export-excel` - Export leads to Excel
- `POST /api/leads/start-workflow-job` - Start background workflow job (avoids timeout issues)
- `GET /api/leads/job-status/:jobId` - Check background job status
- `GET /api/leads/job-result/:jobId` - Get completed job results
- `GET /api/leads/jobs` - List all active jobs (debugging)
- `GET /api/leads/test` - Test OpenAI connection

### Health & Status
- `GET /health` - Server health check
- Rate limiting: 10 requests per minute per IP

## ⚙️ Configuration

### Environment Variables
```env
# Required
APIFY_API_TOKEN=your_apify_token
OPENAI_API_KEY=your_openai_key

# Optional
PORT=3000
NODE_ENV=development
MAX_REQUESTS_PER_MINUTE=10
MAX_LEADS_PER_REQUEST=500
```

### Rate Limiting
- Default: 10 requests per minute per IP
- Automatically scales based on load
- Customize via `MAX_REQUESTS_PER_MINUTE`

### Lead Limits
- Default max: 500 leads per request
- Configurable via `MAX_LEADS_PER_REQUEST`
- UI allows 10-500 range

## 🔍 Data Structure

### Lead Object
```json
{
  "name": "John Doe",
  "title": "CEO", 
  "organization_name": "Example Corp",
  "organization_website_url": "https://example.com",
  "estimated_num_employees": "51-100",
  "email": "john@example.com",
  "email_verified": "Y",
  "linkedin_url": "https://linkedin.com/in/johndoe",
  "industry": "Technology",
  "country": "Singapore",
  "notes": "AI-generated outreach content",
  "conversion_status": "Pending"
}
```

## 🛡️ Security Features

- **Helmet.js**: Security headers
- **CORS**: Configured origin restrictions
- **Rate Limiting**: DDoS protection
- **Input Validation**: Prevents injection attacks
- **Error Handling**: Sanitized error responses

## 🐛 Troubleshooting

### Common Issues

1. **"Configuration Error: Apify API token not configured"**
   - Add your Apify token to `.env` file
   - Restart the server

2. **"OpenAI API error"**
   - Verify OpenAI API key in `.env`
   - Check API key has sufficient credits

3. **"Rate limit exceeded"**
   - Wait 60 seconds and try again
   - Or adjust `MAX_REQUESTS_PER_MINUTE`

4. **"No leads found"**
   - Try different job title/company size combinations
   - Check Apollo.io URL manually

5. **Timeout Issues (500+ leads)**
   - The system now uses background jobs to avoid infrastructure timeouts
   - Large datasets are processed asynchronously with polling updates  
   - Jobs can run for up to 30 minutes without browser timeout
   - Check job progress with real-time status updates

### Debug Mode
```bash
NODE_ENV=development npm start
```
Enables detailed error messages and logging.

### Test Endpoints
- Visit `/api/apollo/test` to check Apify connection
- Visit `/api/leads/test` to check OpenAI connection

## 📈 Performance

- **Batch Processing**: Leads processed in groups of 5
- **Error Recovery**: Individual lead failures don't stop workflow
- **Timeouts**: 2-minute timeout for scraping operations
- **Memory**: Optimized for large lead datasets

## 🔄 Workflow Comparison

### Original n8n Workflow
1. Form submission → Apollo URL generation
2. Apify scraper → Lead data extraction  
3. OpenAI → Personalized outreach generation
4. Google Sheets → Data storage

### This Implementation
1. Form submission → Apollo URL generation ✅
2. Apify scraper → Lead data extraction ✅
3. OpenAI → Personalized outreach generation ✅
4. Excel download → Local data storage ✅

**Result**: 100% feature parity with enhanced UI and offline capability!

## 📞 Support

For issues or questions:
1. Check this README
2. Test API connections via `/api/*/test` endpoints
3. Review server logs for error details
4. Verify environment variable configuration