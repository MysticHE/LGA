# Quick Start Guide

## 🚀 Two Ways to Use the Lead Generator

### Option 1: URL Generation Only (No Server Required)
Perfect for quick Apollo URL generation without any setup.

1. **Open the HTML file directly in your browser:**
   ```
   Double-click: lead-generator.html
   ```

2. **Select multiple options:**
   - Job Titles: Hold Ctrl/Cmd and click multiple (e.g., Founder, CEO, C suite)
   - Company Sizes: Hold Ctrl/Cmd and click multiple (e.g., 1-10, 11-20, 21-50)

3. **Click "🔗 Generate URL Only"**
   - Works instantly without any server
   - Generates Apollo.io search URL with all your filters
   - Copy URL and paste into Apollo.io

### Option 2: Complete Automation (Server Required)
Full workflow with lead scraping and AI outreach generation.

1. **Start the server:**
   ```bash
   cd LGA
   npm install
   npm start
   ```

2. **Open in browser:**
   ```
   http://localhost:3000
   ```

3. **Select multiple options and click "🚀 Complete Workflow"**
   - Generates Apollo URL
   - Scrapes leads automatically
   - Creates AI outreach content  
   - Downloads Excel file

## 🔧 Error Handling

### "Failed to fetch" Error
This happens when you try to use the complete workflow but the server isn't running.

**Solution:**
- Either use "🔗 Generate URL Only" (works without server)
- OR start the server with `npm start`

**Automatic Fallback:**
The system will automatically generate the Apollo URL for you even when the server is down, with a helpful message explaining how to get full functionality.

## 🎯 Multiple Selection Examples

### Example 1: Startup Founders
- **Job Titles**: Founder, Owner, CEO
- **Company Sizes**: 1-10, 11-20
- **Result**: Finds startup founders and CEOs at small companies

### Example 2: Mid-Level Managers
- **Job Titles**: Manager, Director, Head
- **Company Sizes**: 51-100, 101-200, 201-500
- **Result**: Finds managers and directors at medium-sized companies

### Example 3: Enterprise Executives
- **Job Titles**: C suite, VP, Partner
- **Company Sizes**: 1001-2000, 2001-5000, 5001-10000
- **Result**: Finds executives at large enterprises

## 📊 What You Get

### With URL Only:
- Apollo.io search URL with all your filters
- Manual copy/paste into Apollo
- Instant results, no setup

### With Complete Workflow:
- Automated lead scraping (up to 2000 leads)
- AI-generated personalized outreach content
- Excel file download with all data
- Progress tracking and statistics

Both options use the exact same filtering logic as your n8n workflow!