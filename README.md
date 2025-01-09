# AI-Powered Blog Content Audit System

## Overview
This system performs automated analysis and categorization of blog content using AI to evaluate brand alignment, content quality, SEO optimization, and content categorization. It's specifically designed for auditing Act-On's blog content archive, with a focus on evaluating and categorizing content according to brand guidelines, topic relevance, and current marketing strategies.

## Key Features
- Content quality and brand alignment analysis
- Tone and voice evaluation based on brand spectrum
- SEO analysis and optimization recommendations
- Multimedia asset assessment
- Content categorization and tagging
- Personal pronoun detection
- Performance metrics integration
- Cost tracking for API usage

## System Architecture

### Core Components
1. **Blog Scraper** (`scrape_blog.py`)
   - Extracts content from blog posts
   - Captures metadata, images, and structural elements
   - Handles multimedia assessment
   - Detects outdated widgets and elements

2. **AI Analysis** (`ai_analysis.py`)
   - Interfaces with Claude API
   - Performs content categorization
   - Analyzes tone and voice
   - Evaluates SEO elements
   - Assesses quality and brand fit

3. **Analysis Processor** (`make_analysis.py`)
   - Processes scraped content
   - Generates comprehensive Excel reports
   - Applies conditional formatting
   - Integrates performance metrics

## Setup Instructions

### Prerequisites
- Python 3.8+
- pip package manager
- Anthropic API key
- Excel for viewing reports

### Environment Setup
1. Clone the repository
2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # Unix/macOS
venv\Scripts\activate     # Windows
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

4. Create a `.env` file with your API key:
```
ANTHROPIC_API_KEY=your_api_key_here
```

### Directory Structure
```
project/
├── ai_analysis.py
├── scrape_blog.py
├── make_analysis.py
├── resources/
│   ├── yoast-blog-keywords.xlsx
│   └── performance.xlsx
├── output/
│   ├── blog.json
│   └── blog_audit.xlsx
└── .env
```

## Usage

### 1. Scraping Blog Content
Run the scraper to collect blog content:
```bash
python scrape_blog.py
```
This will:
- Scrape specified blog URLs
- Extract content and metadata
- Save raw data to `output/blog.json`

### 2. Running Analysis
Process the scraped content in batches:
```bash
python make_analysis.py --start_index 0 --batch_size 100
```
Parameters:
- `--start_index`: Starting position in the blog list
- `--batch_size`: Number of articles to process in this batch

### 3. Output
The system generates an Excel report (`blog_audit.xlsx`) with:

#### Basic Information
- Title and URL
- Publication/modified dates
- Word count
- Reading level (Gunning Fog)

#### Quality & Brand Fit
- Overall quality score
- Topic relevance
- Brand alignment
- Quality recommendations

#### Tone & Voice
- Challenger/supportive balance
- Natural/conversational score
- Authentic/approachable score
- Gender-neutral/inclusive score
- Personal pronoun detection

#### SEO Analysis
- Target keyword analysis
- Keyword density
- Meta description quality
- Header tag structure
- Optimization recommendations

#### Multimedia Assessment
- Image count and dimensions
- Header image analysis
- Outdated widget detection

#### Content Categorization
- Primary category
- Solution topic
- Use case
- Customer journey stage
- CMO priority
- Marketing activity type
- Target audience

#### Performance Metrics
- Views and users
- Session data
- Engagement metrics
- API cost tracking

## Scoring Guidelines

### Word Count
- Green: 1000-1200 words
- Yellow: 800-999 or 1201-1400 words
- Orange/Red: <800 or >1400 words

### Reading Level (Gunning Fog)
- Green: Grade 9-12
- Yellow: Grade 13-15 or 7-8
- Red: Grade 16+ or <7

### Quality Scores
All quality metrics use a 0-100 scale:
- 90-100: Exceptional
- 80-89: Strong
- 70-79: Good
- 60-69: Needs Improvement
- Below 60: Requires Major Revision
- Below 40: Requires Complete Rewrite

## Maintenance

### Adding New Content Categories
To add new content categories or modify existing ones:
1. Update the categorization logic in `ai_analysis.py`
2. Modify the Excel template in `make_analysis.py`
3. Re-run analysis on affected content

### Updating Brand Guidelines
To update brand voice or content guidelines:
1. Modify the respective prompt templates in `ai_analysis.py`
2. Update scoring criteria if needed
3. Consider re-running analysis on existing content

## Troubleshooting

### Common Issues
1. **API Rate Limiting**
   - Use batch processing
   - Implement retry logic
   - Monitor API costs

2. **Memory Issues**
   - Process content in smaller batches
   - Clear memory between batches
   - Monitor system resources

3. **Excel File Corruption**
   - Always maintain backups
   - Save incremental versions
   - Use error handling when writing files

## Notes
- The system is designed for Act-On's specific needs but can be adapted
- Regular monitoring of API costs is recommended
- Consider periodic revalidation of AI analysis accuracy
- Keep brand guidelines and scoring criteria updated