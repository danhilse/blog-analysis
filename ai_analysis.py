# ai_analysis.py
import os
import json
import anthropic
from dotenv import load_dotenv
from decimal import Decimal

# Load environment variables
load_dotenv()
client = anthropic.Anthropic(api_key=os.getenv('ANTHROPIC_API_KEY'))

INPUT_COST_PER_MTOK = Decimal('3.00')
OUTPUT_COST_PER_MTOK = Decimal('15.00')
class CostTracker:
    def __init__(self):
        self.cost = Decimal('0.00')
        self.input_tokens = 0
        self.output_tokens = 0

    def reset(self):
        self.__init__()

    def add_usage(self, input_tokens, output_tokens):
        input_cost = (Decimal(input_tokens) / Decimal('1000000')) * Decimal('3.00')
        output_cost = (Decimal(output_tokens) / Decimal('1000000')) * Decimal('15.00')
        self.cost += input_cost + output_cost
        self.input_tokens += input_tokens
        self.output_tokens += output_tokens

cost_tracker = CostTracker()

def make_api_call(formatted_prompt, system_prompt="You are an expert content analyst."):
    try:
        token_count = client.beta.messages.count_tokens(
            betas=["token-counting-2024-11-01"],
            model="claude-3-sonnet-20240229",
            system=system_prompt,
            messages=[{"role": "user", "content": formatted_prompt}]
        )
        input_tokens = token_count.input_tokens

        response = client.messages.create(
            model="claude-3-sonnet-20240229",
            max_tokens=1500,
            temperature=0.3,
            system=system_prompt,
            messages=[{"role": "user", "content": formatted_prompt}]
        )

        cost_tracker.add_usage(input_tokens, response.usage.output_tokens)
        return json.loads(response.content[0].text)
    except Exception as e:
        print(f"API call error: {e}")
        return None

def analyze_quality_brand_fit(guidelines):
    prompt = f"""Analyze this content for quality and brand alignment based on these guidelines:

{guidelines}

Return ONLY a JSON object with these exact fields:
{{
    "Overall Quality Score": <0-100 integer>,
    "Brand Alignment Score": <0-100 integer>,
    "Quality Notes/Recommendations": <string with key observations and actionable recommendations>
}}"""
    return make_api_call(prompt)

def analyze_tone_voice(guidelines):
    prompt = f"""Analyze this content's tone and voice based on these guidelines:

{guidelines}

Return ONLY a JSON object with these exact fields:
{{
    "Challenger Percentage": <0-100 integer>,
    "Supportive Percentage": <0-100 integer>,
    "Natural/Conversational Score": <0-100 integer>,
    "Authentic/Approachable Score": <0-100 integer>,
    "Gender-Neutral/Inclusive Score": <0-100 integer>,
    "Tone Notes/Recommendations": <string with observations and suggestions>
}}"""
    return make_api_call(prompt)

def analyze_seo(seo_data):
    prompt = f"""Analyze this content's SEO effectiveness using the provided metadata:

{seo_data}

Return ONLY a JSON object with these exact fields:
{{
    "Keyword Density": <float percentage>,
    "Keyword Integration Score": <0-100 integer>,
    "Meta Description Quality Score": <0-100 integer>,
    "Recommended New Keywords": <array of 3-5 keyword strings>,
    "SEO Notes/Recommendations": <string with actionable improvements>
}}"""
    return make_api_call(prompt)

def analyze_content_categorization(content):
    prompt = f"""You are an expert marketing analyst. Your task is to analyze the given content and determine the most appropriate categories based on the content's focus, themes, and target audience. Return ONLY a valid JSON object.

<Content to analyze> 
{content}
</Content to analyze> 

First, identify the most relevant Use Case based on the content's main focus and detailed descriptions:

GET Stage Use Cases:
- "Identify and Target Audience Segments" - Content about capturing email addresses, first-party data collection, progressive profiling, and landing page optimization
- "Reach New Prospects" - Content about behavioral insights, firmographic data, and customer lifecycle segmentation
- "Personalize Outreach" - Content about automated programs, targeted emails based on behavior, dynamic segmentation, and CRM integration
- "Nurture Prospects" - Content about targeted email programs, thought leadership, and sales funnel progression
- "Deliver Best Leads to Sales" - Content about lead scoring, sales-marketing alignment, and lead quality optimization
- "Empower Sales Intelligence" - Content about ABM insights, behavioral data capture, and sales workflow automation
- "Scale Operations" - Content about CRM integrations, prospect targeting, and automated marketing workflows

KEEP Stage Use Cases:
- "Welcome and Onboard" - Content about automated tasks, behavioral engagement data, and omnichannel marketing programs
- "Drive Product Adoption" - Content about automated welcome series, customer onboarding, and direct mail integration
- "Regular Communication" - Content about transactional emails, brand consistency, email performance, and compliance
- "Automate Renewal" - Content about social media automation, customer re-engagement, and milestone-based communications

GROW Stage Use Cases:
- "Grow Advocates" - Content about automated feedback collection, community building, and customer education
- "Automate Communications" - Content about internal workflows, partner communications, and automated messaging
- "Cross-sell and Upsell" - Content about targeted offers, behavioral insights, loyalty programs, and customer value expansion
- "Marketing Performance" - Content about ROI optimization and marketing effectiveness

OPTIMIZE Stage Use Cases:
- "Data-Driven Marketing" - Content about automation tools, unified customer views, and personalized strategies
- "Scale Marketing Output" - Content about multi-channel campaign coordination, lead nurturing, and conversion tracking
- "Single Source of Truth" - Content about centralized databases, CRM synchronization, and lead scoring systems
- "Marketing/Sales Insights" - Content about integrated reporting and performance analytics

**IF NO USE CASE MATCHES, SELECT "NONE"**
- "No Clear Match" - Content that doesn't fit any specific Use Case

Then, the CMO Priority must match the Use Case's stage:
- GET → "New Customer Acquisition" or "Build Pipeline and Accelerate Sales"
- KEEP → "Deliver Value and Keep Customers"
- GROW → "Improve Brand Loyalty" or "Maximize ARPU"
- OPTIMIZE → "Maximizing MROI"
- NONE → "No Clear Match"

Additional required categorization:

Primary Category - Choose based on content type:
- "Product" - Content focusing on Act-On features/capabilities
- "Industry" - Content about industry trends/challenges
- "Use Case" - Content demonstrating specific applications
- "Thought Leadership" - Educational/strategic content
- "No Clear Match" - Content that doesn't fit any specific category

Solution Topic - Choose based on primary solution discussed:
- "Convert Unknown Visitors to Known Leads" - Website visitor identification
- "Identify and Target Audience Segments" - Audience segmentation
- "Reach New Prospects Through Omni-channel Campaigns" - Multi-channel outreach
- "Personalize Outreach and Communication" - Personalization
- "Scale Demand Generation Operations" - Operational scaling
- "No CLear Topic" - Content that doesn't fit any specific topic

Marketing Activity Type - Choose based on main marketing activity:
- "Email Marketing" - Email campaigns/automation
- "Social Media Marketing" - Social media activities
- "Content Marketing" - Content creation/distribution
- "Lead Generation" - Lead capture/qualification
- "Account-Based Marketing" - ABM strategies
- "Marketing Automation" - Automation processes
- "Analytics and Reporting" - Data analysis
- "Website Optimization" - Website improvements
- "Event Marketing" - Event management
- "Customer Marketing" - Customer-focused campaigns
- "No Clear Activity Type" - Content that doesn't fit any specific activity

Target Audience - Choose based on content's intended reader:
- "Marketing Leaders" - Strategic/executive content
- "Demand Generation Managers" - Demand gen focused
- "Marketing Operations Managers" - Operations focused
- "Digital Marketing Managers" - Digital marketing focused
- "Marketing Automation Specialists" - Technical/platform focused
- "Sales Leaders" - Sales-aligned content
- "Small Business Owners" - SMB focused
- "Enterprise Marketers" - Enterprise focused
- "No Clear Audience" - Content that doesn't fit any specific audience

Return ONLY a JSON object with these exact fields:
{{
    "Primary Category": <selected value>,
    "Solution Topic": <selected value>,
    "Use Case": <selected value>,
    "Customer Journey Stage": <selected value based on Use Case>,
    "CMO Priority": <selected value matching Journey Stage>,
    "Marketing Activity Type": <selected value>,
    "Target Audience": <selected value>
}}"""

    return make_api_call(prompt)

