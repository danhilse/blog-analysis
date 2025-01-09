# ai_analysis.py
import os
import json
import anthropic
from dotenv import load_dotenv
from decimal import Decimal
import time
import re
import traceback

from static import brandGuidelines, voiceGuidelines

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


def fix_json_quotes(json_str: str) -> str:
    """
    Fix JSON string by properly escaping quotes within value strings.
    Uses a simple regex-based approach to find and fix quotes inside JSON string values.
    """

    def escape_value_quotes(match):
        """Helper function to escape quotes within a JSON string value."""
        value = match.group(2)
        # Escape any unescaped quotes within the value
        escaped_value = value.replace('"', '\\"')
        return f'{match.group(1)}"{escaped_value}"'

    # First convert any fancy quotes to regular quotes
    json_str = json_str.replace('"', '"').replace('"', '"')

    # Regex pattern to match JSON string values, considering the key and colon
    pattern = r'("[\w\s]+"\s*:\s*)"(.+?)"(?=\s*[,}])'

    # Replace each string value with properly escaped version
    fixed = re.sub(pattern, escape_value_quotes, json_str)
    return fixed


def make_api_call(formatted_prompt, system_prompt="You are an expert content analyst.", max_retries=2):
    retries = 0

    while retries <= max_retries:
        try:
            # Count tokens
            token_count = client.beta.messages.count_tokens(
                betas=["token-counting-2024-11-01"],
                model="claude-3-sonnet-20240229",
                system=system_prompt + " Return ONLY valid JSON.",
                messages=[{"role": "user", "content": formatted_prompt}]
            )
            input_tokens = token_count.input_tokens

            # Make API call
            response = client.messages.create(
                model="claude-3-sonnet-20240229",
                max_tokens=1500,
                temperature=0.3,
                system=system_prompt + " Return ONLY valid JSON.",
                messages=[{"role": "user", "content": formatted_prompt}]
            )

            response_text = response.content[0].text.strip()

            # Remove any markdown
            if response_text.startswith("```"):
                response_text = re.sub(r'```\w*\n?', '', response_text)
                response_text = response_text.strip()

            print("\nRaw response:")
            print(response_text)

            # Fix quotes and try to parse
            fixed_json = fix_json_quotes(response_text)
            print("\nFixed JSON:")
            print(fixed_json)

            try:
                parsed_response = json.loads(fixed_json)
                cost_tracker.add_usage(input_tokens, response.usage.output_tokens)
                return parsed_response
            except json.JSONDecodeError as je:
                print(f"\nJSON Parse Error: {str(je)}")
                print("Attempting manual quote escape...")
                # Try one more time with manual escaping
                fixed_json = response_text.replace('"White-Collar Mechanic"', '\\"White-Collar Mechanic\\"')
                fixed_json = fixed_json.replace('"Supportive Challenger"', '\\"Supportive Challenger\\"')
                try:
                    parsed_response = json.loads(fixed_json)
                    cost_tracker.add_usage(input_tokens, response.usage.output_tokens)
                    return parsed_response
                except:
                    retries += 1
                    if retries > max_retries:
                        raise
                    time.sleep(2)

        except Exception as e:
            retries += 1
            print(f"\nAttempt {retries} failed with error: {str(e)}")
            if retries > max_retries:
                print(f"Failed after {max_retries} retries. Last error: {str(e)}")
                return None
            time.sleep(2)

    return None


def analyze_quality_brand_fit(content):
    default_response = {
        "Overall Quality Score": "n/a",
        "Topic Relevance": "Error",
        "Brand Alignment": "Error",
        "Quality Notes": "Error analyzing content quality",
        "Brand Alignment Notes": "Error analyzing brand alignment"
    }

    try:
        prompt = f"""You are an expert content evaluator for Act-On. Analyze this content's quality and brand alignment.

<Content to analyze>
{content}
</Content to analyze>

{brandGuidelines}

<Company Focus>
Act-On is a B2B marketing automation platform helping marketers create and optimize multi-channel marketing campaigns. Core topics include: marketing automation, email marketing, lead management, B2B marketing strategies, campaign optimization, marketing analytics, CRM integration, lead generation, customer engagement, and marketing technology for revenue growth.
</Company Focus>

Score these quality factors (0-100):
1. Writing Excellence (clear communication, grammar, structure)
2. Structure & Organization (logical flow, clear hierarchy)
3. Value & Impact (audience focus, actionable insights)
4. Engagement (compelling narrative, examples, CTAs)
5. Topic Relevance (connection to marketing automation/B2B marketing)

Evaluate brand alignment based on:
- Dual personality as "Supportive Challenger" and "White-Collar Mechanic"
- Natural, conversational, direct voice
- Core messaging: agile marketing, innovation, partnership
- Brand values: people first, authenticity, excellence, continuous improvement

**BE CRITICAL - USE FULL SCORING RANGE**
"Score Reference Points": 
{{
  "90-100": "Exceptional content that excels in all areas",
  "80-89": "Strong content with minor issues",
  "70-79": "Solid content needing some improvements",
  "60-69": "Passable content with significant issues",
  "Below 60": "Poor quality content needing major revision",
  "Below 40": "Unacceptable content requiring complete rewrite"}}

Return this exact JSON - no other text/formatting:
{{
    "Overall Quality Score": <integer 0-100>,
    "Topic Relevance": <"On Topic" or "Tangentially Related" or "Off Topic">,
    "Brand Alignment": <"On Brand" or "Mostly on Brand" or "Needs Work" or "Not on Brand">,
    "Quality Notes": "<2-3 sentences, no line breaks>",
    "Brand Alignment Notes": "<2-3 sentences, no line breaks>"
}}"""

        print("\nSending analysis request...")
        response = make_api_call(prompt)

        if not response:
            print("No response from API")
            return default_response

        print("\nValidating response structure...")

        # Validate individual fields
        try:
            quality_score = int(response.get("Overall Quality Score", 0))
            if not 0 <= quality_score <= 100:
                print(f"Invalid quality score: {quality_score}")
                raise ValueError("Quality score must be between 0 and 100")

            topic_relevance = str(response.get("Topic Relevance", ""))
            if topic_relevance not in ["On Topic", "Tangentially Related", "Off Topic"]:
                print(f"Invalid topic relevance: {topic_relevance}")
                raise ValueError("Invalid topic relevance value")

            brand_alignment = str(response.get("Brand Alignment", ""))
            if brand_alignment not in ["On Brand", "Mostly on Brand", "Needs Work", "Not on Brand"]:
                print(f"Invalid brand alignment: {brand_alignment}")
                raise ValueError("Invalid brand alignment value")

        except (ValueError, TypeError) as e:
            print(f"Validation error: {str(e)}")
            print("Raw response:", response)
            return default_response

        # Construct validated result
        result = {
            "Overall Quality Score": quality_score,
            "Topic Relevance": topic_relevance,
            "Brand Alignment": brand_alignment,
            "Quality Notes": str(response.get("Quality Notes", "No quality notes provided")).strip(),
            "Brand Alignment Notes": str(
                response.get("Brand Alignment Notes", "No brand alignment notes provided")).strip()
        }

        print("\nSuccessfully validated response")
        return result

    except Exception as e:
        print(f"\nError in quality analysis: {str(e)}")
        print("Traceback:", traceback.format_exc())
        return default_response

def analyze_tone_voice(content):
    prompt = f"""ou are a brand voice expert analyzing content against Act-On's guidelines. Score strictly using these bands:
- Excellent (90-100): Nearly perfect alignment with guidelines, minimal/no improvements needed
- Good (75-89): Strong alignment, few minor improvements possible  
- Adequate (60-74): Meets basic requirements but needs several improvements
- Needs Work (40-59): Significant misalignment with guidelines
- Poor (0-39): Fails to meet most guideline requirements

Critical Score Caps:
- Use of prohibited terms/jargon caps score at 60
- >3 instances of corporate-speak caps at 70
- Incorrect challenger/supportive balance for channel caps at 65

Analyze against Act-On's voice guidelines:

{voiceGuidelines}

<Content to analyze>
{content}
</Content to analyze>

Analyze the following elements:

1. Challenger vs Supportive Balance
- Calculate the ratio of challenging content (pushing readers, questioning status quo) vs supportive content (guidance, reassurance)
- Consider the channel-appropriate balance according to Act-On's Tone of Voice Spectrum
- Evaluate if the balance matches the content type and purpose

2. Natural/Conversational Quality
- Assess how well the content maintains a straightforward, plain-speaking tone
- Check for corporate-speak, jargon, or overly complex language
- Evaluate the flow and readability of the content

3. Authentic/Approachable Quality
- Look for confidence without arrogance
- Assess professional yet accessible language
- Evaluate the balance of technical expertise and approachability

4. Gender-Neutral/Inclusive Language
- Check for any exclusionary terms or phrases
- Assess overall inclusivity of language
- Evaluate use of gender-neutral pronouns and terminology

Return EXACTLY this JSON structure with no additional text, markdown, or formatting:
{{
    "Challenger Percentage": <integer between 0 and 100>,
    "Supportive Percentage": <integer between 0 and 100>,
    "Natural/Conversational Score": <integer between 0 and 100>,
    "Authentic/Approachable Score": <integer between 0 and 100>,
    "Gender-Neutral/Inclusive Score": <integer between 0 and 100>,
    "Tone Notes/Recommendations": "<exactly 2-3 sentences>"
}}

IMPORTANT:
- Return only valid JSON - no explanations, prefaces, or additional formatting
- Challenger and Supportive Percentages must sum to exactly 100
- All scores must be integers, not floats
- Do not use line breaks within the text fields
- Escape any quotes within text fields using \"
- Ensure all field names match exactly as shown
- Text fields must be wrapped in double quotes

Note: Challenger and Supportive Percentages must total 100%. Scores should reflect how well the content meets each criterion, with 100 being perfect alignment and 0 being complete misalignment."""

    return make_api_call(prompt)

def analyze_seo(content, seo_data):
    """
    Analyzes content for SEO effectiveness using provided metadata.
    """
    prompt = f"""You are an expert SEO analyst. Evaluate this content and metadata for SEO optimization.

<Content>
{content}
</Content>

<SEO Metadata>
Target Keyword: {seo_data['current_target_keyword']}
Meta Description: {'Present' if seo_data['meta_description_present'] else 'Missing'}
H1 Tag: {'Present' if seo_data['h1_present'] else 'Missing'}
H2 Tags: {seo_data['h2_count']}
H3 Tags: {seo_data['h3_count']}
</SEO Metadata>

Analyze these elements and provide scoring:

1. Keyword Analysis
- Calculate exact keyword density
- Evaluate keyword placement and distribution
- Check for keyword stuffing
- Assess semantic relevance

2. Structure Analysis
- Review header hierarchy
- Evaluate content organization
- Check keyword usage in headers

3. Meta Description Review
- Evaluate presence, length, and effectiveness
- Check keyword inclusion and CTA
- Assess value proposition

4. Keyword Opportunities
- Identify related semantic keywords
- Consider search intent
- Look for topic expansions

Return EXACTLY this JSON structure with no additional text, markdown, or formatting:
{{
    "Keyword Density": <number formatted to exactly 2 decimal places>,
    "Keyword Integration Score": <integer between 0 and 100>,
    "Meta Description Quality Score": <integer between 0 and 100>,
    "Recommended New Keywords": ["keyword1", "keyword2", "keyword3"],
    "SEO Notes/Recommendations": "<exactly 2-3 specific recommendations>"
}}

IMPORTANT:
- Return only valid JSON - no explanations, prefaces, or additional formatting
- Keyword Density must be formatted as X.XX (exactly 2 decimal places)
- All scores must be integers, not floats
- Array elements must be wrapped in double quotes
- Do not use line breaks within the text fields
- Escape any quotes within text fields using \"
- Ensure all field names match exactly as shown
- Text fields must be wrapped in double quotes"""

    default_response = {
        "Keyword Density": 0.00,
        "Keyword Integration Score": 0,
        "Meta Description Quality Score": 0,
        "Recommended New Keywords": [],
        "SEO Notes/Recommendations": "Error analyzing SEO content"
    }

    try:
        # Get the response from the API
        response = make_api_call(prompt)

        # If response is None or empty, return default
        if not response:
            return default_response

        # Response is already parsed JSON from make_api_call, no need to parse again
        result = {
            "Keyword Density": float(response.get("Keyword Density", 0)),
            "Keyword Integration Score": int(response.get("Keyword Integration Score", 0)),
            "Meta Description Quality Score": int(response.get("Meta Description Quality Score", 0)),
            "Recommended New Keywords": response.get("Recommended New Keywords", []),
            "SEO Notes/Recommendations": str(response.get("SEO Notes/Recommendations", "No recommendations available"))
        }

        return result

    except Exception as e:
        print(f"Error in SEO analysis: {str(e)}")
        return default_response

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

Return EXACTLY this JSON structure with no additional text, markdown, or formatting:
{{
    "Primary Category": "<exactly one of the specified category values>",
    "Solution Topic": "<exactly one of the specified topic values>",
    "Use Case": "<exactly one of the specified use case values>",
    "Customer Journey Stage": "<exactly one of: GET, KEEP, GROW, OPTIMIZE, NONE>",
    "CMO Priority": "<exactly one of the specified priority values>",
    "Marketing Activity Type": "<exactly one of the specified activity type values>",
    "Target Audience": "<exactly one of the specified audience values>"
}}

IMPORTANT:
- Return only valid JSON - no explanations, prefaces, or additional formatting
- All values must be exactly as specified in the categories above
- All field values must be wrapped in double quotes
- Ensure all field names match exactly as shown
- Do not add any additional fields or comments"""

    return make_api_call(prompt)

