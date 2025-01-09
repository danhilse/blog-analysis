import anthropic
from typing import Dict, List, Any, Optional
import os
from dataclasses import dataclass
from dotenv import load_dotenv
from static import CATEGORIES, USE_CASES, BRAND_GUIDELINES, corporate_urls
import json
import asyncio

# Load environment variables from .env file
load_dotenv()


@dataclass
class AnalysisResult:
    """Container for analysis results with status and error tracking"""
    success: bool
    result: Optional[str] = None
    error: Optional[str] = None


class BlogAnalyzer:
    def __init__(self):
        api_key = os.getenv("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY not found in environment variables")

        self.client = anthropic.Anthropic(api_key=api_key)

        # Use static configurations
        self.categories = CATEGORIES
        self.use_cases = USE_CASES
        self.brand_guidelines = BRAND_GUIDELINES

    def _get_cached_system_prompt(self, prompt_type: str) -> List[Dict[str, Any]]:
        """
        Get system prompts with appropriate cache control based on type.

        Args:
            prompt_type: Type of analysis being performed
        """
        prompts = {
            "categorize": [
                {
                    "type": "text",
                    "text": """You are a content categorization specialist for a B2B marketing technology website. Analyze content and determine the most appropriate category based on its primary focus and purpose."""
                },
                {
                    "type": "text",
                    "text": f"""Here are the available categories:
            {', '.join(self.categories)}

            Guidelines for categorization:

            1. Primary Focus: Categorize based on the main topic and purpose, not just keyword matches
            2. Hierarchy: Choose specific categories over general ones when applicable
            3. Key Distinctions:
               - 'Marketing Automation': platform-specific features and implementation
               - 'Automation Technology & Strategy': broader automation strategy and selection
               - 'Email Marketing': content strategy and campaigns
               - 'Email Deliverability': technical delivery and inbox placement
               - 'Customer Marketing': retention and advocacy programs
               - 'Customer Journey': overall experience mapping and optimization
            4. Default Rules:
               - 'AI and Marketing': only when AI/ML is the primary focus
               - 'Marketing Strategy': only for high-level planning across channels
               - 'Corporate': company news and announcements
               - 'Uncategorized': only when no other category clearly fits"""
                },
                {
                    "type": "text",
                    "text": """Output ONLY valid JSON format with exactly these keys and no others:
            {
                "category": "exact category name from the list",
                "reasoning": "Two sentence explanation of why this category best fits the content's primary focus and purpose"
            }"""
                }
            ],
            "brand_alignment": [
                {
                    "type": "text",
                    "text": "You are a brand alignment specialist analyzing content against brand guidelines.",
                },
                {
                    "type": "text",
                    "text": self.brand_guidelines,
                    "cache_control": {"type": "ephemeral"}
                }
            ],
            "use_case": [
                {
                    "type": "text",
                    "text": "You are a content strategy specialist. Analyze content and determine which use case best matches the content's purpose and outcomes."
                },
                {
                    "type": "text",
                    "text": f'''When categorizing content, follow this prioritization framework:
1. Primary Business Outcome: What is the ultimate goal the content aims to achieve?
2. Target Audience Stage: Is this aimed at prospects, existing customers, or partners?
3. Strategic Intent: Look for the higher-level business objective before tactical implementation
4. Implementation Methods: Consider tactical elements (like nurture campaigns or webinars) as supporting evidence, not primary categorization factors

Key Distinctions:
- Content about tactical methods (email nurture, webinars, etc.) should be categorized by the strategic goal it serves
- When content mentions multiple use cases, prioritize the main business outcome
- Customer lifecycle stage should guide categorization (e.g., new customer content vs. existing customer growth)

                    Here are the possible use cases and their descriptions:

{self._format_use_cases()}''',
                    "cache_control": {"type": "ephemeral"}
                },
                {
                    "type": "text",
                    "text": """Output ONLY valid JSON format with exactly these keys and no others:
            {
                "use case": "exact use case name",
                "reasoning": "Two sentence justification without any line breaks or special characters",
                'next best use case': 'exact use case name of the next best use case'
            }"""
                }
            ],
            "use_case_type_2": [
                {
                    "type": "text",
                    "text": "You are a content strategy specialist. Analyze content and determine which use case best matches the content's purpose and outcomes."
                },
                {
                    "type": "text",
                    "text": f"Here are the possible use cases and their descriptions:\n\n{self._format_use_cases()}",
                    "cache_control": {"type": "ephemeral"}
                },
                {
                    "type": "text",
                    "text": """Output ONLY valid JSON format with exactly these keys and no others:
            {
                "use case": "exact use case name",
                "reasoning": "Two sentence justification without any line breaks or special characters"
            }"""
                }
            ],
            "use_case_multi": [
                {
                    "type": "text",
                    "text": """You are a content strategy specialist analyzing content for all relevant marketing use cases. Identify all use cases that significantly align with the content's purpose, outcomes, and tactical implementation."""
                },
                {
                    "type": "text",
                    "text": f"""Here are the possible use cases and their descriptions:\n\n{self._format_use_cases()}

        When analyzing content:
        1. Consider both strategic goals and tactical implementations
        2. Look for multiple layers of intent and audience
        3. Assess confidence based on:
           - Direct mention and focus on the use case
           - Alignment with use case objectives
           - Depth of coverage for the use case
           - Target audience match""",
                    "cache_control": {"type": "ephemeral"}
                },
                    {
                    "type": "text",
                    "text": """Output ONLY valid JSON format with exactly these keys and no others:
        {
                    "primary_use_case": {
                    "name": "exact use case name that best fits",
                "confidence": float between 0-1,
                "reasoning": "One sentence explanation"
            },
            "additional_use_cases": [
                {
                    "name": "exact use case name",
                    "confidence": float between 0-1,
                    "reasoning": "One sentence explanation",
                    "relationship": "Brief explanation of how this use case relates to or supports the primary use case"
                }
            ],
            "analysis_confidence": float between 0-1
        }

        Note:
        - Only include use cases with confidence > 0.5
        - Order additional use cases by confidence (highest first)
        - Explain relationships between use cases to show content flow and dependencies"""
                }
            ]
        }
        return prompts.get(prompt_type, [])

    def _format_use_cases(self) -> str:
        """Format use cases and descriptions for prompt"""
        return "\n".join([
            f"- {use_case}: {description}"
            for use_case, description in self.use_cases.items()
        ])

    async def analyze_content(self, content: str, analysis_type: str) -> AnalysisResult:
        """
        Analyze content based on specified analysis type.

        Args:
            content: The content to analyze
            analysis_type: Type of analysis to perform

        Returns:
            AnalysisResult containing success status and result/error
        """
        try:
            system_prompt = self._get_cached_system_prompt(analysis_type)

            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=1024,
                system=system_prompt,
                temperature=0.25,
                messages=[
                    {
                        "role": "user",
                        "content": f"Analyze this content:\n\n{content}"
                    }
                ]
            )

            result = response.content[0].text.strip()

            # For categorization, validate the JSON structure and category
            if analysis_type == "categorize":
                try:
                    result_json = json.loads(result)
                    if "category" not in result_json:
                        return AnalysisResult(
                            success=False,
                            result=result,
                            error="Response missing 'category' field"
                        )
                    if result_json["category"] not in self.categories:
                        return AnalysisResult(
                            success=False,
                            result=result,
                            error=f"Invalid category returned: {result_json['category']}"
                        )
                except json.JSONDecodeError as e:
                    return AnalysisResult(
                        success=False,
                        result=result,
                        error=f"Invalid JSON response format: {str(e)}"
                    )

            return AnalysisResult(success=True, result=result)

        except Exception as e:
            return AnalysisResult(
                success=False,
                result=str(e),
                error=f"Analysis failed: {str(e)}"
            )

class AnalysisManager:
    """Manages multiple analysis operations for a blog article"""

    def __init__(self):
        self.analyzer = BlogAnalyzer()

    async def analyze_article(self, content: str) -> Dict[str, AnalysisResult]:
        """
        Perform all analyses on an article

        Args:
            content: The article content to analyze

        Returns:
            Dictionary of analysis results keyed by analysis type
        """
        analyses = {
            "category": await self.analyzer.analyze_content_with_retry(content, "categorize"),
            "brand_alignment": await self.analyzer.analyze_content_with_retry(content, "brand_alignment"),
            "summary": await self.analyzer.analyze_content_with_retry(content, "summarize"),
            "use_case": await self.analyzer.analyze_content_with_retry(content, "use_case"),
            "use_case_type_2": await self.analyzer.analyze_content_with_retry(content, "use_case_type_2"),
            "use_case_multi": await self.analyzer.analyze_content_with_retry(content, "use_case_multi")

        }
        return analyses


def get_analysis_manager() -> AnalysisManager:
    """Factory function to get an AnalysisManager instance."""
    return AnalysisManager()