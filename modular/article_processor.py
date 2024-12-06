import json
import re
import pandas as pd
from datetime import datetime
from urllib.parse import urlparse
import textstat
import os
from typing import Dict, List, Tuple


def clean_content(json_input: str) -> str:
    """
    Cleans and converts JSON content to Markdown.

    Parameters:
        json_input (str): JSON string containing the "content" field.

    Returns:
        str: Cleaned Markdown content.
    """
    try:
        # Parse the JSON input
        data = json.loads(json_input)
        content = data.get("content", "")
    except json.JSONDecodeError:
        # If input is not a valid JSON, assume it's raw content
        content = json_input

    # Remove image markers
    content = re.sub(r'\[CONTENT IMAGE:.*?\]', '', content)

    # Remove source lines
    content = re.sub(r'Source:\s*https?://\S+', '', content)

    # Convert headers
    content = re.sub(r'H2:\s*', '## ', content)

    # Normalize line breaks and whitespace
    content = re.sub(r'\n{3,}', '\n\n', content)
    content = re.sub(r'\n\s*\n', '\n\n', content)
    content = re.sub(r'[ \t]+', ' ', content)

    # Remove remaining brackets
    content = re.sub(r'\[.*?\]', '', content)

    return content.strip()


def calculate_word_count(content: str) -> int:
    """
    Calculate word count from cleaned content.

    Parameters:
        content (str): Raw content string

    Returns:
        int: Number of words
    """
    clean_text = clean_content(content)
    return len(clean_text.split())


def load_yoast_keywords() -> Dict[str, str]:
    """
    Load Yoast keywords from Excel file and create URL mapping.

    Returns:
        Dict[str, str]: Mapping of URLs to keywords
    """
    try:
        yoast_df = pd.read_excel('../resources/yoast-blog-keywords.xlsx', header=None)
        url_to_keyword = {}
        for _, row in yoast_df.iterrows():
            url = row[4]  # URL column
            keyword = row[2]  # Keyword column
            if pd.notna(url) and pd.notna(keyword):
                url_to_keyword[url.strip()] = keyword.strip()
        return url_to_keyword
    except Exception as e:
        print(f"Error loading Yoast keywords: {e}")
        return {}


def import_performance_data() -> pd.DataFrame:
    """
    Import performance metrics from Excel file.

    Returns:
        pd.DataFrame: Performance data indexed by URL slug
        pd.DataFrame: Performance data indexed by URL slug
    """
    try:
        file_path = os.path.join('../resources', 'performance.xlsx')
        performance_df = pd.read_excel(file_path)

        def extract_slug(path):
            if pd.isna(path):
                return ''
            parsed = urlparse(path)
            path_parts = [p for p in parsed.path.strip('/').split('/') if p]
            return path_parts[-1] if path_parts else ''

        performance_df['slug'] = performance_df['Page path + query string'].apply(extract_slug)
        performance_df = performance_df[performance_df['slug'] != '']
        performance_df.set_index('slug', inplace=True)
        return performance_df
    except Exception as e:
        print(f"Warning: Could not load performance data: {str(e)}")
        return None


def count_personal_pronouns(text: str) -> Dict[str, any]:
    """
    Count personal pronouns while excluding those within quotes.

    Parameters:
        text (str): Content to analyze

    Returns:
        Dict[str, any]: Analysis results including counts and examples
    """
    quote_pairs = {'"': '"', "'": "'"}
    quotes: List[Tuple[int, int]] = []
    current_pos = 0
    text_length = len(text)

    # Find quoted regions
    while current_pos < text_length:
        next_quote = None
        next_pos = text_length

        for open_quote in quote_pairs.keys():
            pos = text.find(open_quote, current_pos)
            if pos != -1 and pos < next_pos:
                next_quote = open_quote
                next_pos = pos

        if next_quote is None:
            break

        start = next_pos
        end_quote = quote_pairs[next_quote]
        end = text.find(end_quote, start + 1)

        if end == -1:
            end = text_length - 1

        quotes.append((start, end))
        current_pos = end + 1

    def is_in_quotes(pos: int) -> bool:
        return any(start <= pos <= end for start, end in quotes)

    # Find sentences
    sentence_endings = re.compile(r'([.!?])')
    sentences: List[Tuple[str, int, int]] = []
    start = 0

    for match in sentence_endings.finditer(text):
        end = match.end()
        sentence = text[start:end].strip()
        if sentence:
            sentences.append((sentence, start, end))
        start = end

    if start < text_length:
        sentence = text[start:].strip()
        if sentence:
            sentences.append((sentence, start, text_length))

    # Find pronouns
    pronouns = r'\b(I|me|my|mine|myself)\b'
    matches = []
    sentences_with_pronouns: List[str] = []
    seen_sentences: set = set()

    for match in re.finditer(pronouns, text, re.IGNORECASE):
        if not is_in_quotes(match.start()):
            matches.append(match)
            for sentence, s_start, s_end in sentences:
                if s_start <= match.start() < s_end and sentence not in seen_sentences:
                    sentences_with_pronouns.append(sentence)
                    seen_sentences.add(sentence)
                    break

    return {
        'count': len(matches),
        'found_pronouns': [m.group() for m in matches],
        'quoted_regions': quotes,
        'sentences_with_pronouns': sentences_with_pronouns,
        'flag': len(matches) > 0
    }


def format_seo_data(basic_info: Dict, seo_analysis: Dict, target_keyword: str) -> Dict:
    """
    Format SEO data from content analysis.

    Parameters:
        basic_info (Dict): Basic content information
        seo_analysis (Dict): SEO analysis data
        target_keyword (str): Target keyword from Yoast

    Returns:
        Dict: Formatted SEO data
    """
    return {
        'current_target_keyword': target_keyword,
        'meta_description_present': seo_analysis.get('meta_description', {}).get('present', False),
        'h1_present': seo_analysis.get('headings', {}).get('h1_present', False),
        'h2_count': seo_analysis.get('headings', {}).get('h2_count', 0),
        'h3_count': seo_analysis.get('h3_count', 0)
    }


def parse_date(date_str: str) -> str:
    """
    Parse date string to consistent format.

    Parameters:
        date_str (str): Date string to parse

    Returns:
        str: Formatted date string
    """
    if date_str:
        try:
            return datetime.strptime(date_str.split('T')[0], '%Y-%m-%d').strftime('%Y-%m-%d')
        except ValueError:
            return 'No Date'
    return 'No Date'