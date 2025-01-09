import json
import os
import asyncio
from datetime import datetime
from ai import get_analysis_manager
import helper
from urllib.parse import urlparse
from static import useCaseCats, preSorted


def generate_unique_id(url):
    """Generate a unique ID from URL by removing protocol and special chars"""
    parsed = urlparse(url)
    # Combine netloc and path, remove special chars, convert to lowercase
    unique_id = (parsed.netloc + parsed.path).lower()
    return ''.join(c for c in unique_id if c.isalnum())


def load_processed_data():
    """Load the processed data from JSON file"""
    processed_file = 'output/processed.json'
    try:
        if os.path.exists(processed_file):
            with open(processed_file, 'r') as f:
                return json.load(f)
        return {}
    except json.JSONDecodeError:
        print(f"Warning: Corrupted processed.json found. Creating backup and starting fresh.")
        if os.path.exists(processed_file):
            backup_file = f'output/processed_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
            os.rename(processed_file, backup_file)
        return {}


def save_processed_data(article_id, article_analysis):
    """Save a single article analysis to the processed data file"""
    processed_file = 'output/processed.json'
    max_retries = 3
    retry_delay = 1  # seconds

    for attempt in range(max_retries):
        try:
            # Load current data
            current_data = load_processed_data()

            # Update with new analysis
            current_data[article_id] = article_analysis

            # Save updated data
            with open(processed_file, 'w') as f:
                json.dump(current_data, f, indent=2)

            return True
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"Error saving data (attempt {attempt + 1}): {str(e)}. Retrying...")
                asyncio.sleep(retry_delay)
            else:
                print(f"Failed to save data after {max_retries} attempts: {str(e)}")
                return False


def get_use_case(url):
    # Normalize the URL by trimming whitespace
    normalized_url = url.strip()

    # Look up the category in pre_sorted, return None if not found
    return preSorted.get(normalized_url)

async def analyze_article(article, analysis_manager):
    """Analyze a single article and return its analysis"""
    url = article['url']

    article_analysis = {
        'url': url,
        'title': article.get('basic_info', {}).get('title', ''),
        'category': article.get('basic_info', {}).get('category', None),  # Added category field
        'publication_date': article.get('basic_info', {}).get('publication_date', ''),
        'modified_date': article.get('basic_info', {}).get('modified_date', ''),
        'processed_timestamp': datetime.now().isoformat(),
        'red_flags': {
            'matches': article.get('red_flags', {}).get('matches', [])
        },
        'pre sorted use case': get_use_case(url)
    }

    # Clean the content
    cleaned_content = helper.clean_content(article.get('content', ''))

    # Analyze for category
    category_result = await analysis_manager.analyzer.analyze_content(
        cleaned_content,
        "categorize"
    )

    if category_result.success:
        category_data = json.loads(category_result.result)
        article_analysis['ai_category'] = category_data['category']
        article_analysis['ai_category_reasoning'] = category_data['reasoning']
        article_analysis['ai_category_status'] = 'success'
    else:
        article_analysis['ai_category'] = 'uncategorized'
        article_analysis['ai_category_reasoning'] = None
        article_analysis['ai_category_status'] = 'failed'
        article_analysis['category_error'] = category_result.result
        article_analysis['category_error_message'] = category_result.error

    # Analyze for first use case type
    analysis_result = await analysis_manager.analyzer.analyze_content(
        cleaned_content,
        "use_case"
    )

    if analysis_result.success:
        use_case_data = json.loads(analysis_result.result)
        use_case = use_case_data['use case']
        article_analysis['use_case'] = use_case
        article_analysis['use_case_reasoning'] = use_case_data['reasoning']
        article_analysis['use_case_alt'] = use_case_data['next best use case']
        article_analysis['analysis_status'] = 'success'

        # Look up getKeepGrow and cmoPriority from useCaseCats
        if use_case in useCaseCats['useCases']:
            use_case_info = useCaseCats['useCases'][use_case]
            article_analysis['getKeepGrow'] = use_case_info.get('getKeepGrow')
            article_analysis['cmoPriority'] = use_case_info.get('cmoPriority')
        else:
            article_analysis['getKeepGrow'] = 'unknown'
            article_analysis['cmoPriority'] = 'unknown'
    else:
        article_analysis['use_case'] = 'unclassified'
        article_analysis['use_case_reasoning'] = None
        article_analysis['analysis_status'] = 'failed'
        article_analysis['error'] = analysis_result.result
        article_analysis['error_message'] = analysis_result.error
        article_analysis['getKeepGrow'] = 'unknown'
        article_analysis['cmoPriority'] = 'unknown'

    # Analyze for second use case type
    analysis_result = await analysis_manager.analyzer.analyze_content(
        cleaned_content,
        "use_case_type_2"
    )

    if analysis_result.success:
        use_case_data = json.loads(analysis_result.result)
        use_case_2 = use_case_data['use case']
        article_analysis['use_case_type_2'] = use_case_2
        article_analysis['use_case_reasoning_type_2'] = use_case_data['reasoning']

        # # Look up getKeepGrow and cmoPriority for second use case
        # if use_case_2 in useCaseCats['useCases']:
        #     use_case_info = useCaseCats['useCases'][use_case_2]
        #     article_analysis['getKeepGrow_type_2'] = use_case_info.get('getKeepGrow')
        #     article_analysis['cmoPriority_type_2'] = use_case_info.get('cmoPriority')
        # else:
        #     article_analysis['getKeepGrow_type_2'] = 'unknown'
        #     article_analysis['cmoPriority_type_2'] = 'unknown'
    else:
        article_analysis['use_case_type_2'] = 'unclassified'
        article_analysis['use_case_reasoning_type_2'] = None
        article_analysis['error_2'] = analysis_result.result

    # Analyze for second use case type
    analysis_result = await analysis_manager.analyzer.analyze_content(
        cleaned_content,
        "use_case_multi"
    )

    if analysis_result.success:
        use_case_data = json.loads(analysis_result.result)
        article_analysis['use_case_multi_primary'] = use_case_data['primary_use_case']
        article_analysis['use_case_multi_addl'] = use_case_data['additional_use_cases']

        # # Look up getKeepGrow and cmoPriority for second use case
        # if use_case_2 in useCaseCats['useCases']:
        #     use_case_info = useCaseCats['useCases'][use_case_2]
        #     article_analysis['getKeepGrow_type_2'] = use_case_info.get('getKeepGrow')
        #     article_analysis['cmoPriority_type_2'] = use_case_info.get('cmoPriority')
        # else:
        #     article_analysis['getKeepGrow_type_2'] = 'unknown'
        #     article_analysis['cmoPriority_type_2'] = 'unknown'
    else:
        article_analysis['use_case_multi'] = 'failed'
        article_analysis['error_multi'] = analysis_result.result

    return article_analysis

async def process_all_content():
    """Process all content across all categories"""
    # Initialize output directory
    os.makedirs('output', exist_ok=True)

    # Load input data
    with open('output/blog.json', 'r') as f:
        content_data = json.load(f)

    # Initialize AI analysis manager
    analysis_manager = get_analysis_manager()

    # Process all categories
    for category, articles in content_data['analyses'].items():
        print(f"Processing category: {category}")

        # Skip if category is empty
        if not articles:
            continue

        # Process each article in the category
        for article in articles:
            try:
                article_id = generate_unique_id(article['url'])

                # Analyze article
                article_analysis = await analyze_article(article, analysis_manager)

                # Add category information
                # article_analysis['category'] = category

                # Save article analysis with retries
                if save_processed_data(article_id, article_analysis):
                    print(f"Processed and saved {category} article: {article['url']}")
                else:
                    print(f"Failed to save {category} article: {article['url']}")

            except Exception as e:
                print(f"Error processing article {article['url']}: {str(e)}")
                continue

    print("Processing complete!")


def main():
    """Main entry point with asyncio setup"""
    asyncio.run(process_all_content())


if __name__ == "__main__":
    main()