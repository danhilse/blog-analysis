import json
from typing import Dict, List
import logging
import os
from datetime import datetime
from urllib.parse import urlparse
import sys
import ast
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
import re
from static import corporate_urls


# Enhanced logging setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        # logging.FileHandler('scraper.log')
    ]
)
logger = logging.getLogger(__name__)

class BlogAnalyzer:
    def __init__(self):
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }



    def analyze_webpage(self, url: str) -> Dict:
        """Main method to analyze a webpage."""
        try:
            logger.info(f"Fetching URL: {url}")

            # Determine if we need advanced scraping based on URL
            if 'connect.act-on.com' in url:
                html_content = self._fetch_with_selenium(url)
            else:
                # Use regular requests for the main blog
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                    'Accept-Language': 'en-US,en;q=0.5',
                    'Accept-Encoding': 'gzip, deflate',
                    'Connection': 'keep-alive',
                    'Upgrade-Insecure-Requests': '1',
                    'Cache-Control': 'max-age=0'
                }
                response = requests.get(url, headers=headers, timeout=10)
                response.raise_for_status()
                html_content = response.text

            # Create soup object
            soup = BeautifulSoup(html_content, 'lxml')

            content = self.get_content(soup)

            # Create analysis dictionary
            analysis = {
                'url': url,
                'analysis_timestamp': datetime.now().isoformat(),
                'basic_info': self.get_basic_info(soup),
                'seo_analysis': self.get_seo_analysis(soup),
                'multimedia_assessment': self.get_multimedia_assessment(soup),
                'content': content,
                'red_flags': self.check_red_flags(content),  # Add red flags analysis
                'related_content': self.get_related_content(soup),
                'videos': self.get_videos(soup)
            }

            return analysis

        except Exception as e:
            logger.error(f"Error processing {url}: {str(e)}", exc_info=True)
            return {
                'url': url,
                'analysis_timestamp': datetime.now().isoformat(),
                'error': str(e),
                'status': 'failed'
            }

    def _fetch_with_selenium(self, url: str) -> str:
        """Fetch content using Selenium for pages that require JavaScript."""
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # Run in headless mode
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')

        # Add additional headers to avoid detection
        chrome_options.add_argument(
            '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')

        # Initialize webdriver with ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)

        try:
            driver.get(url)

            # Wait for article content to be present
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'article-content'))
            )

            # Additional wait for dynamic content
            time.sleep(2)

            return driver.page_source
        finally:
            driver.quit()

    def check_red_flags(self, content: str) -> Dict:
        """Check content for specific red flag phrases using regex."""
        red_flags = {
            'matches': [],
            'count': 0
        }

        patterns = [
            r'rethink\s+marketing\s+podcast',
            r'(?:in\s+the\s+comments|tell\s+us\s+in\s+the\s+comments)',
            r'adaptive',
            r'growth\s+marketing\s+platform'
        ]

        for pattern in patterns:
            # Compile pattern with IGNORECASE and MULTILINE flags
            regex = re.compile(pattern, re.IGNORECASE | re.MULTILINE)
            for match in regex.finditer(content):
                red_flags['matches'].append({
                    'pattern': pattern,
                    'matched_text': match.group(),
                    'position': match.start()
                })

        red_flags['count'] = len(red_flags['matches'])
        return red_flags

    def get_basic_info(self, soup: BeautifulSoup) -> Dict:
        """Extract basic information about the blog post."""
        basic_info = {
            'title': '',
            'publication_date': '',
            'modified_date': '',
            'url': '',
            'description': '',
            'category': None  # Added category field
        }

        # First try to get metadata from Yoast SEO schema
        schema_script = soup.find('script', {'type': 'application/ld+json', 'class': 'yoast-schema-graph'})
        if schema_script and schema_script.string:
            try:
                schema_data = json.loads(schema_script.string)
                for item in schema_data.get('@graph', []):
                    if item.get('@type') == 'WebPage':
                        basic_info.update({
                            'publication_date': item.get('datePublished', ''),
                            'modified_date': item.get('dateModified', ''),
                            'url': item.get('url', ''),
                            'description': item.get('description', '')
                        })
                        break
            except json.JSONDecodeError as e:
                logger.error(f"Error parsing JSON-LD schema: {e}")

        # Extract category from the breadcrumbs widget
        breadcrumbs = soup.find('div', class_='breadcrumbs')
        if breadcrumbs:
            # Find all list items in the breadcrumbs
            list_items = breadcrumbs.find_all('li', class_='elementor-icon-list-item')
            for item in list_items:
                # Look for an anchor tag with 'category' in the href
                category_link = item.find('a', href=lambda x: x and '/category/' in x)
                if category_link:
                    category_text = category_link.get_text(strip=True)
                    if category_text:
                        basic_info['category'] = category_text
                        break

        # Fallback to meta tags if needed
        if not basic_info['publication_date']:
            pub_date_meta = soup.find('meta', property='article:published_time')
            if pub_date_meta:
                basic_info['publication_date'] = pub_date_meta.get('content', '')

        if not basic_info['modified_date']:
            mod_date_meta = soup.find('meta', property='article:modified_time')
            if mod_date_meta:
                basic_info['modified_date'] = mod_date_meta.get('content', '')

        # Get title
        title = soup.find('h1', class_='elementor-heading-title')
        if title:
            basic_info['title'] = title.get_text(strip=True)

        # Get URL
        canonical = soup.find('link', rel='canonical')
        if canonical:
            basic_info['url'] = canonical.get('href', '')

        # Get description
        if not basic_info['description']:
            meta_desc = soup.find('meta', attrs={'name': 'description'})
            if meta_desc:
                basic_info['description'] = meta_desc.get('content', '')

        return basic_info
    def get_seo_analysis(self, soup: BeautifulSoup) -> Dict:
        """Analyze SEO elements of the page."""
        seo_analysis = {
            'meta_description': {
                'present': False,
                'content': ''
            },
            'headings': {
                'h1_present': False,
                'h2_count': 0,
                'h3_count': 0
            }
        }

        meta_desc = soup.find('meta', attrs={'name': 'description'})
        if meta_desc:
            seo_analysis['meta_description']['present'] = True
            seo_analysis['meta_description']['content'] = meta_desc.get('content', '')

        h1_tags = soup.find_all('h1')
        h2_tags = soup.find_all('h2')
        h3_tags = soup.find_all('h3')

        seo_analysis['headings']['h1_present'] = len(h1_tags) > 0
        seo_analysis['headings']['h2_count'] = len(h2_tags)
        seo_analysis['headings']['h3_count'] = len(h3_tags)

        return seo_analysis

    def get_multimedia_assessment(self, soup: BeautifulSoup) -> Dict:
        """Assess multimedia content on the page."""
        multimedia = {
            'header_image': None,
            'content_images': [],
            'total_image_count': 0,
            'outdated_widgets': []  # New field
        }

        # Find outdated download button widgets
        main_content = soup.find('div', class_='elementor-widget-theme-post-content')
        if main_content:
            # Find all wp-block-buttons divs
            button_blocks = main_content.find_all('div', class_='wp-block-buttons')
            for block in button_blocks:
                # Check for standard style buttons with download/ebook links
                buttons = block.find_all('div', class_='wp-block-button is-style-standard')
                for button in buttons:
                    link = button.find('a', class_='wp-block-button__link')
                    if link and ('download' in link.get_text().lower() or 'ebook' in link.get_text().lower()):
                        multimedia['outdated_widgets'].append({
                            'type': 'download_button',
                            'text': link.get_text(strip=True),
                            'url': link.get('href', '')
                        })

        multimedia['outdated_widget_count'] = len(multimedia['outdated_widgets'])

        # Find header/featured image
        featured_img = soup.find('div', class_='elementor-widget-theme-post-featured-image')
        if featured_img:
            img = featured_img.find('img')
            if img and not self.is_logo_image(img):
                multimedia['header_image'] = {
                    'src': img.get('src', ''),
                    'alt': img.get('alt', ''),
                    'width': int(img.get('width', 0)) if img.get('width') else 0,
                    'height': img.get('height', '')
                }

        # Find content images
        main_content = soup.find('div', class_='elementor-widget-theme-post-content')
        if main_content:
            for element in main_content.find_all(['img', 'figure']):
                img = None
                if element.name == 'figure':
                    if 'wp-block-image' in element.get('class', []):
                        img = element.find('img')
                elif element.name == 'img':
                    img = element

                if img and not self.is_logo_image(img):
                    header_src = multimedia['header_image']['src'] if multimedia['header_image'] else None
                    if header_src and img.get('src') == header_src:
                        continue

                    image_info = {
                        'src': img.get('src', ''),
                        'alt': img.get('alt', ''),
                        'width': img.get('width', ''),
                        'height': img.get('height', '')
                    }
                    if not any(existing['src'] == image_info['src'] for existing in multimedia['content_images']):
                        multimedia['content_images'].append(image_info)

        multimedia['total_image_count'] = (
            (1 if multimedia['header_image'] else 0) +
            len(multimedia['content_images'])
        )

        return multimedia

    def get_content(self, soup: BeautifulSoup) -> str:
        """Extract the main content in a readable format."""
        content = []

        def get_spaced_text(element):
            """Get text from an element with proper spacing around inline elements."""
            text_parts = []
            for item in element.children:
                if item.name is None:  # Direct text node
                    text_parts.append(item.strip())
                else:  # HTML element
                    text_parts.append(item.get_text().strip())
            return ' '.join(filter(None, text_parts))

        def process_content_block(content_block):
            """Process a content block and extract text content."""
            if not content_block:
                return

            for element in content_block.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'ul', 'ol', 'img', 'figure', 'div']):
                if element.name in ['h1', 'h2', 'h3', 'h4']:
                    content.append(f"\n{element.name.upper()}: {get_spaced_text(element)}")
                elif element.name == 'p':
                    text = get_spaced_text(element)
                    if text:  # Only add non-empty paragraphs
                        content.append(text)
                elif element.name in ['ul', 'ol']:
                    for index, li in enumerate(element.find_all('li', recursive=False)):
                        if element.name == 'ol':
                            # For ordered lists, use the actual number
                            content.append(f"{index + 1}. {get_spaced_text(li)}")
                        else:
                            content.append(f"- {get_spaced_text(li)}")
                elif element.name == 'img' and not self.is_logo_image(element):
                    alt_text = element.get('alt', 'No alt text provided')
                    src = element.get('src', '')
                    if src and alt_text:
                        content.append(f"\n[CONTENT IMAGE: {alt_text}]\nSource: {src}\n")
                elif element.name == 'figure':
                    if 'wp-block-image' in element.get('class', []):
                        img = element.find('img')
                        if img and not self.is_logo_image(img):
                            alt_text = img.get('alt', 'No alt text provided')
                            src = element.get('src', '')
                            if src and alt_text:
                                content.append(f"\n[CONTENT IMAGE: {alt_text}]\nSource: {src}\n")

        # Try different content container selectors in order of preference
        content_containers = [
            # Main blog content (Elementor)
            soup.find('div', class_='elementor-widget-theme-post-content'),
            # Help center article content
            soup.find('section', class_='content article-content'),
            # Help center article body
            soup.find('div', class_='article-body'),
            # Help center main content
            soup.find('main', class_='content'),
            # Generic article content
            soup.find('article'),
            # Last resort: main tag
            soup.find('main')
        ]

        # Log what we found
        logger.info("Content containers found:")
        for i, container in enumerate(content_containers):
            if container:
                logger.info(f"Container {i}: {container.name} with class {container.get('class', [])}")

        # Try each container until we find content
        for container in content_containers:
            if container:
                process_content_block(container)
                if content:  # If we found content, stop looking
                    break

        # If still no content, try to get any text from body
        if not content:
            logger.warning("No content found in standard containers, attempting to extract from body")
            body = soup.find('body')
            if body:
                # Get text while excluding script and style elements
                for script in body(['script', 'style']):
                    script.decompose()
                text = body.get_text()
                if text.strip():
                    content.append(text.strip())

        return '\n\n'.join(content)
    def get_videos(self, soup: BeautifulSoup) -> List[Dict]:
        """Extract video content from the page."""
        videos = []
        video_embeds = soup.find_all('figure', class_='wp-block-embed-youtube')

        for embed in video_embeds:
            iframe = embed.find('iframe')
            if iframe:
                video_info = {
                    'title': iframe.get('title', ''),
                    'embed_url': iframe.get('src', ''),
                    'width': iframe.get('width', ''),
                    'height': iframe.get('height', '')
                }
                videos.append(video_info)

        return videos

    def get_related_content(self, soup: BeautifulSoup) -> List[Dict]:
        """Extract related content cards."""
        related_content = []
        content_cards = soup.find_all('div', class_='jet-listing-grid__item')

        for card in content_cards:
            title_elem = card.find('h3', class_='elementor-heading-title')
            if not title_elem or not title_elem.find('a'):
                continue

            link = title_elem.find('a')
            description = ""
            desc_elem = title_elem.find_next('h3', class_='elementor-heading-title')
            if desc_elem and desc_elem.find('a'):
                description = desc_elem.find('a').get_text(strip=True)

            related_content.append({
                'title': title_elem.get_text(strip=True),
                'url': link['href'],
                'description': description,
                'type': 'blog'
            })

        return related_content

    def is_logo_image(self, img) -> bool:
        """Check if an image is a logo based on specific logo URLs."""
        src = img.get('src', '')
        logo_urls = [
            "https://act-on.com/wp-content/uploads/2023/03/AO-logo_Color_616x225.svg",
            "https://act-on.com/wp-content/uploads/2023/10/AO-logo_Color_Icon-100-200x200.jpg"
        ]
        return src in logo_urls

    def get_content_type(self, url: str) -> str:
        """Determine content type from URL structure."""
        path = urlparse(url).path.lower()
        if '/blog/' in path:
            return 'blog'
        elif '/webinars/' in path or '/on-demand-webinars/' in path:
            return 'webinar'
        elif '/case-studies/' in path:
            return 'case-study'
        elif '/ebooks/' in path or '/white-papers/' in path:
            return 'resource'
        else:
            return 'other'

def get_all_blogs():
    with open('all blogs.txt', 'r', encoding='utf-8') as file:
            # Read the entire file content
            content = file.read().strip()

            # Handle potential variations in the file content
            # Remove any leading/trailing whitespace or linebreaks
            content = content.strip()

            # If the content doesn't start with '[' and end with ']', try to fix it
            if not (content.startswith('[') and content.endswith(']')):
                content = f"[{content}]"

            # Use ast.literal_eval to safely evaluate the string as a Python literal
            blogs_array = ast.literal_eval(content)

            # Verify that we got a list
            if not isinstance(blogs_array, list):
                raise ValueError("File content is not a valid array/list")

            return blogs_array


def filter_valid_urls(urls):
    """
    Filter out data-sheets URLs and duplicates, returning only unique valid URLs.

    Args:
        urls (list): List of URLs to filter

    Returns:
        list: Filtered list of unique, valid URLs
    """
    # Convert to set to remove duplicates while filtering invalid URLs
    unique_valid_urls = {url for url in urls if "/learn/data-sheets/" not in url}
    # Convert back to list and sort for consistency
    return sorted(list(unique_valid_urls))

def main():
    analyzer = BlogAnalyzer()

    # URLs from first presentation
    urls = [
        "https://act-on.com/learn/blog/manufacturing-industry-slow-to-adopt-emerging-digital-marketing-software/",
        "https://act-on.com/learn/blog/retention-marketing-how-we-reached-400-customer-accounts/",
        # "https://act-on.com/learn/data-sheets/advanced-crm-mapping/",
        # "https://act-on.com/learn/blog/how-and-why-you-should-calculate-customer-lifetime-value-clv/",
        # "https://act-on.com/learn/blog/pipeline-generation-face-economic-headwinds-and-win/",
        # "https://act-on.com/learn/blog/what-is-customer-marketing-2/"
    ]

    urls = [
        'https://act-on.com/learn/blog/pipeline-generation-face-economic-headwinds-and-win/',
        'https://act-on.com/learn/blog/5-ways-marketing-leaders-help-sales-expand-pipeline/',
        'https://act-on.com/learn/blog/how-to-reduce-customer-attrition-and-keep-your-best-customers/',
        'https://act-on.com/learn/blog/what-is-lead-scoring-for-marketing-and-what-are-the-benefits/',
        'https://act-on.com/learn/blog/5-lead-nurture-campaigns-that-build-pipeline-and-support-roi/',
        'https://act-on.com/learn/blog/what-is-customer-marketing-2/',
        'https://act-on.com/learn/blog/lead-scoring-model-building-a-framework-to-drive-conversion/',
        'https://act-on.com/learn/blog/lead-scoring-tools-and-tactics-to-convert-customers/',
        'https://act-on.com/learn/blog/feeding-the-funnel-how-to-build-nurture-programs-that-drive-pipeline/',
        'https://act-on.com/learn/blog/use-trigger-campaigns-effectively-nurture-leads/',
        'https://act-on.com/learn/blog/increase-customer-retention-why-multichannel-marketing-is-an-underrated-tool/',
        'https://act-on.com/learn/blog/proven-tactics-for-engaging-and-successful-welcome-emails/',
        'https://act-on.com/learn/blog/how-to-use-automated-customer-segmentation-for-better-results/',
        'https://act-on.com/learn/blog/demand-generation-101-7-tactics-for-generating-high-quality-leads/',
        'https://act-on.com/learn/blog/5-steps-to-increase-conversion-rates-with-account-based-marketing/',
        'https://act-on.com/learn/blog/the-new-rules-of-data-driven-marketing/',
        'https://act-on.com/learn/blog/capture-and-use-first-party-data/',
        'https://act-on.com/learn/blog/sales-and-marketing-alignment-why-it-matters/',
        'https://act-on.com/learn/blog/how-to-align-sales-and-marketing-on-strategy/',
        'https://act-on.com/learn/blog/integrate-sales-and-marketing-software-to-streamline-processes/',
        'https://act-on.com/learn/blog/retention-marketing-how-we-reached-400-customer-accounts/',
        'https://act-on.com/learn/blog/debunking-email-personalization-myths-part-1-of-2/'
    ]

    # urls = get_all_blogs()

    urls = filter_valid_urls(urls)


    # Create output directory if it doesn't exist
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Initialize the analyses structure
    analyses = {
        'metadata': {
            'analysis_date': datetime.now().isoformat(),
            'total_urls': len(urls),
            'successful_analyses': 0,
            'failed_analyses': 0
        },
        'analyses': {
            'blog': [],
            'webinar': [],
            'case-study': [],
            'resource': [],
            'other': []
        }
    }

    # Analyze each URL
    for url in urls:
        logger.info(f"\nAnalyzing URL: {url}")

        # Analyze the webpage
        analysis = analyzer.analyze_webpage(url)

        # Determine content type and update statistics
        content_type = analyzer.get_content_type(url)

        if 'error' in analysis:
            analyses['metadata']['failed_analyses'] += 1
        else:
            analyses['metadata']['successful_analyses'] += 1

        # Add analysis to appropriate category
        analyses['analyses'][content_type].append(analysis)

        logger.info(f"Analysis completed for: {url}")

    # Create output filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = 'blog.json'
    output_path = os.path.join(output_dir, filename)

    # Write analyses to JSON file
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(analyses, f, indent=2, ensure_ascii=False)

    logger.info(f"\nAnalysis complete:")
    logger.info(f"Total URLs processed: {analyses['metadata']['total_urls']}")
    logger.info(f"Successful analyses: {analyses['metadata']['successful_analyses']}")
    logger.info(f"Failed analyses: {analyses['metadata']['failed_analyses']}")
    logger.info(f"Results saved to: {output_path}")

if __name__ == "__main__":
    main()
