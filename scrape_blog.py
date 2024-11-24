import requests
from bs4 import BeautifulSoup
import json
from typing import Dict, List
import logging
import os
from datetime import datetime
from urllib.parse import urlparse
import sys
import ast


# Enhanced logging setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('scraper.log')
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
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5',
                # Let requests handle compression automatically
                'Accept-Encoding': 'gzip, deflate',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
                'Cache-Control': 'max-age=0'
            }

            # Make request with automatic content decoding
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()

            # Debug response
            logger.info(f"Response status code: {response.status_code}")
            logger.info(f"Response headers: {dict(response.headers)}")
            logger.info(f"Response encoding: {response.encoding}")

            # Get decoded text content
            html_content = response.text

            # Debug the raw HTML
            logger.info("First 1000 characters of decoded HTML:")
            logger.info(html_content[:1000])

            # Create soup with lxml parser
            soup = BeautifulSoup(html_content, 'lxml')

            # Debug soup structure
            logger.info(f"Found <head> tag: {soup.find('head') is not None}")
            head = soup.find('head')
            if head:
                logger.info("Head contents:")
                logger.info(head.prettify()[:1000])

                # Debug script tags
                scripts = head.find_all('script')
                logger.info(f"Found {len(scripts)} script tags in head")
                for i, script in enumerate(scripts):
                    logger.info(f"Script {i} type: {script.get('type')}")
                    logger.info(f"Script {i} class: {script.get('class')}")
                    if script.get('type') == 'application/ld+json':
                        logger.info(f"JSON-LD content: {script.string[:200]}")

            analysis = {
                'url': url,
                'analysis_timestamp': datetime.now().isoformat(),
                'basic_info': self.get_basic_info(soup),
                'seo_analysis': self.get_seo_analysis(soup),
                'multimedia_assessment': self.get_multimedia_assessment(soup),
                'content': self.get_content(soup),
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

    def get_basic_info(self, soup: BeautifulSoup) -> Dict:
        """Extract basic information about the blog post."""
        basic_info = {
            'title': '',
            'publication_date': '',
            'modified_date': '',
            'url': '',
            'description': ''
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
            'total_image_count': 0
        }

        # Find header/featured image
        featured_img = soup.find('div', class_='elementor-widget-theme-post-featured-image')
        if featured_img:
            img = featured_img.find('img')
            if img and not self.is_logo_image(img):
                multimedia['header_image'] = {
                    'src': img.get('src', ''),
                    'alt': img.get('alt', ''),
                    'width': img.get('width', ''),
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

        # Extract main content
        main_content = soup.find('div', class_='elementor-widget-theme-post-content')
        if main_content:
            for element in main_content.find_all(['h2', 'h3', 'h4', 'p', 'ul', 'ol', 'img', 'figure', 'div']):
                if element.name in ['h2', 'h3', 'h4']:
                    content.append(f"\n{element.name.upper()}: {get_spaced_text(element)}")
                elif element.name == 'p':
                    content.append(get_spaced_text(element))
                elif element.name in ['ul', 'ol']:
                    for li in element.find_all('li'):
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
                            src = img.get('src', '')
                            if src and alt_text:
                                content.append(f"\n[CONTENT IMAGE: {alt_text}]\nSource: {src}\n")

            return '\n\n'.join(content)
        return ""

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

def main():
    analyzer = BlogAnalyzer()
    urls = [
        "https://act-on.com/learn/blog/manufacturing-industry-slow-to-adopt-emerging-digital-marketing-software/",
        "https://act-on.com/learn/blog/retention-marketing-how-we-reached-400-customer-accounts/",
        "https://act-on.com/learn/data-sheets/advanced-crm-mapping/",
        "https://act-on.com/learn/blog/how-and-why-you-should-calculate-customer-lifetime-value-clv/",
        "https://act-on.com/learn/blog/pipeline-generation-face-economic-headwinds-and-win/",
        "https://act-on.com/learn/blog/what-is-customer-marketing-2/"
    ]

    # urls = get_all_blogs()



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
