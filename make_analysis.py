import json
from textstat import textstat
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
import pandas as pd
from datetime import datetime
from urllib.parse import urlparse
import textstat
import os

def clean_content(json_input):
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

    # Remove image markers like [CONTENT IMAGE: ...]
    content = re.sub(r'\[CONTENT IMAGE:.*?\]', '', content)

    # Remove source lines like Source: https://...
    content = re.sub(r'Source:\s*https?://\S+', '', content)

    # Convert 'H2:' headers to Markdown '## '
    content = re.sub(r'H2:\s*', '## ', content)

    # Normalize line breaks and remove excessive whitespace
    content = re.sub(r'\n{3,}', '\n\n', content)  # Replace 3+ newlines with 2
    content = re.sub(r'\n\s*\n', '\n\n', content)  # Ensure double newlines for paragraphs
    content = re.sub(r'[ \t]+', ' ', content)      # Replace multiple spaces/tabs with single space

    # Remove any remaining unwanted brackets or markdown artifacts
    content = re.sub(r'\[.*?\]', '', content)

    # Trim leading and trailing whitespace
    content = content.strip()

    return content

def calculate_word_count(content):
    clean_text = clean_content(content)
    return len(clean_text.split())


def add_no_match_highlighting(ws):
    """
    Adds conditional highlighting to mark cells containing 'No Clear Match', 'NONE',
    or similar values in red.

    Parameters:
        ws (openpyxl.worksheet.worksheet.Worksheet): The active worksheet
    """
    # Get header row values
    header_row = [cell.value for cell in ws[2]]  # Headers are in row 2

    # Columns to check for "No Clear" values
    target_columns = [
        'Primary Category',
        'Solution Topic',
        'Use Case',
        'Customer Journey Stage',
        'CMO Priority',
        'Marketing Activity Type',
        'Target Audience'
    ]

    # Get column indices for target columns
    column_indices = {col: header_row.index(col) + 1 for col in target_columns if col in header_row}

    # Values to highlight in red
    highlight_values = [
        'No Clear Match',
        'NONE',
        'No Clear Topic',
        'No Clear Activity Type',
        'No Clear Audience'
    ]

    # Red fill for matching cells
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    white_font = Font(color='FFFFFF', bold=True)

    # Apply highlighting to matching cells
    for row in range(3, ws.max_row + 1):  # Start from data rows
        for col_name, col_idx in column_indices.items():
            cell = ws.cell(row=row, column=col_idx)
            if cell.value in highlight_values:
                cell.fill = red_fill
                cell.font = white_font

def style_excel_file(filename):
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    # Extract title from first row and remove it
    title_cell = ws['A2'].value  # Save title value

    # Define brand colors with lighter shades for subheaders
    section_colors = {
        'Title': {
            'header': '#193661',  # Dark blue
            'subheader': '#C1C9D4'  # Lighter blue
        },
        'Basic Information': {
            'header': '#00babe',  # Teal
            'subheader': '#E5F9F9'  # Lighter teal
        },
        'Quality & Brand Fit': {
            'header': '#e34e64',  # Salmon
            'subheader': '#FFEFEF'  # Lighter salmon
        },
        'Tone & Voice': {
            'header': '#193661',  # Dark blue
            'subheader': '#C1C9D4'  # Lighter blue
        },
        'SEO Analysis': {
            'header': '#00babe',  # Teal
            'subheader': '#E5F9F9'  # Lighter teal
        },
        'Multimedia Assessment': {
            'header': '#e34e64',  # Salmon
            'subheader': '#FFEFEF'  # Lighter salmon
        },
        'Content Categorization': {
            'header': '#193661',  # Dark blue
            'subheader': '#C1C9D4'  # Lighter blue
        },
        'Performance Metrics': {
            'header': '#00babe',  # Teal
            'subheader': '#E5F9F9'  # Lighter teal
        },
        'Cost Analysis': {            # Added missing section
            'header': '#e34e64',      # Salmon
            'subheader': '#FFEFEF'    # Lighter salmon
        }
    }

    # Define sections and their ranges
    sections = {
        'Title': (1, 1),
        'Basic Information': (2, 7),  # Updated to include Modified Date
        'Quality & Brand Fit': (8, 10),
        'Tone & Voice': (11, 16),
        'SEO Analysis': (17, 26),  # Updated for split H2/H3
        'Multimedia Assessment': (27, 34),
        'Content Categorization': (35, 41),
        'Performance Metrics': (42, 47),
        'Cost Analysis': (48, 48)  # New section
    }

    # Freeze top two rows and first column
    ws.freeze_panes = 'B3'

    # Set specific column widths
    ws.column_dimensions['B'].width = 15  # URL column
    for col in range(1, ws.max_column + 1):
        if get_column_letter(col) != 'B':  # All non-URL columns
            ws.column_dimensions[get_column_letter(col)].width = 30

    # Style metric headers (second row)
    row2 = ws[1]  # Get the second row (index 1)
    for col_idx, cell in enumerate(row2, 1):
        # Find which section this column belongs to
        for section, (start, end) in sections.items():
            if start <= col_idx <= end:
                subheader_color = section_colors[section]['subheader'].replace('#', '')
                cell.fill = PatternFill(start_color=subheader_color,
                                      end_color=subheader_color,
                                      fill_type='solid')
                cell.font = Font(bold=True, color='444444')
                cell.alignment = openpyxl.styles.Alignment(horizontal='center',
                                                         vertical='center',
                                                         wrap_text=True)
                break

    # Add and style section headers
    ws.insert_rows(1)
    current_col = 1
    for section, (start, end) in sections.items():
        if current_col <= end:
            cell = ws.cell(row=1, column=current_col)
            cell.value = section
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color=section_colors[section]['header'].replace('#', ''),
                                  end_color=section_colors[section]['header'].replace('#', ''),
                                  fill_type='solid')
            cell.alignment = openpyxl.styles.Alignment(horizontal='center',
                                                     vertical='center',
                                                     wrap_text=True)

            if start != end:
                ws.merge_cells(start_row=1, start_column=current_col, end_row=1, end_column=end)
            current_col = end + 1

    # Get column indices for conditional formatting
    header_row = [cell.value for cell in ws[2]]  # Get header row values
    word_count_col = header_row.index('Word Count') + 1
    reading_level_col = header_row.index('Reading Level (Gunning Fog)') + 1
    header_image_width_col = header_row.index('Header Image Width') + 1

    # Add conditional formatting for Word Count and Reading Level
    ws.conditional_formatting.add(
        f'{get_column_letter(word_count_col)}3:{get_column_letter(word_count_col)}{ws.max_row}',
        ColorScaleRule(
            start_type='min',
            start_color='FF0000',  # Red
            mid_type='percentile',
            mid_value=50,
            mid_color='FFFF00',  # Yellow
            end_type='max',
            end_color='00FF00'  # Green
        )
    )

    ws.conditional_formatting.add(
        f'{get_column_letter(reading_level_col)}3:{get_column_letter(reading_level_col)}{ws.max_row}',
        ColorScaleRule(
            start_type='min',
            start_color='FF0000',  # Red
            mid_type='percentile',
            mid_value=50,
            mid_color='FFFF00',  # Yellow
            end_type='max',
            end_color='00FF00'  # Green
        )
    )

    # Add conditional formatting for Header Image Width
    for row in range(3, ws.max_row + 1):
        cell = ws.cell(row=row, column=header_image_width_col)
        try:
            width = int(cell.value)
            if width >= 1200:
                cell.fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')  # Green
            elif width >= 800:
                cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # Yellow
            else:
                cell.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')  # Red
        except (ValueError, TypeError):
            pass

    # Format data cells (including URL special formatting)
    for row in ws.iter_rows(min_row=3):  # Start from data rows
        for cell in row:
            # Special formatting for URL column
            if cell.column == 2:  # URL column (B)
                cell.alignment = openpyxl.styles.Alignment(horizontal='left',
                                                         vertical='center',
                                                         wrap_text=False)
                cell.font = Font(color='0563C1', underline='single')
                if cell.value and isinstance(cell.value, str):
                    cell.hyperlink = cell.value
            else:
                cell.alignment = openpyxl.styles.Alignment(horizontal='left',
                                                         vertical='center',
                                                         wrap_text=True)
                cell.font = Font(color='444444')

    # Add alternating row colors for better readability
    for row in range(3, ws.max_row + 1):
        fill_color = 'F7F9FB' if row % 2 == 0 else 'FFFFFF'
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if not cell.fill.start_color.rgb:  # Only apply if no other fill color is present
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    # Add thin borders to all cells
    thin_border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='thin', color='E3E3E3'),
        right=openpyxl.styles.Side(style='thin', color='E3E3E3'),
        top=openpyxl.styles.Side(style='thin', color='E3E3E3'),
        bottom=openpyxl.styles.Side(style='thin', color='E3E3E3')
    )

    # Add the new highlighting for "No Clear Match" values
    add_no_match_highlighting(ws)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border

    cost_col = header_row.index('API Cost') + 1
    for row in range(3, ws.max_row + 1):
        cell = ws.cell(row=row, column=cost_col)
        cell.number_format = '$#,##0.00000'


    # Restore title in its new position
    ws['A3'] = title_cell

    wb.save(filename)

def process_blog_data(data, output_file):
    # Create DataFrame
    df = create_blog_audit_df(data)

    # Save to Excel
    df.to_excel(output_file, index=False)

    # Apply styling
    style_excel_file(output_file)

    return df

def load_yoast_keywords():
    """Load Yoast keywords data and create URL to keyword mapping"""
    try:
        # Read the Yoast keywords Excel file
        yoast_df = pd.read_excel('resources/yoast-blog-keywords.xlsx', header=None)

        # Create a dictionary mapping URLs to keywords
        url_to_keyword = {}
        for _, row in yoast_df.iterrows():
            url = row[4]  # URL is in column 5 (index 4)
            keyword = row[2]  # Keyword is in column 3 (index 2)
            if pd.notna(url) and pd.notna(keyword):
                url_to_keyword[url.strip()] = keyword.strip()

        return url_to_keyword
    except Exception as e:
        print(f"Error loading Yoast keywords: {e}")
        return {}

def import_performance_data():
    """
    Imports performance data from resources/performance.xlsx and processes it for integration
    with blog audit data.
    """
    try:
        file_path = os.path.join('resources', 'performance.xlsx')
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

from ai_analysis import analyze_content_categorization, cost_tracker

def create_blog_audit_df(articles_data):
    """
    Creates a DataFrame containing audit information for multiple articles.
    Includes performance metrics from Google Analytics data.
    """
    # Set up logging

    url_to_keyword = load_yoast_keywords()
    all_articles = []

    # Import performance data
    performance_df = import_performance_data()

    for article in articles_data:
        cost_tracker.reset()  # Reset cost tracker for new article

        # Safely extract 'content'
        content = article.get('content', '')
        clean_text = clean_content(content)

        # Safely extract 'basic_info'
        basic_info = article.get('basic_info', {})
        url = basic_info.get('url', 'No URL')
        title = basic_info.get('title', 'No Title')
        publication_date_str = basic_info.get('publication_date')
        modified_date_str = basic_info.get('modified_date')

        # Extract target keyword
        target_keyword = url_to_keyword.get(url, 'Not Found')

        # Extract slug for performance data matching
        slug = urlparse(url).path.strip('/').split('/')[-1] if url != 'No URL' else 'No Slug'

        # Safely extract 'multimedia_assessment'
        multimedia = article.get('multimedia_assessment', {})
        header_image = multimedia.get('header_image') or {}

        # Safeguard against missing 'width' and 'height'
        try:
            header_width = int(header_image.get('width', 0))
        except (TypeError, ValueError):
            header_width = 0

        try:
            header_height = int(header_image.get('height', 0))
        except (TypeError, ValueError):
            header_height = 0

        header_src = header_image.get('src', 'No Source')
        header_alt = header_image.get('alt', 'No Alt Text')

        # Safely extract 'content_images'
        content_images = multimedia.get('content_images', [])
        content_widths = []
        for img in content_images:
            try:
                width = int(img.get('width', 0))
                content_widths.append(width)
            except (TypeError, ValueError):
                content_widths.append(0)

        min_content_width = min(content_widths) if content_widths else 0

        # Get performance metrics if available
        performance_metrics = {
            'Total Views': 0,
            'Total Users': 0,
            'Total Sessions': 0,
            'Engagement Rate': 0.0,
            'Average Time on Page': 0.0,
            'Bounce Rate': 0.0
        }

        if performance_df is not None and slug in performance_df.index:
            metrics = performance_df.loc[slug]
            performance_metrics = {
                'Total Views': metrics.get('Views', 0),
                'Total Users': metrics.get('Total users', 0),
                'Total Sessions': metrics.get('Sessions', 0),
                'Engagement Rate': metrics.get('Engagement rate', 0.0),
                'Average Time on Page': metrics.get('Average session duration', 0.0),
                'Bounce Rate': metrics.get('Bounce rate', 0.0)
            }

        # Safely extract 'seo_analysis'
        seo_analysis = article.get('seo_analysis', {})
        meta_description = seo_analysis.get('meta_description', {})
        meta_present = meta_description.get('present', False)

        headings = seo_analysis.get('headings', {})
        h1_present = headings.get('h1_present', False)
        h2_count = headings.get('h2_count', 0)
        h3_count = headings.get('h3_count', 0)

        # Parse dates safely
        def parse_date(date_str):
            if date_str:
                try:
                    return datetime.strptime(date_str.split('T')[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                except ValueError:
                    return 'No Date'
            return 'No Date'

        publication_date = parse_date(publication_date_str)
        modified_date = parse_date(modified_date_str)

        categorization = analyze_content_categorization(clean_text)

        # print(categorization)

        # Extract other necessary fields with defaults
        article_data = {
            # Basic Information
            'Title': title,
            'URL': url,
            'Slug': slug,
            'Publication Date': publication_date,
            'Modified Date': modified_date,
            'Word Count': calculate_word_count(clean_text),
            'Reading Level (Gunning Fog)': round(textstat.gunning_fog(clean_text), 1) if clean_text else 0.0,

            # Quality & Brand Fit
            'Overall Quality Score': 'TBD',
            'Brand Alignment Score': 'TBD',
            'Quality Notes/Recommendations': 'TBD',

            # Tone & Voice
            'Challenger Percentage': 'TBD',
            'Supportive Percentage': 'TBD',
            'Natural/Conversational Score': 'TBD',
            'Authentic/Approachable Score': 'TBD',
            'Gender-Neutral/Inclusive Score': 'TBD',
            'Tone Notes/Recommendations': 'TBD',

            # SEO Analysis
            'Current Target Keyword': target_keyword,
            'Keyword Density': 'TBD',
            'Keyword Integration Score': 'TBD',
            'Meta Description Present': 'Yes' if meta_present else 'No',
            'Meta Description Quality Score': 'TBD',
            'H1 Tag Present': 'Yes' if h1_present else 'No',
            'H2 Tags': h2_count,
            'H3 Tags': h3_count,
            'Recommended New Keywords': 'TBD',
            'SEO Notes/Recommendations': 'TBD',

            # Multimedia Assessment
            'Number of Images': multimedia.get('total_image_count', 0),
            'Header Image Width': header_width,
            'Header Image Height': header_height,  # Added for completeness
            'Header Image Src': header_src,
            'Header Image Alt': header_alt,
            'Minimum Content Image Width': min_content_width,
            'Unmodified Stock Images Flag': 'TBD',
            'Outdated Widgets Flag': 'TBD',

            # Update Content Categorization section with AI results
            'Primary Category': categorization["Primary Category"],
            'Solution Topic': categorization["Solution Topic"],
            'Use Case': categorization["Use Case"],
            'Customer Journey Stage': categorization["Customer Journey Stage"],
            'CMO Priority': categorization["CMO Priority"],
            'Marketing Activity Type': categorization["Marketing Activity Type"],
            'Target Audience': categorization["Target Audience"],

            # Performance Metrics (from Google Analytics)
            **performance_metrics,

            'API Cost': f"${cost_tracker.cost:.4f}",  # Add cost as last field
        }

        all_articles.append(article_data)

    df = pd.DataFrame(all_articles)

    # Convert numeric columns to appropriate types
    numeric_columns = ['Total Views', 'Total Users', 'Total Sessions',
                      'Engagement Rate', 'Average Time on Page', 'Bounce Rate']
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Format time-based metrics
    df['Average Time on Page'] = df['Average Time on Page'].round(2)
    df['Engagement Rate'] = df['Engagement Rate'].round(4)
    df['Bounce Rate'] = df['Bounce Rate'].round(4)

    return df

def process_content_data(json_data, output_file):
    """
    Process multiple content pieces from the JSON structure and create an Excel report.

    Parameters:
        json_data (dict): Parsed JSON data containing multiple content analyses
        output_file (str): Path to save the Excel output
    """
    all_content = []

    # Process each content type
    for content_type, content_list in json_data['analyses'].items():
        # Skip empty content types
        if not content_list:
            continue

        # Add each piece of content to the list
        for content in content_list:
            content['content_type'] = content_type  # Add content type to the data
            all_content.append(content)

    # Create DataFrame with all content
    df = create_blog_audit_df(all_content)

    # Save to Excel
    df.to_excel(output_file, index=False)

    # Apply styling
    style_excel_file(output_file)

    return df

# Example usage
if __name__ == "__main__":
    input_file = 'output/blog.json'
    output_file = 'output/blog_audit.xlsx'

    # Read the JSON file
    with open(input_file, 'r') as file:
        data = json.load(file)

    # Process the data
    df = process_content_data(data, output_file)
    print(f"Excel file created: {output_file}")
