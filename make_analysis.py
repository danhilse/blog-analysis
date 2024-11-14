import pandas as pd
import json
from textstat import textstat
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

def style_excel_file(filename):
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    # Extract title from first row and remove it from data
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
        }
    }

    # Define sections and their ranges
    sections = {
        'Title': (1, 1),
        'Basic Information': (2, 5),
        'Quality & Brand Fit': (6, 8),
        'Tone & Voice': (9, 14),
        'SEO Analysis': (15, 23),
        'Multimedia Assessment': (24, 28),
        'Content Categorization': (29, 35),
        'Performance Metrics': (36, 40)
    }

    # Freeze top two rows and first column
    ws.freeze_panes = 'B3'

    # Set specific column widths
    ws.column_dimensions['B'].width = 15  # URL column
    for col in range(1, ws.max_column + 1):
        if get_column_letter(col) != 'B':  # All non-URL columns
            ws.column_dimensions[get_column_letter(col)].width = 30

    # Style metric headers (second row) - Important to do this first!
    row2 = ws[1]  # Get the second row (index 1)
    for col_idx, cell in enumerate(row2, 1):
        # Find which section this column belongs to
        for section, (start, end) in sections.items():
            if start <= col_idx <= end:
                # Get the subheader color for this section
                subheader_color = section_colors[section]['subheader'].replace('#', '')
                # Apply the fill color and other styles
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
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')

    # Add thin borders to all cells
    thin_border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='thin', color='E3E3E3'),
        right=openpyxl.styles.Side(style='thin', color='E3E3E3'),
        top=openpyxl.styles.Side(style='thin', color='E3E3E3'),
        bottom=openpyxl.styles.Side(style='thin', color='E3E3E3')
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border

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


def create_blog_audit_df(articles_data):
    """
    Creates a DataFrame containing audit information for multiple articles.

    Parameters:
        articles_data (list): List of article dictionaries containing analysis data

    Returns:
        pandas.DataFrame: DataFrame containing audit information for all articles
    """
    all_articles = []

    # Process each article
    for article in articles_data:
        clean_text = clean_content(article['content'])

        article_data = {
            # Basic Information
            'Title': article['basic_info']['title'],
            'URL': article['basic_info']['url'],
            'Publication Date': datetime.strptime(article['basic_info']['publication_date'].split('T')[0], '%Y-%m-%d').strftime('%Y-%m-%d') if article['basic_info']['publication_date'] else 'No Date',
            'Word Count': calculate_word_count(clean_text),
            'Reading Level (Gunning Fog)': round(textstat.gunning_fog(clean_text), 1),

            # Quality & Brand Fit (placeholders)
            'Overall Quality Score': 'TBD',
            'Brand Alignment Score': 'TBD',
            'Quality Notes/Recommendations': 'TBD',

            # Tone & Voice (placeholders)
            'Challenger Percentage': 'TBD',
            'Supportive Percentage': 'TBD',
            'Natural/Conversational Score': 'TBD',
            'Authentic/Approachable Score': 'TBD',
            'Gender-Neutral/Inclusive Score': 'TBD',
            'Tone Notes/Recommendations': 'TBD',

            # SEO Analysis
            'Current Target Keyword': 'TBD',
            'Keyword Density': 'TBD',
            'Keyword Integration Score': 'TBD',
            'Meta Description Present': 'Yes' if article['seo_analysis']['meta_description']['present'] else 'No',
            'Meta Description Quality Score': 'TBD',
            'H1 Tag Present': 'Yes' if article['seo_analysis']['headings']['h1_present'] else 'No',
            'Number of H2/H3 Tags': f"H2: {article['seo_analysis']['headings']['h2_count']}, H3: {article['seo_analysis']['headings']['h3_count']}",
            'Recommended New Keywords': 'TBD',
            'SEO Notes/Recommendations': 'TBD',

            # Multimedia Assessment
            'Number of Images': article['multimedia_assessment']['total_image_count'],
            'Missing Images Flag': 'TBD',
            'Low Quality Images Flag': 'TBD',
            'Unmodified Stock Images Flag': 'TBD',
            'Outdated Widgets Flag': 'TBD',

            # Content Categorization (placeholders)
            'Primary Category': 'TBD',
            'Solution Topic': 'TBD',
            'Use Case': 'TBD',
            'Customer Journey Stage': 'TBD',
            'CMO Priority': 'TBD',
            'Marketing Activity Type': 'TBD',
            'Target Audience': 'TBD',

            # Performance Metrics (placeholders)
            'Total Views': 'TBD',
            'Average Time on Page': 'TBD',
            'Bounce Rate': 'TBD',
            'Conversion Rate': 'TBD',
            'Inbound Links Count': 'TBD'
        }

        all_articles.append(article_data)

    return pd.DataFrame(all_articles)

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
