import re, json

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
