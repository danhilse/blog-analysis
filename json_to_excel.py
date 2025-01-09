import json
import pandas as pd
from datetime import datetime


def process_json_to_excel():
    # Read the JSON file
    with open('output/processed.json', 'r') as file:
        data = json.load(file)

    # Create a list to store the processed records
    processed_records = []

    # Process each record
    for key, value in data.items():
        record = {
            'title': value.get('title', ''),
            'url': value.get('url', ''),
            'publication_date': value.get('publication_date', ''),
            'processed_timestamp': value.get('processed_timestamp', ''),
            'red flags': ', '.join([str(match) for match in value.get('red_flags', {}).get('matches', [])]),
            'category': value.get('category', ''),
            'ai category': value.get('ai_category', ''),
            'ai category reason': value.get('ai_category_reasoning', ''),
            'pre sorted use case': value.get('pre sorted use case', ''),
            'use case': value.get('use_case', ''),
            'use case reasoning': value.get('use_case_reasoning', ''),
            'use case alt': value.get('use_case_alt', ''),
            'CMO Priority': value.get('cmoPriority', ''),
            'Get/Keep/Grow': value.get('getKeepGrow', ''),
            'Use Case Type 2': value.get('use_case_type_2', ''),
            'Use Case Reasoning Type 2': value.get('use_case_reasoning_type_2', ''),
            'CMO Priority Type 2': value.get('cmoPriority_type_2', ''),
            'Get/Keep/Grow Type 2': value.get('getKeepGrow_type_2', '')
        }

        processed_records.append(record)

    # Convert to DataFrame
    df = pd.DataFrame(processed_records)

    # Convert timestamps to a more readable format
    for col in ['publication_date', 'processed_timestamp']:
        df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d %H:%M:%S')

    # Write to Excel
    output_file = 'output/processed.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Successfully processed {len(processed_records)} records to {output_file}")


if __name__ == "__main__":
    process_json_to_excel()