import os
from batch_processor import BatchProcessor, style_excel_file


def main():
    # Define file paths
    json_file = '../output/all.json'
    excel_file = '../output/analysis_output.xlsx'

    # Ensure output directory exists
    os.makedirs(os.path.dirname(excel_file), exist_ok=True)

    # Initialize processor
    processor = BatchProcessor(
        json_file=json_file,
        excel_file=excel_file
    )

    try:
        # Process in batches
        start_index = 1  # Start from article 0
        batch_size = 1  # Process 100 articles at a time

        last_index = processor.process_batch(start_index, batch_size)

        if last_index is not None:
            print(f"Successfully processed articles up to index {last_index}")

            # Add a delay before styling
            import time
            time.sleep(2)

            # Only apply styling if this is the final batch
            try:
                style_excel_file(excel_file)
                print("Styling applied successfully")
            except Exception as e:
                print(f"Error applying styles: {e}")
        else:
            print("Error occurred during processing")

    except Exception as e:
        print(f"Error in main process: {e}")


if __name__ == "__main__":
    main()