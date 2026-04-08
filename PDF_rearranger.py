import yaml
import os
import sys
from pypdf import PdfReader, PdfWriter

def rearrange_pdf_via_config(yaml_path):
    # 1. Load the YAML configuration
    try:
        with open(yaml_path, 'r') as f:
            config = yaml.safe_load(f)
    except Exception as e:
        print(f"Error reading YAML file: {e}")
        return

    # 2. Extract Paths and Data
    input_pdf_path = config['settings']['input_path']
    output_pdf_path = config['settings']['output_path']
    items = config['extractions']

    # 3. Validation
    if not os.path.exists(input_pdf_path):
        print(f"Error: Input PDF not found at {input_pdf_path}")
        return

    # 4. Sort by Index
    items.sort(key=lambda x: x['index'])
    
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    
    print(f"Reading from: {input_pdf_path}")
    
    # 5. Process Pages
    for item in items:
        name = item.get('file_name', 'Unknown Doc')
        start = item['start_pg']
        end = item['end_pg']
        
        print(f"Processing Index {item['index']}: {name} (Pages {start}-{end})")
        
        try:
            for page_num in range(start - 1, end):
                writer.add_page(reader.pages[page_num])
        except IndexError:
            print(f"  [!] Warning: Range {start}-{end} exceeds PDF length. Skipping.")

    # 6. Save Final File
    try:
        with open(output_pdf_path, "wb") as f_out:
            writer.write(f_out)
        print(f"\nSuccess! Saved to: {output_pdf_path}")
    except Exception as e:
        print(f"Error saving output file: {e}")

if __name__ == "__main__":
    # Get the directory of the executable/script
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    config_location = os.path.join(base_path, "config.yaml")

    if not os.path.exists(config_location):
        print(f"Critical Error: 'config.yaml' must be in the same folder as this program.")
    else:
        rearrange_pdf_via_config(config_location)
    
    input("\nPress Enter to close this window...")