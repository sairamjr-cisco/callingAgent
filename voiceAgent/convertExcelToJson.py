import os
import json
import pandas as pd
import datetime

def convert_excel_to_json(excel_path, json_path=None):
    if not os.path.exists(excel_path):
        print(f"Error: File {excel_path} does not exist.")
        return

    # If output json path is not specified, construct it based on excel file name
    if json_path is None:
        json_path = os.path.splitext(excel_path)[0] + '.json'

    print(f"Reading Excel file: {excel_path}")
    excel_file = pd.ExcelFile(excel_path)
    sheet_names = excel_file.sheet_names
    print(f"Sheets found: {sheet_names}")

    data_to_serialize = {}
    
    for sheet_name in sheet_names:
        # Read the sheet, handling any potential empty/null issues
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        
        # Convert Timestamp and other non-serializable objects to string or appropriate type
        # pandas can output directly to dict/json using to_dict
        records = df.to_dict(orient='records')
        
        # We need a clean serialization function to handle special types like NaT, NaN, Timestamp, datetime
        def sanitize_value(val):
            if pd.isna(val):
                return None
            if isinstance(val, (pd.Timestamp, pd.Timedelta)):
                return str(val)
            if isinstance(val, (datetime.datetime, datetime.date, datetime.time)):
                if isinstance(val, (datetime.datetime, datetime.date)):
                    return val.strftime('%m/%d/%Y')
                return str(val)
            return val

        sanitized_records = []
        for record in records:
            sanitized_record = {str(k): sanitize_value(v) for k, v in record.items()}
            sanitized_records.append(sanitized_record)
            
        data_to_serialize[sheet_name] = sanitized_records

    # If there is only one sheet, simplify the structure to be just the list of that sheet's records
    if len(sheet_names) == 1:
        final_data = data_to_serialize[sheet_names[0]]
    else:
        final_data = data_to_serialize

    # Write to JSON file with indent for readability
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(final_data, f, ensure_ascii=False, indent=4)
        
    print(f"Successfully converted and saved to: {json_path}")

if __name__ == "__main__":
    # Scan the local 'input' directory for Excel files
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(base_dir, "input")
    
    if os.path.exists(input_dir):
        excel_files = [f for f in os.listdir(input_dir) if f.endswith((".xlsx", ".xls"))]
        if excel_files:
            # Prioritize files containing 'customer_data' or 'customer'
            target_file = None
            for f in excel_files:
                if "customer" in f.lower():
                    target_file = f
                    break
            if not target_file:
                target_file = excel_files[0]
            
            excel_file_path = os.path.join(input_dir, target_file)
            json_file_path = os.path.join(input_dir, "customer_data.json")
            
            # Ensure the output directory is exactly the input folder
            convert_excel_to_json(excel_file_path, json_file_path)
        else:
            print("No Excel files found in the input folder.")
    else:
        print("Input folder does not exist.")
