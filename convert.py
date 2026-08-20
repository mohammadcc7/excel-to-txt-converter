import os
import pandas as pd

input_dir = "input"
output_dir = "output"

os.makedirs(output_dir, exist_ok=True)
os.makedirs(input_dir, exist_ok=True)

if os.path.exists(input_dir):
    for file_name in os.listdir(input_dir):
        if file_name.endswith(('.xlsx', '.xls')):
            file_path = os.path.join(input_dir, file_name)
            base_name = os.path.splitext(file_name)[0]
            
            try:
                excel_file = pd.ExcelFile(file_path)
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    out_name = f"{base_name}_{sheet_name}.txt" if len(excel_file.sheet_names) > 1 else f"{base_name}.txt"
                    txt_path = os.path.join(output_dir, out_name)
                    df.to_csv(txt_path, sep='\t', index=False, encoding='utf-8-sig')
                    print(f"Converted: {file_name}")
            except Exception as e:
                print(f"Error {file_name}: {e}")
