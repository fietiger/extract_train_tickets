import pdfplumber
import pandas as pd
import os
import sys
import re

def clean_text(text):
    if not text:
        return ""
    return str(text).replace('\n', ' ').strip()

def process_didi_pdfs(input_dir, output_file):
    if not os.path.exists(input_dir):
        print(f"Error: Directory '{input_dir}' does not exist.")
        return

    all_trips = []
    pdf_files = [f for f in os.listdir(input_dir) if f.endswith('.pdf') and '行程报销单' in f]
    
    if not pdf_files:
        print(f"No Didi reimbursement PDF files found in '{input_dir}'.")
        return

    for file_name in pdf_files:
        file_path = os.path.join(input_dir, file_name)
        print(f"Processing: {file_name}")
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table:
                            continue
                        
                        header_found = False
                        header = []
                        for row in table:
                            clean_row = [clean_text(cell) for cell in row]
                            
                            if "上车时间" in clean_row:
                                header_found = True
                                header = clean_row
                                continue
                            
                            if header_found:
                                if not any(clean_row) or "合计" in "".join(clean_row):
                                    header_found = False
                                    continue
                                
                                try:
                                    city_idx = header.index("城市")
                                    time_idx = header.index("上车时间")
                                except (ValueError, NameError):
                                    city_idx = 3
                                    time_idx = 2
                                
                                # 1. Remove Day of Week (e.g. "11-09 08:25 周日" -> "11-09 08:25")
                                time_val = clean_row[time_idx]
                                time_val = re.split(r'\s*周\s*[一二三四五六日]\s*', time_val)[0].strip()
                                clean_row[time_idx] = time_val
                                
                                # 2. Remove spaces in City (e.g. "武汉 市" -> "武汉市")
                                city_val = clean_row[city_idx]
                                clean_row[city_idx] = city_val.replace(" ", "")
                                
                                all_trips.append(clean_row + [file_name])
                                
        except Exception as e:
            print(f"Error processing {file_name}: {e}")

    if all_trips:
        columns = header + ["来源文件"] if header else None
        df = pd.DataFrame(all_trips, columns=columns)
        df = df.dropna(how='all')
        
        df.to_excel(output_file, index=False)
        print(f"Success! Saved {len(df)} trips to: {output_file}")
    else:
        print("No valid trip info extracted.")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python process_didi.py <input_dir> <output_xlsx>")
        sys.exit(1)
    
    input_dir = sys.argv[1]
    output_xlsx = sys.argv[2]
    process_didi_pdfs(input_dir, output_xlsx)
