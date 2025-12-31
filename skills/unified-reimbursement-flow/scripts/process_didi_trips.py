import pdfplumber
import pandas as pd
import os
import sys
import re

def clean_text(text):
    if not text: return ""
    return str(text).replace('\n', ' ').strip()

def process_didi_pdfs(input_dir, output_file):
    all_trips = []
    pdf_files = [f for f in os.listdir(input_dir) if f.endswith('.pdf') and '行程报销单' in f]
    for file_name in pdf_files:
        file_path = os.path.join(input_dir, file_name)
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table: continue
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
                            except:
                                city_idx, time_idx = 3, 2
                            time_val = re.split(r'\s*周\s*[一二三四五六日]\s*', clean_row[time_idx])[0].strip()
                            clean_row[time_idx] = time_val
                            clean_row[city_idx] = clean_row[city_idx].replace(" ", "")
                            all_trips.append(clean_row + [file_name])
    if all_trips:
        columns = header + ["来源文件"] if header else None
        pd.DataFrame(all_trips, columns=columns).to_excel(output_file, index=False)

if __name__ == "__main__":
    process_didi_pdfs(sys.argv[1], sys.argv[2])
