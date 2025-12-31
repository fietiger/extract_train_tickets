# -*- coding: utf-8 -*-
import os
import re
import csv
import pandas as pd
from datetime import datetime
import pdfplumber
from pathlib import Path

class TrainTicketExtractor:
    def __init__(self):
        self.extracted_data = []
        
    def extract_text_from_pdf(self, pdf_path):
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
                return text
        except Exception as e:
            print(f"Error reading PDF {pdf_path}: {str(e)}")
            return ""
    
    def parse_train_ticket_info(self, text, filename):
        ticket_info = {
            'filename': filename,
            'invoice_number': '',
            'date': '',
            'departure_station': '',
            'arrival_station': '',
            'price': '',
            'passenger_name': '',
            'train_number': '',
            'departure_time': '',
            'seat_type': '',
            'seat_number': ''
        }
        text = re.sub(r'\s+', ' ', text).strip()
        lines = text.split('\n')
        
        invoice_match = re.search(r'发票号码:(\d+)', text)
        if invoice_match:
            ticket_info['invoice_number'] = invoice_match.group(1)
        
        date_match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)\s*\d{1,2}:\d{2}', text)
        if date_match:
            chinese_date = date_match.group(1)
            date_parts = re.findall(r'\d+', chinese_date)
            if len(date_parts) >= 3:
                year, month, day = date_parts[0], date_parts[1].zfill(2), date_parts[2].zfill(2)
                ticket_info['date'] = f"{year}-{month}-{day}"
            else:
                ticket_info['date'] = chinese_date
        
        train_station_match = re.search(r'([^\s]+)\s+([GDKTZCY]\d{1,4})\s+([^\s]+)', text)
        if train_station_match:
            ticket_info['departure_station'] = train_station_match.group(1)
            ticket_info['train_number'] = train_station_match.group(2)
            ticket_info['arrival_station'] = train_station_match.group(3)
        
        time_seat_match = re.search(r'\d{4}年\d{1,2}月\d{1,2}日\s+(\d{1,2}:\d{2})开\s+(.+)', text)
        if time_seat_match:
            ticket_info['departure_time'] = time_seat_match.group(1)
            seat_info = time_seat_match.group(2).strip()
            seat_match = re.search(r'(\d+车[^\s]*)\s*([^\s]*)', seat_info)
            if seat_match:
                ticket_info['seat_number'] = seat_match.group(1)
                if seat_match.group(2):
                    ticket_info['seat_type'] = seat_match.group(2)
            else:
                seat_type_match = re.search(r'(一等座|二等座|硬座|软座|硬卧|软卧|商务座|特等座|动卧)', seat_info)
                if seat_type_match:
                    ticket_info['seat_type'] = seat_type_match.group(1)
        
        price_match = re.search(r'￥(\d+\.?\d*)', text)
        if price_match:
            ticket_info['price'] = price_match.group(1)
        
        name_match = re.search(r'\d+\*+\d+\s+([^\s]{2,4})', text)
        if name_match:
            ticket_info['passenger_name'] = name_match.group(1)
        
        return ticket_info
    
    def process_pdf_files(self, directory_path):
        pdf_files = list(Path(directory_path).rglob("*.pdf"))
        if not pdf_files:
            print(f"No PDF files found in {directory_path}")
            return
        
        print(f"Found {len(pdf_files)} PDF files to process...")
        for pdf_file in pdf_files:
            print(f"Processing: {pdf_file}")
            text = self.extract_text_from_pdf(pdf_file)
            if text:
                ticket_info = self.parse_train_ticket_info(text, pdf_file.name)
                self.extracted_data.append(ticket_info)
                print(f"  - Extracted data for {pdf_file.name}")
            else:
                print(f"  - Failed to extract text from {pdf_file.name}")
    
    def save_to_xlsx(self, output_file="train_tickets_extracted.xlsx"):
        if not self.extracted_data:
            print("No data to save.")
            return
        
        unique_data = {}
        for ticket in self.extracted_data:
            invoice_num = ticket.get('invoice_number', '')
            if invoice_num and invoice_num not in unique_data:
                unique_data[invoice_num] = ticket
            elif not invoice_num:
                unique_data[ticket['filename']] = ticket
        
        deduplicated_data = list(unique_data.values())
        for ticket_data in deduplicated_data:
            departure = ticket_data.get('departure_station', '')
            arrival = ticket_data.get('arrival_station', '')
            ticket_data['route'] = f"{departure} → {arrival}" if departure and arrival else ""
        
        headers = [
            'filename', 'invoice_number', 'date', 'train_number', 'departure_station', 
            'arrival_station', 'route', 'departure_time', 
            'passenger_name', 'seat_type', 'seat_number', 'price'
        ]
        
        try:
            df = pd.DataFrame(deduplicated_data)
            for h in headers:
                if h not in df.columns:
                    df[h] = ""
            
            if 'invoice_number' in df.columns:
                df['invoice_number'] = df['invoice_number'].apply(lambda x: f'="{x}"' if x else "")
            if 'price' in df.columns:
                df['price'] = df['price'].apply(lambda x: f'={x}' if x else "")
            
            df = df[headers]
            df.to_excel(output_file, index=False)
            print(f"Data successfully saved to {output_file}")
        except Exception as e:
            print(f"Error saving to XLSX: {str(e)}")

def main():
    import sys
    # Default values
    target_dir = os.getcwd()
    output_file = "火车票汇总信息表.xlsx"
    
    # Handle command line arguments
    if len(sys.argv) > 1:
        target_dir = sys.argv[1]
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
        # Ensure it has .xlsx extension
        if not output_file.lower().endswith('.xlsx'):
            output_file += '.xlsx'
            
    print(f"Target dir: {target_dir}")
    print(f"Output file: {output_file}")
    
    extractor = TrainTicketExtractor()
    extractor.process_pdf_files(target_dir)
    extractor.save_to_xlsx(output_file)


if __name__ == "__main__":
    main()
