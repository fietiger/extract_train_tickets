# -*- coding: utf-8 -*-
import os
import re
import pandas as pd
import pdfplumber
from pathlib import Path
import sys

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
        
        train_station_match = re.search(r'([^\s]+)\s+([GDKTZCY]\d{1,4})\s+([^\s]+)', text)
        if train_station_match:
            ticket_info['departure_station'] = train_station_match.group(1)
            ticket_info['train_number'] = train_station_match.group(2)
            ticket_info['arrival_station'] = train_station_match.group(3)
        
        time_seat_match = re.search(r'\d{4}年\d{1,2}月\d{1,2}日\s+(\d{1,2}:\d{2})开\s+(.+)', text)
        if time_seat_match:
            ticket_info['departure_time'] = time_seat_match.group(1)
        
        # Try different patterns for price
        price_patterns = [
            r'￥\s*(\d+\.?\d*)',
            r'票价:\s*￥?\s*(\d+\.?\d*)',
            r'二等座\s*￥?\s*(\d+\.?\d*)',
            r'座\s*￥?\s*(\d+\.?\d*)'
        ]
        
        for pattern in price_patterns:
            price_match = re.search(pattern, text)
            if price_match:
                ticket_info['price'] = float(price_match.group(1))
                break



        
        name_match = re.search(r'\d+\*+\d+\s+([^\s]{2,4})', text)
        if name_match:
            ticket_info['passenger_name'] = name_match.group(1)
        
        return ticket_info
    
    def process_pdf_files(self, directory_path):
        pdf_files = list(Path(directory_path).rglob("*.pdf"))
        for pdf_file in pdf_files:
            text = self.extract_text_from_pdf(pdf_file)
            if text:
                ticket_info = self.parse_train_ticket_info(text, pdf_file.name)
                self.extracted_data.append(ticket_info)
    
    def save_to_xlsx(self, output_file):
        if not self.extracted_data: return
        df = pd.DataFrame(self.extracted_data)
        if 'invoice_number' in df.columns:
            df['invoice_number'] = df['invoice_number'].apply(lambda x: f'="{x}"' if x else "")
        df.to_excel(output_file, index=False)



if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "火车票"
    output = sys.argv[2] if len(sys.argv) > 2 else "火车票.xlsx"
    ext = TrainTicketExtractor()
    ext.process_pdf_files(target)
    ext.save_to_xlsx(output)
