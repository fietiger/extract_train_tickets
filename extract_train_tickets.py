# -*- coding: utf-8 -*-
"""
Train Ticket Information Extractor
Extracts information from train ticket PDF files and saves to CSV
"""

import os
import re
import csv
from datetime import datetime
import pdfplumber
from pathlib import Path

class TrainTicketExtractor:
    def __init__(self):
        self.extracted_data = []
        
    def extract_text_from_pdf(self, pdf_path):
        """Extract text content from PDF file"""
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
        """Parse train ticket information from extracted text"""
        # Initialize default values
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
            'arrival_time': '',
            'seat_type': '',
            'seat_number': ''
        }
        
        # Clean text - remove extra spaces and newlines
        text = re.sub(r'\s+', ' ', text).strip()
        lines = text.split('\n')
        
        # Extract invoice number
        invoice_pattern = r'发票号码:(\d+)'
        invoice_match = re.search(invoice_pattern, text)
        if invoice_match:
            ticket_info['invoice_number'] = invoice_match.group(1)
        
        # Extract date (travel date, not invoice date)
        for line in lines:
            # Look for travel date pattern like "2025年08月31日"
            date_match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)\s*\d{1,2}:\d{2}', line)
            if date_match:
                chinese_date = date_match.group(1)
                # Convert Chinese date format to ISO format (YYYY-MM-DD)
                date_parts = re.findall(r'\d+', chinese_date)
                if len(date_parts) >= 3:
                    year, month, day = date_parts[0], date_parts[1].zfill(2), date_parts[2].zfill(2)
                    ticket_info['date'] = f"{year}-{month}-{day}"
                else:
                    ticket_info['date'] = chinese_date  # Fallback to original if parsing fails
                break
        
        # Extract train number and stations from the same line
        for line in lines:
            # Pattern like "上海虹桥 D935 深圳北"
            train_station_match = re.search(r'([^\s]+)\s+([GDKTZCY]\d{1,4})\s+([^\s]+)', line)
            if train_station_match:
                ticket_info['departure_station'] = train_station_match.group(1)
                ticket_info['train_number'] = train_station_match.group(2)
                ticket_info['arrival_station'] = train_station_match.group(3)
                break
        
        # Extract departure time and seat info from the same line
        for line in lines:
            # Pattern like "2025年08月31日 20:05开 11车034号上铺 动卧"
            time_seat_match = re.search(r'\d{4}年\d{1,2}月\d{1,2}日\s+(\d{1,2}:\d{2})开\s+(.+)', line)
            if time_seat_match:
                ticket_info['departure_time'] = time_seat_match.group(1)
                seat_info = time_seat_match.group(2).strip()
                
                # Parse seat information
                # Pattern like "11车034号上铺 动卧" or "03车08F号 二等座" or "12车无座 二等座"
                seat_match = re.search(r'(\d+车[^\s]*)\s*([^\s]*)', seat_info)
                if seat_match:
                    ticket_info['seat_number'] = seat_match.group(1)
                    if seat_match.group(2):
                        ticket_info['seat_type'] = seat_match.group(2)
                else:
                    # Try to extract seat type only
                    seat_type_match = re.search(r'(一等座|二等座|硬座|软座|硬卧|软卧|商务座|特等座|动卧)', seat_info)
                    if seat_type_match:
                        ticket_info['seat_type'] = seat_type_match.group(1)
                break
        
        # Extract price
        for line in lines:
            # Pattern like "￥720.00"
            price_match = re.search(r'￥(\d+\.?\d*)', line)
            if price_match:
                ticket_info['price'] = price_match.group(1)
                break
        
        # Extract passenger name
        for line in lines:
            # Pattern like "4127281981****2515 程文涛"
            name_match = re.search(r'\d+\*+\d+\s+([^\s]{2,4})', line)
            if name_match:
                ticket_info['passenger_name'] = name_match.group(1)
                break
        
        return ticket_info
    
    def process_pdf_files(self, directory_path):
        """Process all PDF files in the directory"""
        pdf_files = list(Path(directory_path).glob("*.pdf"))
        
        if not pdf_files:
            print("No PDF files found in the directory.")
            return
        
        print(f"Found {len(pdf_files)} PDF files to process...")
        
        for pdf_file in pdf_files:
            print(f"Processing: {pdf_file.name}")
            
            # Extract text from PDF
            text = self.extract_text_from_pdf(pdf_file)
            
            if text:
                # Parse ticket information
                ticket_info = self.parse_train_ticket_info(text, pdf_file.name)
                self.extracted_data.append(ticket_info)
                print(f"  - Extracted data for {pdf_file.name}")
            else:
                print(f"  - Failed to extract text from {pdf_file.name}")
    
    def save_to_csv(self, output_file="train_tickets_extracted.csv"):
        """Save extracted data to CSV file with deduplication based on invoice number"""
        if not self.extracted_data:
            print("No data to save.")
            return
        
        # Remove duplicates based on invoice number
        unique_data = {}
        for ticket in self.extracted_data:
            invoice_num = ticket.get('invoice_number', '')
            if invoice_num and invoice_num not in unique_data:
                unique_data[invoice_num] = ticket
            elif not invoice_num:  # Keep tickets without invoice number
                # Use filename as fallback key for tickets without invoice number
                unique_data[ticket['filename']] = ticket
        
        # Convert back to list
        deduplicated_data = list(unique_data.values())
        
        # Add route column to each ticket
        for ticket_data in deduplicated_data:
            departure = ticket_data.get('departure_station', '')
            arrival = ticket_data.get('arrival_station', '')
            if departure and arrival:
                ticket_data['route'] = f"{departure} → {arrival}"
            else:
                ticket_data['route'] = ""
        
        # Define CSV headers (with route column added)
        headers = [
            'filename', 'invoice_number', 'date', 'train_number', 'departure_station', 
            'arrival_station', 'route', 'departure_time', 'arrival_time', 
            'passenger_name', 'seat_type', 'seat_number', 'price'
        ]
        
        try:
            with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=headers)
                writer.writeheader()
                
                for ticket_data in deduplicated_data:
                    # Format invoice number to prevent Excel from treating it as a number
                    formatted_ticket = ticket_data.copy()
                    if formatted_ticket.get('invoice_number'):
                        formatted_ticket['invoice_number'] = f'="{formatted_ticket["invoice_number"]}"'
                    writer.writerow(formatted_ticket)
            
            print(f"Data successfully saved to {output_file}")
            print(f"Total unique tickets: {len(deduplicated_data)} (removed {len(self.extracted_data) - len(deduplicated_data)} duplicates)")
            
        except Exception as e:
            print(f"Error saving to CSV: {str(e)}")
    
    def print_summary(self):
        """Print a summary of extracted data"""
        if not self.extracted_data:
            print("No data extracted.")
            return
        
        print("\n" + "="*50)
        print("EXTRACTION SUMMARY")
        print("="*50)
        
        for i, ticket in enumerate(self.extracted_data, 1):
            print(f"\nTicket {i}: {ticket['filename']}")
            print(f"  Invoice Number: {ticket['invoice_number']}")
            print(f"  Date: {ticket['date']}")
            print(f"  Train: {ticket['train_number']}")
            print(f"  Route: {ticket['departure_station']} → {ticket['arrival_station']}")
            print(f"  Time: {ticket['departure_time']} - {ticket['arrival_time']}")
            print(f"  Passenger: {ticket['passenger_name']}")
            print(f"  Seat: {ticket['seat_type']} {ticket['seat_number']}")
            print(f"  Price: ¥{ticket['price']}")

def main():
    """Main function"""
    # Get current directory
    current_dir = os.getcwd()
    print(f"Processing PDF files in: {current_dir}")
    
    # Create extractor instance
    extractor = TrainTicketExtractor()
    
    # Process all PDF files in current directory
    extractor.process_pdf_files(current_dir)
    
    # Print summary
    extractor.print_summary()
    
    # Save to CSV
    extractor.save_to_csv()

if __name__ == "__main__":
    main()