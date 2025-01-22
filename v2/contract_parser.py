"""
Enhanced Contract Parser v2.0
Extracts contract details with performance tracking and metrics
"""

import os
import PyPDF2
import re
import pandas as pd
from datetime import datetime
import time

class ContractParser:
    def __init__(self, directory_path):
        self.directory_path = directory_path
        self.metrics = {
            'start_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'processed_files': 0,
            'successful_extractions': 0
        }

    def extract_text_from_pdf(self, pdf_path):
        """Extract text from PDF with performance tracking."""
        start_time = time.time()
        file_size = os.path.getsize(pdf_path) / 1024

        try:
            with open(pdf_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                all_text = ''
                for page in pdf_reader.pages:
                    all_text += page.extract_text()
                
                processing_time = time.time() - start_time
                return all_text, {
                    'file_size_kb': round(file_size, 2),
                    'processing_time_seconds': round(processing_time, 3),
                    'processing_speed_kb_per_sec': round(file_size / processing_time, 2),
                    'page_count': len(pdf_reader.pages)
                }
        except Exception as e:
            return None, {'error': str(e)}

    def scrape_contract_details(self, pdf_text):
        """Extract contract-specific details using regex patterns."""
        patterns = {
            'Contract ID': r"Contract ID:\s*(\S+)",
            'POP': r"POP:\s*([0-9\-]+ to [0-9\-]+)",
            'Contract Value': r"Contract Value:\s*\$?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)",
            'CLIN Description': r"CLIN Description:\s*(.*)",
            'PO Number': r"PO Number:\s*(\S+)",
            'Line Item': r"Line Item:\s*(.*)",
            'Line Item Value': r"Line Item Value:\s*\$?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)"
        }

        results = {}
        found_patterns = 0

        for key, pattern in patterns.items():
            match = re.search(pattern, pdf_text)
            results[key] = match.group(1) if match else None
            if match:
                found_patterns += 1

        success_rate = (found_patterns / len(patterns)) * 100
        return results, success_rate

    def process_directory(self):
        """Process all PDFs in directory with metrics tracking."""
        all_data = []
        all_metrics = []
        total_start_time = time.time()

        for filename in os.listdir(self.directory_path):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(self.directory_path, filename)
                print(f"Processing {filename}")

                # Extract text and performance metrics
                pdf_text, file_metrics = self.extract_text_from_pdf(pdf_path)
                if pdf_text:
                    # Extract contract details and success rate
                    contract_details, success_rate = self.scrape_contract_details(pdf_text)
                    
                    # Add filename and metrics
                    contract_details['Filename'] = filename
                    file_metrics['success_rate'] = success_rate
                    
                    all_data.append(contract_details)
                    all_metrics.append(file_metrics)
                    
                    self.metrics['processed_files'] += 1
                    if success_rate > 50:  # Consider extraction successful if over 50% fields found
                        self.metrics['successful_extractions'] += 1

        # Calculate overall metrics
        total_time = time.time() - total_start_time
        self.metrics.update({
            'total_processing_time': round(total_time, 2),
            'average_time_per_file': round(total_time / len(all_data), 2) if all_data else 0,
            'success_rate': round((self.metrics['successful_extractions'] / self.metrics['processed_files'] * 100), 2) if self.metrics['processed_files'] > 0 else 0
        })

        return pd.DataFrame(all_data), pd.DataFrame(all_metrics), self.metrics

    def export_results(self, df_data, df_metrics, output_dir):
        """Export results to Excel with separate sheets for data and metrics."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = os.path.join(output_dir, f'contract_analysis_{timestamp}.xlsx')

        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_data.to_excel(writer, sheet_name='Contract Details', index=False)
            df_metrics.to_excel(writer, sheet_name='File Metrics', index=False)
            pd.DataFrame([self.metrics]).to_excel(writer, sheet_name='Overall Metrics', index=False)

        return excel_path

def main():
    # Update this path to your PDF directory
    parser = ContractParser('/Users/tgalarneau2024/Github stuff/PDF-Contract-Parser/PDF-Storage')
    
    try:
        # Process PDFs and get results
        contract_data, file_metrics, overall_metrics = parser.process_directory()
        
        # Export results
        output_file = parser.export_results(contract_data, file_metrics, parser.directory_path)
        
        # Print summary
        print("\nPROCESSING SUMMARY")
        print("=" * 50)
        print(f"Total Files Processed: {overall_metrics['processed_files']}")
        print(f"Successful Extractions: {overall_metrics['successful_extractions']}")
        print(f"Success Rate: {overall_metrics['success_rate']}%")
        print(f"Total Processing Time: {overall_metrics['total_processing_time']} seconds")
        print(f"\nResults exported to: {output_file}")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
