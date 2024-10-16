import PyPDF2
import re
import pandas as pd

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        all_text = ''
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            all_text += page.extract_text()
        return all_text

# Function to scrape relevant data from the extracted text
def scrape_contract_details(pdf_text):
    # Use regex to find specific details
    contract_id_pattern = r"Contract ID:\s*(\S+)"
    pop_pattern = r"POP:\s*([0-9\-]+ to [0-9\-]+)"
    contract_value_pattern = r"Contract Value:\s*\$?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)"
    clin_description_pattern = r"CLIN Description:\s*(.*)"
    po_number_pattern = r"PO Number:\s*(\S+)"
    line_item_pattern = r"Line Item:\s*(.*)"
    line_item_value_pattern = r"Line Item Value:\s*\$?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)"

    contract_id = re.search(contract_id_pattern, pdf_text)
    pop = re.search(pop_pattern, pdf_text)
    contract_value = re.search(contract_value_pattern, pdf_text)
    clin_description = re.search(clin_description_pattern, pdf_text)
    po_number = re.search(po_number_pattern, pdf_text)
    line_item = re.search(line_item_pattern, pdf_text)
    line_item_value = re.search(line_item_value_pattern, pdf_text)

    return {
        "Contract ID": contract_id.group(1) if contract_id else None,
        "POP": pop.group(1) if pop else None,
        "Contract Value": contract_value.group(1) if contract_value else None,
        "CLIN Description": clin_description.group(1) if clin_description else None,
        "PO Number": po_number.group(1) if po_number else None,
        "Line Item": line_item.group(1) if line_item else None,
        "Line Item Value": line_item_value.group(1) if line_item_value else None
    }

# Function to process the PDF and return the data
def process_pdf_to_dataframe(pdf_path):
    pdf_text = extract_text_from_pdf(pdf_path)
    contract_details = scrape_contract_details(pdf_text)

    # Convert the result to a pandas DataFrame
    df = pd.DataFrame([contract_details])
    return df

# Replace this with the actual path to your PDF file
pdf_file_path = '/Users/tgalarneau2024/Documents/16-306-Cost-plus-fixed-fee-contracts-.pdf'
df = process_pdf_to_dataframe(pdf_file_path)

# Output the extracted data
print(df)

# Save the extracted data to an Excel file
df.to_excel('extracted_contract_data.xlsx', index=False)

