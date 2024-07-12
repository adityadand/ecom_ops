import os
import re
from PyPDF2 import PdfReader

# Function to extract amounts from a string
def extract_amount(text):
    # Regular expression to find amounts in the format ₹ 24,777.00
    amount_pattern = r'₹\s*([\d,]+\.\d{2})'
    amounts = re.findall(amount_pattern, text)
    return [float(amount.replace(',', '')) for amount in amounts]

# Get the current working directory
folder_path = os.getcwd()

# Initialize total
total_amount = 0.0

# Iterate through each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.pdf'):  # Assuming invoices are in PDF format
        file_path = os.path.join(folder_path, filename)
        with open(file_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            full_text = ""
            for page in pdf_reader.pages:
                full_text += page.extract_text()
            amounts = extract_amount(full_text)
            if amounts:
                print("File:", filename)
                for amount in amounts:
                    print("Amount:", "₹", amount)
                    total_amount += amount
                print()

# Print the total amount
print("Total amount:", "₹", round(total_amount, 2))
