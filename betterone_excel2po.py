import pandas as pd

# Load the formatted Excel file
file_path = 'datafile.xlsx'

# Read the Excel file with headers in the 4th row
df = pd.read_excel(file_path, header=0)

# Clean column names by stripping extra spaces
df.columns = df.columns.str.strip()

# Print column names to check if headers are loaded correctly
print("Columns in the DataFrame:", df.columns.tolist())

# Ensure 'ASIN' column exists
asin_column_name = 'ASIN'
if asin_column_name not in df.columns:
    raise KeyError(f"Column '{asin_column_name}' not found")

# Extract PO columns dynamically
po_columns_prefix = 'PO: '
po_columns = [col for col in df.columns if col.startswith(po_columns_prefix)]

# Function to extract PO identifier from column names
def extract_po_info(po_column_name):
    # Extract PO identifier from column name
    po_identifier = po_column_name.split(': ')[1]
    return po_identifier

# Create a dictionary to store results PO-wise
results = {}

# Iterate over each row in the DataFrame
for _, row in df.iterrows():
    asin = row[asin_column_name]  # Access ASIN column
    
    # Process each PO column
    for po_col in po_columns:
        po_identifier = extract_po_info(po_col)
        quantity = row[po_col]  # Access quantity from PO column
        
        if pd.notna(quantity):  # Only include rows where quantity is not NaN
            # Remove decimal places if quantity is a float
            formatted_quantity = f"{int(quantity)}" if quantity.is_integer() else f"{quantity:.0f}"
            
            # Initialize PO entry if it doesn't exist
            if po_identifier not in results:
                results[po_identifier] = []
            
            # Format the result and append to the PO entry
            result = f"{asin} - {formatted_quantity} UNITS"
            results[po_identifier].append(result)

# Output results to a text file
with open('output.txt', 'w') as f:
    for po_id, entries in results.items():
        f.write(f"PO: {po_id} (ETRADE)\n")
        for entry in entries:
            f.write(f"{entry}\n")
        f.write("\n")

print("Processing complete. Results saved to output.txt")
