import os
import pandas as pd

# Specify the folder containing the Excel files
folder_path = ''
output_file = 'formatted_combined_output.xlsx'

# Initialize a list to store dataframes
combined_dataframes = []

# Loop through each file in the folder
for filename in os.listdir(folder_path):
  if filename.endswith('.xlsx'):
    file_path = os.path.join(folder_path, filename)
    
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Extract the required columns
    if 'ASIN' in df.columns and 'Quantity Outstanding' in df.columns:
      df_extracted = df[['ASIN', 'Quantity Outstanding']].copy()
      
      # Check existing number of columns
      num_columns = len(df_extracted.columns)

      # Create MultiIndex header based on column count
      if num_columns == 2:  # If only ASIN and Quantity columns
        headers = pd.MultiIndex.from_tuples([
          (filename, 'ASIN'),
          (filename, 'Quantity')
        ])
      else:  # If there are more columns (add empty string level)
        headers = pd.MultiIndex.from_tuples([
          (filename, ''),
          (filename, 'ASIN'),
          (filename, 'Quantity')
        ])
      
      df_extracted.columns = headers
      
      # Append the dataframe with MultiIndex header to list
      combined_dataframes.append(df_extracted)

# Concatenate all dataframes horizontally
final_df = pd.concat(combined_dataframes, axis=1)

# Save the combined dataframe to a new Excel file with index
final_df.to_excel(output_file, index=True)
