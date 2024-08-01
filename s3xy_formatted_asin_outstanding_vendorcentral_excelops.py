import os
import pandas as pd

# Get the folder containing the script file
folder_path = os.path.dirname(os.path.realpath(__file__))
output_file = os.path.join(folder_path, 'formatted_combined_output.xlsx')

# Initialize a variable to store the combined dataframe
combined_df = None

# Loop through each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx') and filename != 'formatted_combined_output.xlsx':  # Exclude the output file
        file_path = os.path.join(folder_path, filename)
        
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Extract the required columns
        if 'ASIN' in df.columns and 'Quantity Outstanding' in df.columns:
            df_extracted = df[['ASIN', 'Quantity Outstanding']].copy()
            df_extracted.set_index('ASIN', inplace=True)
            df_extracted.columns = [filename]  # Rename the column to the filename
            
            # Merge the data into the combined dataframe
            if combined_df is None:
                combined_df = df_extracted
            else:
                combined_df = combined_df.merge(df_extracted, left_index=True, right_index=True, how='outer')

# Convert the combined dataframe to a final dataframe
if combined_df is not None:
    final_df = combined_df.reset_index()

    # Rename the columns to include 'Quantity' with filename
    final_df.columns = ['ASIN'] + [f'{col}_Quantity' for col in final_df.columns[1:]]

    # Save the combined dataframe to a new Excel file
    final_df.to_excel(output_file, index=False)
else:
    print("No files with the required columns found.")
