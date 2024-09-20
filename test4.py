import pandas as pd
file_path = '311.xlsx'  # path of excel file 
df = pd.read_excel(file_path, header=None) 

df.iloc[:, 2] = pd.to_numeric(df.iloc[:, 2], errors='coerce') # Ensure the third column is numeric and error handling
filtered_df_260 = df[df.iloc[:, 2] == 260]  # Step 2: Filter rows based on the value '260' and '480' in the third column
filtered_df_480 = df[df.iloc[:, 2] == 480]

combined_filtered_df = pd.concat([filtered_df_260, filtered_df_480]) # Combine both filtered DataFrames
print("Combined filtered data based on 260 or 480 in the third column:")
print(combined_filtered_df)

combined_filtered_df.to_excel('combined_filtered_260_480.xlsx', index=False) # Save the combined filtered rows to a new Excel file if needed

def load_and_deduplicate_cell_ids(file_path):  # Function to load and deduplicate cell IDs from a text file
    with open(file_path, 'r') as file:
        cell_ids = [line.strip().split('-')[-1].strip() for line in file.readlines()]
    return list(set(cell_ids))
cell_ids_mathew = load_and_deduplicate_cell_ids('mathew.txt')   #Load and deduplicate cell IDs for Mathew, Peter, and Sarah
cell_ids_peter = load_and_deduplicate_cell_ids('peter.txt')
cell_ids_sarah = load_and_deduplicate_cell_ids('sarah.txt')

combined_filtered_df.iloc[:, 4] = combined_filtered_df.iloc[:, 4].astype(str).str.strip() # Step 5: Convert the relevant column in the combined DataFrame to string 
mask_mathew = combined_filtered_df.iloc[:, 4].isin(cell_ids_mathew)  # Filter for Mathew
matching_rows_mathew = combined_filtered_df[mask_mathew]
matching_rows_mathew.to_excel('final_filtered_mathew.xlsx', index=False)

mask_peter = combined_filtered_df.iloc[:, 4].isin(cell_ids_peter)  # Filter for Peter
matching_rows_peter = combined_filtered_df[mask_peter]
matching_rows_peter.to_excel('final_filtered_peter.xlsx', index=False)

mask_sarah = combined_filtered_df.iloc[:, 4].isin(cell_ids_sarah) # Filter for Sarah
matching_rows_sarah = combined_filtered_df[mask_sarah]
matching_rows_sarah.to_excel('final_filtered_sarah.xlsx', index=False)

print("Filtered and saved data for Mathew, Peter, and Sarah.")
