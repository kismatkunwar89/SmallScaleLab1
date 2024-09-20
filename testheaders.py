import pandas as pd

# Step 1: Load the Excel file without headers
file_path = '311.xlsx'  # Update with the correct path to your Excel file
df = pd.read_excel(file_path, header=None)  # Load data without assuming the first row is the header

# Step 1.1: Define column headers
headers = ['Network Type', 'MCC', 'MNC', 'LAC', 'Cell ID', 'ARFCN', 'Longitude', 'Latitude', 'Altitude', 'Signal Strength', 'Network Type Indicator', 'Timestamp Start', 'Timestamp End', 'Additional Flag']
df.columns = headers  # Assign headers to the DataFrame

# Ensure the third column (MNC) is numeric, forcing errors to NaN
df['MNC'] = pd.to_numeric(df['MNC'], errors='coerce')

# Step 2: Filter rows based on the value '260' and '480' in the MNC (third column)
filtered_df_260 = df[df['MNC'] == 260]
filtered_df_480 = df[df['MNC'] == 480]

# Combine both filtered DataFrames
combined_filtered_df = pd.concat([filtered_df_260, filtered_df_480])

print("Combined filtered data based on 260 or 480 in the MNC column:")
print(combined_filtered_df)

# Optional: Save the combined filtered rows to a new Excel file if needed
combined_filtered_df.to_excel('combined_filtered_260_480.xlsx', index=False)

# Step 3: Function to load and deduplicate cell IDs from a text file
def load_and_deduplicate_cell_ids(file_path):
    with open(file_path, 'r') as file:
        cell_ids = [line.strip().split('-')[-1].strip() for line in file.readlines()]
    return list(set(cell_ids))

# Step 4: Load and deduplicate cell IDs for Mathew, Peter, and Sarah
cell_ids_mathew = load_and_deduplicate_cell_ids('mathew.txt')
cell_ids_peter = load_and_deduplicate_cell_ids('peter.txt')
cell_ids_sarah = load_and_deduplicate_cell_ids('sarah.txt')

# Step 5: Convert the relevant column (Cell ID) in the combined DataFrame to string and strip extra spaces
combined_filtered_df['Cell ID'] = combined_filtered_df['Cell ID'].astype(str).str.strip()

# Step 6: Filter the combined DataFrame based on matching cell IDs for each individual

# Filter for Mathew
mask_mathew = combined_filtered_df['Cell ID'].isin(cell_ids_mathew)
matching_rows_mathew = combined_filtered_df[mask_mathew]
matching_rows_mathew.to_excel('final_filtered_mathew.xlsx', index=False)

# Filter for Peter
mask_peter = combined_filtered_df['Cell ID'].isin(cell_ids_peter)
matching_rows_peter = combined_filtered_df[mask_peter]
matching_rows_peter.to_excel('final_filtered_peter.xlsx', index=False)

# Filter for Sarah
mask_sarah = combined_filtered_df['Cell ID'].isin(cell_ids_sarah)
matching_rows_sarah = combined_filtered_df[mask_sarah]
matching_rows_sarah.to_excel('final_filtered_sarah.xlsx', index=False)

print("Filtered and saved data for Mathew, Peter, and Sarah.")
