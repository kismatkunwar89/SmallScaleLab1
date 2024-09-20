import pandas as pd

# Step 1: Load the Excel file without headers
file_path = '311.xlsx'  # Update with the correct path to your Excel file
df = pd.read_excel(file_path, header=None)  # Load data without assuming the first row is the header

# Print column names to ensure we are using the correct column for filtering
print("Column names:", df.columns)

# Step 2: Verify data in the third column
print("Sample values from the third column:")
print(df.iloc[:, 2].head())

# Ensure the column is numeric, forcing errors to NaN
df.iloc[:, 2] = pd.to_numeric(df.iloc[:, 2], errors='coerce')

# Print unique values in the third column to debug
print("Unique values in the third column after conversion:")
print(df.iloc[:, 2].unique())

# Step 3: Filter rows based on the value '260' in the third column
filtered_df_260 = df[df.iloc[:, 2] == 260]

print("Filtered data based on 260 in the third column:")
print(filtered_df_260)

# Step 4: Filter rows based on the value '480' in the third column
filtered_df_480 = df[df.iloc[:, 2] == 480]

print("Filtered data based on 480 in the third column:")
print(filtered_df_480)

# Step 5: Combine the results of both filters
combined_filtered_df = pd.concat([filtered_df_260, filtered_df_480])

print("Combined filtered data based on 260 or 480 in the third column:")
print(combined_filtered_df)

# Optional: Save the combined filtered rows to a new Excel file if needed
combined_filtered_df.to_excel('combined_filtered_260_480.xlsx', index=False)

# Step 6: Load and deduplicate cell IDs from mathew.txt
def load_and_deduplicate_cell_ids(file_path):
    with open(file_path, 'r') as file:
        cell_ids = [line.strip().split('-')[-1].strip() for line in file.readlines()]
    return list(set(cell_ids))

cell_ids = load_and_deduplicate_cell_ids('mathew.txt')

cell_ids = load_and_deduplicate_cell_ids('peter.txt')

cell_ids = load_and_deduplicate_cell_ids('sarah.txt')

print("Loaded and deduplicated cell IDs from mathew.txt:")
print(cell_ids[:10])  # Print a sample for inspection

# Step 7: Convert the relevant column in the combined DataFrame to string and strip extra spaces
combined_filtered_df.iloc[:, 4] = combined_filtered_df.iloc[:, 4].astype(str).str.strip()

# Convert cell_ids to strings to ensure compatibility
cell_ids = [str(cell_id).strip() for cell_id in cell_ids]

# Step 8: Filter the combined DataFrame based on matching cell IDs
mask = combined_filtered_df.iloc[:, 4].isin(cell_ids)
matching_rows = combined_filtered_df[mask]

print("Final filtered rows based on matching CELLID:")
print(matching_rows)

# Optional: Save the final filtered rows to a new Excel file if needed
matching_rows.to_excel('final_filtered_combined_cell_ids.xlsx', index=False)
