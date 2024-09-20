import pandas as pd

# Function to load cell IDs from the text file and remove duplicates
def load_and_deduplicate_cell_ids(file_path):
    with open(file_path, 'r') as file:
        # Strip spaces and extract the third part (ensure everything is a string)
        cell_ids = [line.strip().split('-')[-1].strip() for line in file.readlines()]
    # Remove duplicates by converting the list to a set and then back to a list
    return list(set(cell_ids))

# Step 1: Load the Excel file (filtered on the value '480' in the third column)
excel_file_path = '311.xlsx'  # Update with the correct path to your Excel file
df = pd.read_excel(excel_file_path, header=None)  # Load data without assuming the first row is the header

# Step 2: Display column names to ensure we use the correct column for filtering
print("Column names:", df.columns)

# Step 3: Filter rows based on the value '480' in the third column (index 2)
filtered_df = df[df.iloc[:, 2] == 480]  # iloc[:, 2] refers to the third column (0-based index)

print("Filtered data based on 480 in the third column:")
print(filtered_df)

# Save the filtered rows to a new Excel file if needed
filtered_df.to_excel('filtered_480.xlsx', index=False)

# Step 4: Load and deduplicate cell IDs from mathew.txt
cell_ids = load_and_deduplicate_cell_ids('mathew.txt')

print("Loaded and deduplicated cell IDs from mathew.txt:")
print(cell_ids[:10])  # Print a sample for inspection

# Step 5: Convert the relevant column in the filtered DataFrame to string and strip extra spaces
filtered_df.iloc[:, 4] = filtered_df.iloc[:, 4].astype(str).str.strip()

# Convert cell_ids to strings to ensure compatibility
cell_ids = [str(cell_id).strip() for cell_id in cell_ids]

# Step 6: Filter the filtered DataFrame based on matching cell IDs
mask = filtered_df.iloc[:, 4].isin(cell_ids)
matching_rows = filtered_df[mask]

print("Final filtered rows based on matching CELLID:")
print(matching_rows)

# Optional: Save the final filtered rows to a new Excel file if needed
matching_rows.to_excel('final_filtered_480_cell_ids.xlsx', index=False)
