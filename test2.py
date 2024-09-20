import pandas as pd

# Step 1: Load the Excel file without headers
file_path = '480.xlsx'  # Update with the correct path to your Excel file
df = pd.read_excel(file_path, header=None)  # Load data without assuming the first row is the header

# Print the first few rows to verify the structure
print("Initial data from Excel file:")
print(df.head())

# Print column names to verify which column has the CELLID
print("Column names in the DataFrame (Index):")
print(df.columns)

# Print sample values from column 4 to verify the content
print("Sample values from column 4:")
print(df.iloc[:, 4].head())

# Step 2: Load the cell IDs from the text file (mathew.txt) and remove duplicates
with open('mathew.txt', 'r') as file:
    # Strip spaces and extract the third part (ensure everything is a string)
    cell_ids = [line.strip().split('-')[-1].strip() for line in file.readlines()]

# Remove duplicates by converting the list to a set and then back to a list
cell_ids = list(set(cell_ids))

print("Loaded and deduplicated cell IDs from mathew.txt:")
print(cell_ids[:10])  # Print a sample for inspection

# Step 3: Convert the 4th column in the DataFrame to string and strip extra spaces
df.iloc[:, 4] = df.iloc[:, 4].astype(str).str.strip()

# Convert cell_ids to strings to ensure compatibility
cell_ids = [str(cell_id).strip() for cell_id in cell_ids]

# Step 4: Filter the DataFrame based on matching cell IDs
# Create a mask where cell IDs in the 4th column match any ID from cell_ids
mask = df.iloc[:, 4].isin(cell_ids)

# Apply the mask to get matching rows
matching_rows = df[mask]

# Print the final matching rows
print("Final filtered rows based on matching CELLID:")
print(matching_rows)

# Optional: Save the filtered rows to a new Excel file if needed
matching_rows.to_excel('filtered_480_cell_ids.xlsx', index=False)
