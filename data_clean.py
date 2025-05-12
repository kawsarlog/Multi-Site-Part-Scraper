import os
import pandas as pd

excel_file = "Parts for Xref 4.22.25.xlsx"
folder_path = "output"
output_file = "output/final_output.csv"
excel_file_output = "output/final_output_file.xlsx"

# Step 1: Merge all CSV files
csv_files = [f for f in os.listdir(folder_path) if f.endswith(".csv")]

merged_df = None
for file in csv_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_csv(file_path)
    df.fillna("n/a", inplace=True)
    if merged_df is None:
        merged_df = df
    else:
        merged_df = pd.merge(merged_df, df, on=["Part Number Reformat", "DESCRIPTION", "Invoice Reference"], how='outer')
merged_df.fillna("n/a", inplace=True)
merged_df.to_csv(output_file, index=False)
print(f"Merged CSV created: {output_file}")

# Step 2: Add missing parts from Excel
merged_df = pd.read_csv(output_file, dtype=str).fillna("N/A")
input_df = pd.read_excel(excel_file, dtype=str).fillna("N/A")
input_df = input_df.rename(columns={
    "Part Number": "Part Number Reformat",
    "DISCRIPTION": "DESCRIPTION"
})
missing_parts_df = input_df[~input_df["Part Number Reformat"].isin(merged_df["Part Number Reformat"])].copy()
for col in merged_df.columns:
    if col not in ["Part Number Reformat", "DESCRIPTION"]:
        missing_parts_df[col] = "N/A"
final_df = pd.concat([merged_df, missing_parts_df], ignore_index=True)
final_df.to_csv(f"{folder_path}/merged_output_with_missing.csv", index=False)

# Step 3: Merge with Excel again and clean columns
merged_df = pd.read_csv(f"{folder_path}/merged_output_with_missing.csv", dtype=str).fillna("N/A")
excel_df = pd.read_excel(excel_file, dtype=str).fillna("N/A")
excel_df = excel_df.rename(columns={
    "Part Number": "Part Number Reformat",
    "DISCRIPTION": "DESCRIPTION"
})
final_df = pd.merge(merged_df, excel_df, on=["Part Number Reformat", "DESCRIPTION"], how="left")

# Remove unwanted _x columns
final_df.drop(columns=[
    "ID_x", "PN + Desc_x", "Part Num from Desc_x", "Part Num wo Code_x"
], inplace=True, errors='ignore')

# Rename _y columns by removing the suffix
final_df.columns = [
    col[:-2] if col.endswith('_y') else col for col in final_df.columns
]

# Final save to Excel
final_df.to_excel(excel_file_output, index=False)
print("✅ Done!")
print(f"✅ Cleaned output file name is {excel_file_output}")
