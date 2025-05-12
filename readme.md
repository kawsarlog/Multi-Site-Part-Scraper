# Multi-Site Part Scraper & Data Merger Tool

This project automates part number lookups from four different websites using a shared Excel input, and merges all data into a clean, unified Excel output file.

## Folder Structure Overview 

```
project-folder/
│
├── Aurora.py
├── FleetPride.py
├── Meritor.py
├── Napa.py
├── data_clean.py
├── Parts for Xref 4.22.25.xlsx ← Main Excel Input File
├── output/
│   ├── Aurora_output.csv
│   ├── fleetpride_output.csv
│   ├── meritorpartsxpress_output.csv
│   ├── napa_output.csv
│   ├── final_output.csv
│   ├── merged_output_with_missing.csv
│   └── final_output_file.xlsx ← Final Cleaned Excel File
```

## Requirements Installation

Install dependencies before running the scripts: 

```bash
pip install -r requirements.txt
```

If installation fails, you may need to install specific versions:

```bash
pip install pandas==2.2.3 requests==2.32.3 selenium==4.27.1 seleniumbase==4.33.3 openpyxl==3.1.5
```

## Input File Format

The input file must contain exactly 6 columns with the following headers:
1. ID
2. Part Number Reformat
3. DESCRIPTION
4. PN + Desc
5. Part Num from Desc
6. Part Num wo Code

⚠️ Column order is critical as the program uses index-based access to these columns.

DO NOT change the structure or column names except the file name.

Allowed: Renaming file to something like My_New_Input_File.xlsx is acceptable.
Not Allowed: Changing the column names or number of columns.

## Usage Instructions

### Step 1: Place Input File
Put the Excel input file (e.g., Parts for Xref 4.22.25.xlsx) in the same folder as all Python scripts. You can rename the file, but if renamed, update all script lines where the file is mentioned:
```python
excel_file = "YourNewFileName.xlsx"
```

### Step 2: Run the 4 Scraper Scripts
Execute each of these scripts individually to scrape data:
```bash
python Aurora.py
python FleetPride.py
python Meritor.py
python Napa.py
```

Each script will:
- 📥 Load the input Excel file
- 🔍 Scrape data from the respective website
- 💾 Save the results as a .csv file in the output/ folder

### Step 3: Run the Final Merger Script
After all scrapers are complete, run the final script to clean and merge everything:
```bash
python data_clean.py
```

This script will:
- 🔄 Merge all 4 CSV files from the output/ folder
- 🧩 Fill missing values with "N/A"
- 🔍 Match original input parts with scraped data
- ➕ Identify and append missing parts
- 💾 Save:
  - final_output.csv (Raw merged CSV)
  - merged_output_with_missing.csv (Includes unmatched rows)
  - final_output_file.xlsx (Final Cleaned Excel Output)

You will see this message after successful completion:
```
✅ Done! ✅ Output file name is output/final_output_file.xlsx
```

## Important Notes

- ✅ You can change the Excel file name, but must update the `excel_file = "..."` line in data_clean.py and all scraper scripts.
- ❌ Do not modify the column order or headers inside the Excel file.
- 📁 All scraped .csv outputs must be saved in the output/ directory for the merger to work correctly.
- 📊 The final output will always be saved as output/final_output_file.xlsx.

### Terms for Aurora

⚠️ By using this tool with Aurora, you acknowledge and agree that cookies must be enabled for proper functionality. The Aurora scraper component requires cookies for session management, authentication, and data retrieval. Using this tool with Aurora indicates your acceptance of their cookie policy and terms of service.

## Customization (If Needed)

If you're using a different input file name (e.g., ClientData.xlsx), update this line in all scripts:
```python
excel_file = "ClientData.xlsx"
```