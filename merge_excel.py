import pandas as pd
import os
from tqdm import tqdm

# Define the input and output paths
input_folder = "sheet"  # Folder where your files are
output_file = "combined_kw.xlsx"  # Final merged file

# List all files inside the folder
all_files = os.listdir(input_folder)

# Create an empty list to collect all dataframes
all_dfs = []

# Setup tqdm fancy progress bar
for filename in tqdm(all_files, desc="📂 Merging Files", unit="file"):
    file_path = os.path.join(input_folder, filename)
    try:
        # Try reading as Excel first
        try:
            df = pd.read_excel(file_path)
        except Exception:
            # If Excel reading fails, try reading as CSV
            df = pd.read_csv(file_path, encoding='utf-8', engine='python')
        
        all_dfs.append(df)
    except Exception as e:
        print(f"❌ Error reading {filename}: {e}")

# Concatenate all dataframes together
if all_dfs:
    combined_df = pd.concat(all_dfs, ignore_index=True)
    combined_df.to_excel(output_file, index=False)
    print(f"\n✅ All files merged and saved to '{output_file}' successfully!")
else:
    print("\n⚠️ No readable files found or failed to read.")
