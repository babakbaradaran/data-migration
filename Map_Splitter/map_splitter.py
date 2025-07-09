import pandas as pd
import os

# Get the current directory where the script is located
current_dir = os.getcwd()

# Define required columns
required_columns = ['old_job', 'PHASE', 'Old Phase Code Description', 'New Phase', 'New Phase Description']

# Loop through all Excel files in the directory usind for loop and os
for file_name in os.listdir(current_dir):
    if file_name.endswith(".xlsx") and not file_name.startswith("m") and "~$" not in file_name:
        try:
            file_path = os.path.join(current_dir, file_name)
            print(f"\nProcessing file: {file_name}")

            # Read the 'Sheet1' from the file
            df = pd.read_excel(file_path, sheet_name='Sheet1')

            # Check if all required columns exist
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                print(f" Skipping {file_name}: missing columns: {missing_columns}")
                continue

            # Drop NaNs in 'old_job' and get unique values
            unique_old_jobs = df['old_job'].dropna().unique()

            # Loop through each unique old_job
            for old_job in unique_old_jobs:
                filtered_df = df[df['old_job'] == old_job][[
                    'PHASE', 'Old Phase Code Description', 'New Phase', 'New Phase Description'
                ]]

                output_filename = f"m{old_job}.xlsx"
                output_path = os.path.join(current_dir, output_filename)

                filtered_df.to_excel(output_path, index=False)
                print(f"Created: {output_filename} with {len(filtered_df)} rows")

        except Exception as e:
            print(f"Error processing {file_name}: {e}")

print("\n All eligible files have been processed.")
