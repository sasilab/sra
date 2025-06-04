import pandas as pd
import numpy as np
from datetime import datetime
import pickle
import os
import re

# Folder containing the Excel files
folder_path = r'C:\Users\sas1924s\Downloads\Sasi_KickPV\PV Dataset\Measurement_Comparison_2025_02_18\Measurement_Comparison_2025_02_18\Measurement_Comparision_Data'  # Change this to your actual folder path

# Pattern to identify valid measurement files
file_pattern = re.compile(r'Measurement_Comparison_\d{4}_\d{2}_\d{2}.*\.xlsx')

# Create empty lists to store all data
all_sma_data = []
all_pvpm_data = []

# Iterate over all Excel files in the folder
for filename in os.listdir(folder_path):
    if not file_pattern.match(filename):
        print(f'Skipping unrelated file: {filename}')
        continue
    
    file_path = os.path.join(folder_path, filename)
    print(f'\nProcessing file: {filename}')
    
    # Load the Excel file
    try:
        excel = pd.ExcelFile(file_path)
    except Exception as e:
        print(f'  Failed to read {filename}: {e}')
        continue
    
    for sheet_name in excel.sheet_names:
        if sheet_name.lower() == 'overview':
            print(f'  Skipping Overview sheet')
            continue

        print(f'  Processing sheet: {sheet_name}')
        
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            timestamp = datetime.strptime(sheet_name, '%Y_%m_%d_%H_%M_%S')
        except Exception as e:
            print(f'    Skipping sheet due to error: {e}')
            continue

        mod_irr = df.iloc[1, 1] if not pd.isna(df.iloc[1, 1]) else None
        mod_temp = df.iloc[1, 7] if not pd.isna(df.iloc[1, 7]) else None

        data_rows = df.iloc[4:, :]

        sma_current_list = []
        sma_voltage_list = []
        pvpm_current_list = []
        pvpm_voltage_list = []

        for idx, row in data_rows.iterrows():
            if not pd.isna(row[0]) and not pd.isna(row[1]):
                sma_current_list.append(row[0])
                sma_voltage_list.append(row[1])
            if not pd.isna(row[4]) and not pd.isna(row[5]):
                pvpm_voltage_list.append(row[4])
                pvpm_current_list.append(row[5])

        sma_row = {
            'timestamp': timestamp,
            'mod_temp': mod_temp,
            'mod_irr': mod_irr,
            'Gut': 1,
            'I': sma_current_list,
            'V': sma_voltage_list
        }
        pvpm_row = {
            'timestamp': timestamp,
            'mod_temp': mod_temp,
            'mod_irr': mod_irr,
            'I': pvpm_current_list,
            'V': pvpm_voltage_list
        }

        all_sma_data.append(sma_row)
        all_pvpm_data.append(pvpm_row)

        print(f'    Extracted SMA: {len(sma_current_list)} I, {len(sma_voltage_list)} V')
        print(f'    Extracted PVPM: {len(pvpm_current_list)} I, {len(pvpm_voltage_list)} V')

# Create pandas DataFrames
sma_df = pd.DataFrame(all_sma_data)
pvpm_df = pd.DataFrame(all_pvpm_data)

# Summary
print('\n=== Summary ===')
print(f'Total SMA records: {len(sma_df)}')
print(f'Total PVPM records: {len(pvpm_df)}')

# Save to pickle
sma_pickle_path = 'sma_all_files.pkl'
pvpm_pickle_path = 'pvpm_all_files.pkl'

with open(sma_pickle_path, 'wb') as f:
    pickle.dump(sma_df, f)
with open(pvpm_pickle_path, 'wb') as f:
    pickle.dump(pvpm_df, f)

print(f'\nSaved SMA DataFrame to {sma_pickle_path}')
print(f'Saved PVPM DataFrame to {pvpm_pickle_path}')
