import pandas as pd
import numpy as np
from datetime import datetime
import pickle

# Load the Excel file
excel_file = 'Measurement_Comparison_2025_02_18.xlsx'

# Get all sheet names
excel = pd.ExcelFile(excel_file)
sheet_names = excel.sheet_names

# Create empty dataframes to store all data
all_sma_data = []
all_pvpm_data = []

# Loop through all sheets except 'Overview'
for sheet_name in sheet_names:
    if sheet_name == 'Overview':
        print(f'Skipping {sheet_name} sheet')
        continue
    
    print(f'Processing sheet: {sheet_name}')
    
    # Read the Excel sheet
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    
    # Extract timestamp from sheet name
    timestamp_str = sheet_name
    timestamp = datetime.strptime(timestamp_str, '%Y_%m_%d_%H_%M_%S')
    
    # Extract mod_irr from SMA (second row of column B)
    mod_irr = df.iloc[1, 1]  # 0-based indexing, so second row is index 1, column B is index 1
    
    # Extract mod_temp from PVPM (second row of column H)
    mod_temp = df.iloc[1, 7]  # Column H is index 7
    
    # Get the data rows (starting from row 5, which is index 4)
    data_rows = df.iloc[4:, :]
    
    # Create lists for SMA current and voltage values
    sma_current_list = []
    sma_voltage_list = []
    
    # Extract SMA current and voltage values
    for idx, row in data_rows.iterrows():
        # Check if we have valid data (not all NaN)
        if not pd.isna(row[0]) and not pd.isna(row[1]):
            # Get current (I) and voltage (V) from SMA columns
            current = row[0]  # MPP tracker B: I
            voltage = row[1]  # MPP tracker B: V
            
            sma_current_list.append(current)
            sma_voltage_list.append(voltage)
    
    # Create lists for PVPM current and voltage values
    pvpm_current_list = []
    pvpm_voltage_list = []
    
    # Extract PVPM current and voltage values
    for idx, row in data_rows.iterrows():
        # Check if we have valid data (not all NaN)
        if not pd.isna(row[4]) and not pd.isna(row[5]):
            # Get current (I) and voltage (V) from PVPM columns
            voltage = row[4]  # U in V
            current = row[5]  # I in A
            
            pvpm_current_list.append(current)
            pvpm_voltage_list.append(voltage)
    
    # Create row data for SMA
    sma_row = {
        'timestamp': timestamp,
        'mod_temp': mod_temp,
        'mod_irr': mod_irr,
        'Gut': 1,
        'I': sma_current_list,
        'V': sma_voltage_list
    }
    
    # Create row data for PVPM
    pvpm_row = {
        'timestamp': timestamp,
        'mod_temp': mod_temp,
        'mod_irr': mod_irr,
        'Gut': 1,
        'I': pvpm_current_list,
        'V': pvpm_voltage_list
    }
    
    # Append to the data lists
    all_sma_data.append(sma_row)
    all_pvpm_data.append(pvpm_row)
    
    print(f'  Extracted SMA data: {len(sma_current_list)} current values, {len(sma_voltage_list)} voltage values')
    print(f'  Extracted PVPM data: {len(pvpm_current_list)} current values, {len(pvpm_voltage_list)} voltage values')

# Create pandas DataFrames from all collected data
sma_df = pd.DataFrame(all_sma_data)
pvpm_df = pd.DataFrame(all_pvpm_data)

# Print summary of the dataframes
print('\nSMA DataFrame summary:')
print(f'Number of rows (sheets processed): {len(sma_df)}')
print(f'Columns: {sma_df.columns.tolist()}')

print('\nPVPM DataFrame summary:')
print(f'Number of rows (sheets processed): {len(pvpm_df)}')
print(f'Columns: {pvpm_df.columns.tolist()}')

# Save dataframes to pickle files
sma_pickle_path = 'sma_all_sheets.pkl'
pvpm_pickle_path = 'pvpm_all_sheets.pkl'

# Save to pickle files
with open(sma_pickle_path, 'wb') as f:
    pickle.dump(sma_df, f)

with open(pvpm_pickle_path, 'wb') as f:
    pickle.dump(pvpm_df, f)

print(f'\nSaved SMA DataFrame to {sma_pickle_path}')
print(f'Saved PVPM DataFrame to {pvpm_pickle_path}')
