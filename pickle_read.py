import pandas as pd
import numpy as np
from datetime import datetime
import pickle

# Load the Excel file
excel_file = r'F:\Krishnas_Space\code\Others\Measurement_Comparison_2025_02_18\Measurement_Comparison_2025_02_18.xlsx'
sheet_name = '2025_02_18_10_41_00'

# Read the Excel sheet
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

# Extract timestamp from sheet name
timestamp_str = sheet_name
timestamp = datetime.strptime(timestamp_str, '%Y_%m_%d_%H_%M_%S')

# Extract mod_irr from SMA (second row of column B)
mod_irr = df.iloc[1, 1]  # 0-based indexing, so second row is index 1, column B is index 1
print(f'Extracted mod_irr: {mod_irr}')

# Extract mod_temp from PVPM (second row of column H)
mod_temp = df.iloc[1, 7]  # Column H is index 7
print(f'Extracted mod_temp: {mod_temp}')

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

# Create single-row SMA dataframe with lists
sma_df = pd.DataFrame({
    'timestamp': [timestamp],
    'mod_temp': [mod_temp],
    'mod_irr': [mod_irr],
    'Gut': [1],
    'I': [sma_current_list],
    'V': [sma_voltage_list]
})

# Create single-row PVPM dataframe with lists
pvpm_df = pd.DataFrame({
    'timestamp': [timestamp],
    'mod_temp': [mod_temp],
    'mod_irr': [mod_irr],
    'Gut': [1],
    'I': [pvpm_current_list],
    'V': [pvpm_voltage_list]
})

# Print the dataframes to verify
print('\nSMA DataFrame (single row with lists):')
print(sma_df)
print('\nSMA current list length:', len(sma_df.loc[0, 'I']))
print('SMA voltage list length:', len(sma_df.loc[0, 'V']))

print('\nPVPM DataFrame (single row with lists):')
print(pvpm_df)
print('\nPVPM current list length:', len(pvpm_df.loc[0, 'I']))
print('PVPM voltage list length:', len(pvpm_df.loc[0, 'V']))

# Save dataframes to pickle files
sma_pickle_path = 'sma.pkl'
pvpm_pickle_path = 'pvpm.pkl'

# Save to pickle files
with open(sma_pickle_path, 'wb') as f:
    pickle.dump(sma_df, f)

with open(pvpm_pickle_path, 'wb') as f:
    pickle.dump(pvpm_df, f)

print(f'\nSaved SMA DataFrame to {sma_pickle_path}')
print(f'Saved PVPM DataFrame to {pvpm_pickle_path}')