import pandas as pd

# Load the pickle file
df_sma = pd.read_pickle("sma.pkl")
df_pvpm = pd.read_pickle("pvpm.pkl")

# Show first few rows
print("SMA Sheet:")
print(df_sma.head())

print("\nPVPM Sheet:")
print(df_pvpm.head())
