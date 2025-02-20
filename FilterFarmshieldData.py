import pandas as pd

# File paths
input_file = "FarmshieldData(29thJan-7thFeb).xlsx"
output_file = "Filtered_Farmers.xlsx"

# Define the sheet name
sheet_name = 'FarmshieldData(29th-7th)'

# Define the columns of interest
columns_of_interest = [
    'Record_ID_system_date', 'Record_ID_system_time', 'AIR_TEMPERATURE', 
    'AIR_HUMIDITY', 'LIGHT_INTENSITY', 'CARBON_DIOXIDE', 'SOIL_MOISTURE', 
    'SOIL_NITROGEN', 'SOIL_PHOSPHOROUS', 'SOIL_POTASSIUM', 'SOIL_TEMPERATURE'
]

# Define the date range for filtering
start_date = '2025-01-30 12:15:00'
end_date = '2025-02-08 12:15:01'

try:
    # Load the Excel file
    df = pd.read_excel(input_file, sheet_name=sheet_name, dtype=str)
    
    # Ensure the necessary columns exist
    required_columns = {'Record_ID_system_date', 'Record_ID_system_time', 'Farmer'}
    if not required_columns.issubset(df.columns):
        missing_cols = required_columns - set(df.columns)
        raise KeyError(f"Missing required columns: {missing_cols}")
    
    # Convert date and time columns to datetime
    df['Record_ID_system_date'] = pd.to_datetime(df['Record_ID_system_date'], errors='coerce')
    df['datetime'] = pd.to_datetime(df['Record_ID_system_date'].astype(str) + ' ' + df['Record_ID_system_time'], errors='coerce')
    
    # Convert filtering dates to datetime
    start_datetime = pd.to_datetime(start_date)
    end_datetime = pd.to_datetime(end_date)
    
    # Filter the data within the datetime range
    df = df[(df['datetime'] >= start_datetime) & (df['datetime'] <= end_datetime)]
    
    # Get unique Farmer types
    farmer_types = df['Farmer'].unique()
    
    # Create a dictionary to store dataframes for each Farmer type
    filtered_sheets = {}
    
    for farmer in farmer_types:
        farmer_df = df[df['Farmer'] == farmer]
        available_columns = [col for col in columns_of_interest if col in farmer_df.columns]
        filtered_sheets[farmer] = farmer_df[available_columns]
    
    # Save to an Excel file with multiple sheets
    with pd.ExcelWriter(output_file) as writer:
        for farmer, data in filtered_sheets.items():
            data.to_excel(writer, sheet_name=farmer, index=False)
    
    print(f"Filtered data successfully saved to: {output_file}")

except KeyError as e:
    print(f"Column not found in the Excel file: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
