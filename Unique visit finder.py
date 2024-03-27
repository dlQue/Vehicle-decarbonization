
import pandas as pd

# Load the data from the specified sheet "SepNov"
file_path = '0 GeoTab full year.xlsx'
sheet_name = 'Sheet1 (2)'

# Load the data
data_sep_nov = pd.read_excel(file_path, sheet_name=sheet_name)

# Filter data for specified vehicles
vehicles = ['46', '47', '14A', '51', '38A']  # Update this list if you have different vehicle identifiers
filtered_data = data_sep_nov[data_sep_nov['Vehicle'].isin(vehicles)]

# Extract date from "Trip Started" to identify visits per day
filtered_data['Date'] = filtered_data['Trip Started'].dt.date

# Count unique and total visits per day per vehicle
# Assuming a "visit" is defined by a unique start location
visits_per_vehicle_per_day = filtered_data.groupby(['Date', 'Vehicle'])['End Location Modified'].agg(['nunique', 'count']).reset_index()
visits_per_vehicle_per_day.rename(columns={'nunique': 'Unique Visits', 'count': 'Total Visits'}, inplace=True)

# Specify your desired output file path
output_file_path = 'your_output_file_path_here11.xlsx'  # Update this path

# Save the processed data to an Excel file
visits_per_vehicle_per_day.to_excel(output_file_path, index=False, sheet_name='VisitsPerVehiclePerDay')

print(f"Output saved to {output_file_path}")
