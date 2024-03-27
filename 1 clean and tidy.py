import pandas as pd

file_path = 'Nov 2023.xlsx'
xls = pd.ExcelFile(file_path)

# Get all sheet names
sheet_names = xls.sheet_names
#print(sheet_names)

# Initialize an empty DataFrame for merged data
merged_data_with_dates = pd.DataFrame()

first_sheet = True
for sheet in sheet_names:
    if sheet.startswith("Day"):
        # Read the specific range (A3:O220) and the date from cell I1
        day_data = pd.read_excel(xls, sheet_name=sheet, usecols="A:O", skiprows=2, nrows=217)
        
        date = pd.read_excel(xls, sheet_name=sheet, header=None).iloc[0, 8]  # Read directly from cell I1
        # Add the date as a new column
        day_data['Date'] = date

        # Append to the merged DataFrame
        merged_data_with_dates = pd.concat([merged_data_with_dates, day_data])

# Reset index of the merged DataFrame
merged_data_with_dates.reset_index(drop=True, inplace=True)

# Removing empty rows based on the first column (assumed to be 'Time called in' in this case)
cleaned_data = merged_data_with_dates.dropna(subset=[merged_data_with_dates.columns[0]])

# Reset index after removing empty rows
cleaned_data.reset_index(drop=True, inplace=True)

cleaned_file_path = 'Merged_Nov_2023.xlsx'

with pd.ExcelWriter(cleaned_file_path) as writer:
    cleaned_data.to_excel(writer, index=False, sheet_name='Cleaned Data')
#cleaned_file_path

# Filling missing values with 0 in the 'Orders' and '# of passengers' columns
filled_data = pd.read_excel('Merged_Nov_2023.xlsx')
filled_data['Orders'] = filled_data['Orders'].fillna(0)
filled_data['# of passengers'] = filled_data['# of passengers'].fillna(0)
filled_data['#14 MICHAEL'] = filled_data['#14-A BLAIR'].fillna(0)
filled_data['#38A BLAIR OS'] = filled_data['#38 '].fillna(0)
filled_data['#46 JEREMY'] = filled_data['#46 JEREMY'].fillna(0)
filled_data['#47 PAOLO'] = filled_data['#47 PAOLO'].fillna(0)
filled_data['#51 '] = filled_data['#51 '].fillna(0)

filled_file_path = 'Filled_Merged_Nov_2023.xlsx'

with pd.ExcelWriter(filled_file_path) as writer:
    filled_data.to_excel(writer, index=False, sheet_name='Filled Data')
#filled_file_path


