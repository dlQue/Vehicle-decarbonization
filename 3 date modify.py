import pandas as pd
from dateutil.parser import parse
import numpy as np

# convert various date formats to 'yyyy-mm-dd'
def convert_date(date_entry):
    try:
        # Convert to string if not already a string
        if not isinstance(date_entry, str):
            date_entry = str(date_entry)
        # Use dateutil's parse to handle various formats
        date_obj = parse(date_entry)
        # Convert to 'yyyy-mm-dd' format
        return date_obj.strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        # Return NaN for unparseable dates
        return np.nan

def main():
    file_path = 'Merged_3mon.xlsx'
    output_file_path = 'modified_dates.xlsx' 

    data = pd.read_excel(file_path)

    data['Date'] = data['Date'].apply(convert_date)

    data.to_excel(output_file_path, index=False)
    print(f"File saved as '{output_file_path}'")

if __name__ == "__main__":
    main()
