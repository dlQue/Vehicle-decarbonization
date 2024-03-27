import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

file_path_oct = 'Merged_3mon.xlsx'
vehicle_data_oct = pd.read_excel(file_path_oct)

# Function to convert 'HH:MM:SS' to minutes
def convert_to_minutes(time_str):
    if pd.isnull(time_str) or time_str == '':
        return None
    try:
        h, m, s = map(int, time_str.split(':'))
        return h * 60 + m + s / 60
    except ValueError:
        return None

# Apply the conversion to the 'Time for Call' column
vehicle_data_oct['Trip Duration'] = vehicle_data_oct['Time for Call'].apply(convert_to_minutes)

# Round the trip duration to the nearest 5 minutes
vehicle_data_oct['Rounded Trip Duration'] = 5 * np.round(vehicle_data_oct['Trip Duration'] / 5)

# Count the frequency of each rounded trip duration
duration_counts = vehicle_data_oct['Rounded Trip Duration'].value_counts()


# Visualization
plt.figure(figsize=(20, 12))  # Increase figure size to accommodate tables

# Most Common Destinations
valid_destinations = vehicle_data_oct[~vehicle_data_oct['To'].isin(['nan', 'Mileage', 'Start of Day', 'End of Day', 'Total for Day', 'Truck Washed', '.', 'DOCK'])]
destination_counts = valid_destinations['To'].value_counts().head(10)
plt.subplot(2, 3, 1)  # Change to 2, 3, 1 for the first subplot in the first row
sns.barplot(x=destination_counts.values, y=destination_counts.index)
plt.title('Top 10 Most Common Destinations')
plt.xlabel('Count')
plt.ylabel('Destination')

# Most Common Trip Durations
plt.subplot(2, 3, 2)  # Change to 2, 3, 2 for the second subplot in the first row
sns.barplot(x=duration_counts.values, y=duration_counts.index.astype(str))
plt.title('Most Common Trip Durations (Rounded to 5 min)')
plt.xlabel('Count')
plt.ylabel('Duration (minutes)')

# Destinations in the Top 75% Quantile
plt.subplot(2, 3, 3)  # Change to 2, 3, 3 for the third subplot in the first row
# Calculate the frequency of each destination
destination_counts = valid_destinations['To'].value_counts()
# Determine the 75th percentile value
quantile_75 = destination_counts.quantile(0.75)
# Filter destinations that are at or above the 75th percentile
top_destinations = destination_counts[destination_counts >= quantile_75]
sns.barplot(x=top_destinations.values, y=top_destinations.index)
plt.title('Destinations in the Top 75% Quantile')
plt.xlabel('Count')
plt.ylabel('Destination')


plt.tight_layout()
plt.subplots_adjust(bottom=0)  # Adjust this value as needed to fit tables

plt.show()








