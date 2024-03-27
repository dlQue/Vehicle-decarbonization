import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import datetime

new_file_path = '0 Combined shortcut.xlsx'
new_vehicle_data = pd.read_excel(new_file_path)
#print(new_vehicle_data.head())

# Performing a basic data quality check for missing values and unique value counts
missing_values = new_vehicle_data.isnull().sum()
unique_values = new_vehicle_data.nunique()

# Creating a summary report of the data quality
data_quality_summary = pd.DataFrame({'Missing Values': missing_values, 'Unique Values': unique_values})
#print(data_quality_summary)
# Visualizing data quality summary
fig, ax = plt.subplots(2, 1, figsize=(15, 10))
data_quality_summary['Missing Values'].plot(kind='bar', ax=ax[0], title='Missing Values per Column')
data_quality_summary['Unique Values'].plot(kind='bar', ax=ax[1], title='Unique Values per Column')
plt.tight_layout()
plt.savefig('data_quality_summary.png')
#plt.show()

# Define a function to detect and return outliers
def detect_outliers(df, col):
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    outliers = df[(df[col] < lower_bound) | (df[col] > upper_bound)]
    return outliers

# Detect outliers in 'Trip Duration (Minutes)'
outliers_trip_duration = detect_outliers(new_vehicle_data, 'End Odometer')

# Display outliers
print("Outliers in End Odometer:")
print(outliers_trip_duration)

# Visualize outliers using a boxplot
plt.figure(figsize=(10, 6))
sns.boxplot(x=new_vehicle_data['End Odometer'])
plt.title('Boxplot of End Odometer')
plt.savefig('End_Odometer_outliers.png')
plt.show()

# Adjusted function to convert time string or time object to total minutes
def convert_to_minutes(time_value):
    if pd.isna(time_value):
        return 0
    if isinstance(time_value, str):
        time_parts = datetime.datetime.strptime(time_value, '%H:%M:%S.%f')
    elif isinstance(time_value, datetime.time):
        time_parts = time_value
    else:
        return 0
    total_minutes = time_parts.hour * 60 + time_parts.minute + time_parts.second / 60
    return total_minutes

# Ensure 'Stop Duration' is converted to a numeric format (total minutes)
new_vehicle_data['Stop Duration (Minutes)'] = new_vehicle_data['Stop Duration'].apply(convert_to_minutes)

# Now detect outliers in 'Stop Duration (Minutes)'
outliers_stop_duration = detect_outliers(new_vehicle_data, 'Stop Duration (Minutes)')

# Display outliers
print("Outliers in Stop Duration (Minutes):")
print(outliers_stop_duration)

# Visualize outliers using a boxplot
plt.figure(figsize=(10, 6))
sns.boxplot(x=new_vehicle_data['Stop Duration (Minutes)'])
plt.title('Boxplot of Stop Duration (Minutes)')
plt.savefig('stop_duration_outliers.png')
plt.show()




# Converting 'Trip Started' and 'Trip Ended' to datetime for duration calculation
new_vehicle_data['Trip Started'] = pd.to_datetime(new_vehicle_data['Trip Started'])
new_vehicle_data['Trip Ended'] = pd.to_datetime(new_vehicle_data['Trip Ended'])

# Calculating trip duration in minutes
new_vehicle_data['Trip Duration (Minutes)'] = (new_vehicle_data['Trip Ended'] - new_vehicle_data['Trip Started']).dt.total_seconds() / 60

# Analysis of trip frequency per vehicle
trip_frequency_per_vehicle = new_vehicle_data['Vehicle'].value_counts()

# Analysis of average and median trip duration
average_trip_duration = new_vehicle_data['Trip Duration (Minutes)'].mean()
median_trip_duration = new_vehicle_data['Trip Duration (Minutes)'].median()

print(trip_frequency_per_vehicle, average_trip_duration, median_trip_duration)

# Visualizing trip frequency per vehicle
plt.figure(figsize=(12, 6))
trip_frequency_per_vehicle.plot(kind='bar', title='Trip Frequency per Vehicle')
plt.xlabel('Vehicle')
plt.ylabel('Number of Trips')
plt.savefig('trip_frequency_per_vehicle.png')
plt.show()

# Visualizing average and median trip duration
fig, ax = plt.subplots()
ax.bar(['Average Duration', 'Median Duration'], [average_trip_duration, median_trip_duration])
plt.title('Average and Median Trip Duration')
plt.ylabel('Duration (Minutes)')
plt.savefig('avg_median_trip_duration.png')
plt.show()



# Setting the style for the plots
sns.set(style="whitegrid")

# For the histogram of trip durations
plt.figure(figsize=(12, 6))
sns.histplot(new_vehicle_data['Trip Duration (Minutes)'], bins=50, kde=True)
plt.title('Distribution of Trip Durations')
plt.xlabel('Duration (Minutes)')
plt.ylabel('Frequency')
plt.xlim(0, new_vehicle_data['Trip Duration (Minutes)'].quantile(0.95))
plt.savefig('trip_durations_histogram.png')  # Save the figure
plt.show()

# For the histogram of trip start times
# Converting 'Trip Started' to datetime
new_vehicle_data['Trip Started'] = pd.to_datetime(new_vehicle_data['Trip Started'])

# Extracting the hour from 'Trip Started'
new_vehicle_data['Hour of Day Started'] = new_vehicle_data['Trip Started'].dt.hour

# Now you can create the histogram for 'Hour of Day Started'
plt.figure(figsize=(12, 6))
sns.histplot(new_vehicle_data['Hour of Day Started'], bins=24, kde=False)
plt.title('Distribution of Trip Start Times')
plt.xlabel('Hour of Day')
plt.ylabel('Number of Trips')
plt.xticks(range(0, 24))
plt.savefig('trip_start_times_histogram.png')  # Save the figure
plt.show()




# Reapplying the conversion
new_vehicle_data['Driving Duration (Minutes)'] = new_vehicle_data['Driving Duration'].apply(convert_to_minutes)
new_vehicle_data['Idling Duration (Minutes)'] = new_vehicle_data['Idling Duration'].apply(convert_to_minutes)

# Recalculating total duration
new_vehicle_data['Total Duration (Minutes)'] = new_vehicle_data['Driving Duration (Minutes)'] + new_vehicle_data['Idling Duration (Minutes)']

# Reattempting the efficiency analysis plot
plt.figure(figsize=(12, 6))
sns.scatterplot(data=new_vehicle_data, x='Total Duration (Minutes)', y='Maximum Speed', hue='Idling Duration (Minutes)', size='Idling Duration (Minutes)', sizes=(20, 180), alpha=0.5)
plt.title('Efficiency Analysis: Total Duration vs. Maximum Speed with Idling Duration')
plt.xlabel('Total Duration (Minutes)')
plt.ylabel('Maximum Speed (km/h)')
plt.legend(title='Idling Duration (Minutes)', bbox_to_anchor=(0.9, 1), loc=2)
plt.savefig('efficiency_analysis_scatterplot.png')  # Save the figure
plt.show()


# Work Hours Utilization Analysis

# Assuming standard work hours to be from 9:00 to 17:00 (9 AM to 5 PM)
work_hours_start = 9
work_hours_end = 17

# Function to check if a time is within work hours
def is_within_work_hours(time):
    return work_hours_start <= time.hour < work_hours_end

# Checking if trip starts or ends during work hours
new_vehicle_data['Start During Work Hours'] = new_vehicle_data['Trip Started'].apply(is_within_work_hours)
new_vehicle_data['Stop During Work Hours'] = new_vehicle_data['Trip Ended'].apply(is_within_work_hours)

# Calculating the proportion of trips starting and ending during work hours
start_during_work_hours_proportion = new_vehicle_data['Start During Work Hours'].mean()
stop_during_work_hours_proportion = new_vehicle_data['Stop During Work Hours'].mean()

print(start_during_work_hours_proportion, stop_during_work_hours_proportion)

# Visualizing proportion of trips during work hours
fig, ax = plt.subplots(1, 1, figsize=(8, 6))
ax.bar(['Start During Work Hours', 'Stop During Work Hours'], [start_during_work_hours_proportion, stop_during_work_hours_proportion])
ax.set_title('Proportion of Trips During Work Hours')
ax.set_ylabel('Proportion')
plt.tight_layout()
plt.savefig('work_hours_proportions.png')
plt.show()




# Adjusting the approach to calculate odometer differences
# Sorting the data by vehicle and trip start time
sorted_vehicle_data = new_vehicle_data.sort_values(by=['Vehicle', 'Trip Started'])

# Calculating the difference in odometer readings for each trip, per vehicle
sorted_vehicle_data['Odometer Difference'] = sorted_vehicle_data.groupby('Vehicle')['End Odometer'].diff().fillna(0)

# Plotting the odometer difference for each vehicle over time
plt.figure(figsize=(15, 7))
for vehicle, data in sorted_vehicle_data.groupby('Vehicle'):
    plt.plot(data['Trip Started'], data['Odometer Difference'], label=vehicle)
plt.title('Odometer Reading Changes Over Time by Vehicle')
plt.xlabel('Time')
plt.ylabel('Odometer Difference (km)')
plt.legend(title='Vehicle')
plt.savefig('odometer_difference_plot.png')  # Save the figure
plt.show()

# Summary statistics for odometer differences
odometer_diff_summary = sorted_vehicle_data.groupby('Vehicle')['Odometer Difference'].describe()
print(odometer_diff_summary)

# Visualizing odometer differences summary statistics with different colors for each vehicle
plt.figure(figsize=(12, 6))
colors = ['blue', 'green', 'red', 'purple', 'orange']  # Adjust colors as needed

for (i, (vehicle, data)) in enumerate(odometer_diff_summary.iterrows()):
    plt.boxplot(data, positions=[i + 1], patch_artist=True, boxprops=dict(facecolor=colors[i]))

plt.title('Odometer Differences Summary Statistics by Vehicle')
plt.xticks(range(1, len(odometer_diff_summary) + 1), odometer_diff_summary.index)
plt.ylabel('Odometer Difference (km)')
plt.savefig('odometer_diff_stats_by_vehicle.png')
plt.show()



# Location-Based Analysis: Most common start and end locations
# Identifying the most common start and end locations for each vehicle
most_common_start_locations = sorted_vehicle_data.groupby('Vehicle')['Start Location'].agg(pd.Series.mode)
most_common_end_locations = sorted_vehicle_data.groupby('Vehicle')['End Location'].agg(pd.Series.mode)

# Creating a summary DataFrame
location_summary = pd.DataFrame({
    'Most Common Start Location': most_common_start_locations,
    'Most Common End Location': most_common_end_locations
})

print(location_summary)



