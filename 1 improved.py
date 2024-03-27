import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import datetime


fuel_economy_data = pd.DataFrame({
    'Vehicle': ['14A', '38A', '46', '47', '51'],
    'Distance': [8710.31, 5178.45, 6246.02, 8350.34, 3075.63],
    'Fuel Used': [4526.72, 1527.08, 2627.37, 3428.09, 2602.14],
    'Fuel Economy': [51.97, 29.49, 42.06, 41.05, 84.61]
})


# load data
def load_data(file_path):
    return pd.read_excel(file_path)

# Data Quality Report Function
def data_quality_report(df):
    missing_values = df.isnull().sum()
    unique_values = df.nunique()
    return pd.DataFrame({'Missing Values': missing_values, 'Unique Values': unique_values})

# Generic Plotting Functions
def plot_bar(data, title, filename):
    plt.figure(figsize=(15, 10))
    data.plot(kind='bar', title=title)
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()

def plot_box(data, title, filename):
    plt.figure(figsize=(10, 6))
    sns.boxplot(x=data)
    plt.title(title)
    plt.savefig(filename)
    plt.close()
    
def plot_avg_median_trip_duration(avg_duration, median_duration, filename):
    fig, ax = plt.subplots()
    ax.bar(['Average Duration', 'Median Duration'], [avg_duration, median_duration])
    plt.title('Average and Median Trip Duration')
    plt.ylabel('Duration (Minutes)')
    plt.savefig(filename)
    plt.close()

def plot_trip_duration_histogram(df, column, filename):
    plt.figure(figsize=(12, 6))
    sns.histplot(df[column], bins=50, kde=True)
    plt.title('Distribution of Trip Durations')
    plt.xlabel('Duration (Minutes)')
    plt.ylabel('Frequency')
    plt.xlim(0, df[column].quantile(0.95))  # Adjust the range if necessary
    plt.savefig(filename)
    plt.close()

# Function to prepare data for grouped bar chart
def prepare_hourly_trip_data(df, start_column, end_column):
    start_hours = df[start_column].dt.hour.value_counts().sort_index()
    end_hours = df[end_column].dt.hour.value_counts().sort_index()
    hourly_data = pd.DataFrame({'Start Hours': start_hours, 'End Hours': end_hours})
    hourly_data = hourly_data.fillna(0)  # Fill missing hours with 0
    return hourly_data

# Function to plot grouped bar chart for trip start and end times
def plot_grouped_bar_chart(data, filename):
    ax = data.plot(kind='bar', figsize=(12, 6), alpha=0.75)
    plt.title('Distribution of Trip Start and End Times')
    plt.xlabel('Hour of Day')
    plt.ylabel('Number of Trips')
    plt.xticks(range(0, 24), rotation=0)
    plt.legend(title='Trip Times')
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()

# Function to calculate odometer differences
def calculate_odometer_differences(df):
    df_sorted = df.sort_values(by=['Vehicle', 'Trip Started'])
    df_sorted['Odometer Difference'] = df_sorted.groupby('Vehicle')['End Odometer'].diff().fillna(0)
    return df_sorted

# Function to plot odometer changes over time by vehicle
def plot_odometer_changes(df, filename):
    plt.figure(figsize=(15, 7))
    for vehicle, data in df.groupby('Vehicle'):
        plt.plot(data['Trip Started'], data['Odometer Difference'], label=vehicle)
    plt.title('Odometer Reading Changes Over Time by Vehicle')
    plt.xlabel('Time')
    plt.ylabel('Odometer Difference (km)')
    plt.legend(title='Vehicle')
    plt.savefig(filename)
    plt.close()
    
def plot_odometer_diff_summary(df, filename):
    plt.figure(figsize=(12, 6))

    # Preparing data for boxplot
    data_to_plot = [group['Odometer Difference'].values for _, group in df.groupby('Vehicle')]
    vehicle_names = df['Vehicle'].unique()

    # Define colors for each vehicle
    colors = ['blue', 'green', 'red', 'purple', 'orange']  # Extend or reduce this list based on the number of vehicles

    # Create a boxplot for each vehicle
    boxprops = dict(linestyle='-', linewidth=1)
    medianprops = dict(linestyle='-', linewidth=2, color='firebrick')
    for i, data in enumerate(data_to_plot):
        color = colors[i % len(colors)]  # Cycle through colors if there are more vehicles than colors
        plt.boxplot(data, positions=[i + 1], patch_artist=True, boxprops=dict(facecolor=color, **boxprops), medianprops=medianprops, showfliers=True)

    plt.title('Odometer Differences Summary Statistics by Vehicle')
    plt.xticks(range(1, len(vehicle_names) + 1), vehicle_names)
    plt.ylabel('Odometer Difference (km)')
    plt.savefig(filename)
    plt.close()

# Function to plot Efficiency Analysis
def plot_efficiency_analysis(df, filename):
    plt.figure(figsize=(12, 6))

    # Using Seaborn to create the scatter plot
    sns.scatterplot(data=df, x='Total Duration (Minutes)', y='Maximum Speed', 
                    hue='Idling Duration (Minutes)', size='Idling Duration (Minutes)', 
                    sizes=(20, 180), alpha=0.5)

    plt.title('Efficiency Analysis: Total Duration vs. Maximum Speed with Idling Duration')
    plt.xlabel('Total Duration (Minutes)')
    plt.ylabel('Maximum Speed (km/h)')
    plt.legend(title='Idling Duration (Minutes)', bbox_to_anchor=(0.9, 1), loc=2)
    plt.savefig(filename)
    plt.close()

# Outlier Detection Function
def detect_outliers(df, col):
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    return df[(df[col] < lower_bound) | (df[col] > upper_bound)]

# Time Conversion Function
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

def convert_time_columns(df):
    df['Trip Started'] = pd.to_datetime(df['Trip Started'])
    df['Trip Ended'] = pd.to_datetime(df['Trip Ended'])
    df['Trip Duration (Minutes)'] = (df['Trip Ended'] - df['Trip Started']).dt.total_seconds() / 60
    df['Stop Duration (Minutes)'] = df['Stop Duration'].apply(convert_to_minutes)
    df['Driving Duration (Minutes)'] = df['Driving Duration'].apply(convert_to_minutes)  # Ensure this conversion
    df['Idling Duration (Minutes)'] = df['Idling Duration'].apply(convert_to_minutes)

# Trip Analysis Functions
def trip_analysis(df):
    # Use the converted 'Driving Duration (Minutes)' for calculation
    trip_frequency = df['Vehicle'].value_counts()
    avg_duration = df['Driving Duration (Minutes)'].mean()  # Updated column name
    median_duration = df['Driving Duration (Minutes)'].median()  # Updated column name
    return trip_frequency, avg_duration, median_duration

# Location Analysis Function
def location_analysis(df):
    most_common_start = df.groupby('Vehicle')['Start Location'].agg(pd.Series.mode)
    most_common_end = df.groupby('Vehicle')['End Location'].agg(pd.Series.mode)
    return pd.DataFrame({'Most Common Start Location': most_common_start, 'Most Common End Location': most_common_end})


# Distance Covered Analysis
def calculate_distance_covered(df):
    df['Date'] = pd.to_datetime(df['Trip Started']).dt.date
    df['Distance Covered'] = df['End Odometer'] - df['Start Odometer']
    daily = df.groupby(['Vehicle', 'Date'])['Distance Covered'].sum()
    weekly = df.groupby(['Vehicle', pd.Grouper(key='Trip Started', freq='W-MON')])['Distance Covered'].sum()
    monthly = df.groupby(['Vehicle', pd.Grouper(key='Trip Started', freq='M')])['Distance Covered'].sum()
    yearly = df.groupby(['Vehicle', pd.Grouper(key='Trip Started', freq='Y')])['Distance Covered'].sum()
    return daily, weekly, monthly, yearly

# Plot Daily Distance Covered
def plot_daily_distance_covered(data, filename):
    plt.figure(figsize=(15, 7))
    for vehicle in data.index.get_level_values(0).unique():
        vehicle_data = data.xs(vehicle, level='Vehicle')
        plt.plot(vehicle_data.index, vehicle_data.values, label=f'Vehicle {vehicle}')
    plt.title('Daily Distance Covered per Vehicle')
    plt.xlabel('Date')
    plt.ylabel('Distance Covered (km)')
    plt.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()

# Plot Weekly/Monthly/Yearly Distance Covered
def plot_aggregated_distance_covered(data, title, filename):
    plt.figure(figsize=(15, 7))
    data.unstack(level=0).plot(kind='bar', stacked=True)
    plt.title(title)
    plt.xlabel('Time Period')
    plt.ylabel('Distance Covered (km)')
    plt.legend(title='Vehicle')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()

# Fuel Economy Analysis
# Merge with Main Dataset
def merge_fuel_data(main_df, fuel_df):
    return pd.merge(main_df, fuel_df, on='Vehicle', how='left')

# Plotting Average Fuel Economy per Vehicle
def plot_avg_fuel_economy(data, filename):
    plt.figure(figsize=(10, 6))
    sns.barplot(x=data.index, y=data.values)
    plt.title('Average Fuel Economy per Vehicle')
    plt.xlabel('Vehicle')
    plt.ylabel('Average Fuel Economy (km/l)')
    plt.savefig(filename)
    plt.close()
    
# Plotting Distance Covered vs. Fuel Used per Vehicle
def plot_distance_vs_fuel_used(distance_data, fuel_used_data, filename):
    plt.figure(figsize=(10, 6))
    sns.scatterplot(x=distance_data.values, y=fuel_used_data.values)  # Ensure both are Series
    for i, txt in enumerate(distance_data.index):
        plt.annotate(txt, (distance_data.iloc[i], fuel_used_data.iloc[i]))  # Use iloc for positional indexing
    plt.title('Distance Covered vs. Fuel Used per Vehicle')
    plt.xlabel('Total Distance Covered (km)')
    plt.ylabel('Total Fuel Used (L)')
    plt.savefig(filename)
    plt.close()

# Plotting Fuel Economy Trend Over Time
def plot_fuel_economy_trend(data, filename):
    plt.figure(figsize=(12, 8))
    for vehicle in data.index.get_level_values(0).unique():
        vehicle_data = data.loc[vehicle]
        sns.lineplot(x=vehicle_data.index.astype('str'), y=vehicle_data.values, label=vehicle)  # Convert index to string for plotting
    plt.title('Fuel Economy Trend Over Time')
    plt.xlabel('Month')
    plt.ylabel('Fuel Economy (km/l)')
    plt.xticks(rotation=45)
    plt.legend(title='Vehicle')
    plt.savefig(filename)
    plt.close()


# Main Analysis Script
def main():
    sns.set(style="whitegrid")

    new_vehicle_data = load_data('0 Combined shortcut.xlsx')
    # Convert Time Columns
    convert_time_columns(new_vehicle_data)
    # Prepare and merge fuel economy data
    merged_data = merge_fuel_data(new_vehicle_data, fuel_economy_data)

    # Data Quality Report
    data_quality_summary = data_quality_report(new_vehicle_data)
    plot_bar(data_quality_summary['Missing Values'], 'Missing Values per Column', 'data_quality_missing_values.png')
    plot_bar(data_quality_summary['Unique Values'], 'Unique Values per Column', 'data_quality_unique_values.png')

    # Outliers Detection and Visualization
    outliers_end_odometer = detect_outliers(new_vehicle_data, 'End Odometer')
    plot_box(new_vehicle_data['End Odometer'], 'Boxplot of End Odometer', 'End_Odometer_outliers.png')

    # Correctly plotting the converted 'Stop Duration (Minutes)' column
    plot_box(new_vehicle_data['Stop Duration (Minutes)'], 'Boxplot of Stop Duration', 'Stop_Duration_Boxplot.png')

    # Trip Frequency and Duration Analysis
    trip_freq, avg_trip_duration, med_trip_duration = trip_analysis(new_vehicle_data)
    print(trip_freq, avg_trip_duration, med_trip_duration)
    plot_bar(trip_freq, 'Trip Frequency per Vehicle', 'trip_frequency_per_vehicle.png')
    
    # Plotting Average and Median Trip Duration
    plot_avg_median_trip_duration(avg_trip_duration, med_trip_duration, 'avg_median_trip_duration.png')
    
    # Plotting Trip Durations Histogram
    plot_trip_duration_histogram(new_vehicle_data, 'Trip Duration (Minutes)', 'trip_durations_histogram.png')

    # Prepare Data for Grouped Bar Chart
    hourly_trip_data = prepare_hourly_trip_data(new_vehicle_data, 'Trip Started', 'Trip Ended')

    # Plotting Grouped Bar Chart for Trip Start and End Times
    plot_grouped_bar_chart(hourly_trip_data, 'trip_start_end_times_grouped_bar.png')

    # Prepare Data for Odometer Changes Plot and Summary
    odometer_data = calculate_odometer_differences(new_vehicle_data)

    # Plotting Odometer Changes Over Time by Vehicle
    plot_odometer_changes(odometer_data, 'odometer_changes_over_time.png')

    # Plotting Odometer Differences Summary Statistics by Vehicle with Colors
    new_vehicle_data['Total Duration (Minutes)'] = new_vehicle_data['Driving Duration (Minutes)'] + new_vehicle_data['Idling Duration (Minutes)']
    plot_odometer_diff_summary(odometer_data, 'odometer_diff_summary_by_vehicle.png')

    # Plotting Efficiency Analysis
    plot_efficiency_analysis(new_vehicle_data, 'efficiency_analysis_scatterplot.png')

    # Location-Based Analysis
    location_summary = location_analysis(new_vehicle_data)
    print(location_summary)    
    
    # Convert Odometer readings to numeric values if not already
    new_vehicle_data['Start Odometer'] = pd.to_numeric(new_vehicle_data['Start Odometer'], errors='coerce')
    new_vehicle_data['End Odometer'] = pd.to_numeric(new_vehicle_data['End Odometer'], errors='coerce')

    # Distance Covered Analysis
    daily_dist, weekly_dist, monthly_dist, yearly_dist = calculate_distance_covered(new_vehicle_data)
    print(daily_dist, weekly_dist, monthly_dist, yearly_dist)
    # You can add code here to print these results or plot them as needed.
    plot_daily_distance_covered(daily_dist, 'daily_distance_covered.png')
    plot_aggregated_distance_covered(weekly_dist, 'Weekly Distance Covered per Vehicle', 'weekly_distance_covered.png')
    plot_aggregated_distance_covered(monthly_dist, 'Monthly Distance Covered per Vehicle', 'monthly_distance_covered.png')
    plot_aggregated_distance_covered(yearly_dist, 'Yearly Distance Covered per Vehicle', 'yearly_distance_covered.png')
    
    # Analysis 1: Average Fuel Economy per Vehicle
    avg_fuel_economy = merged_data.groupby('Vehicle')['Fuel Economy'].mean()
    plot_avg_fuel_economy(avg_fuel_economy, 'avg_fuel_economy_per_vehicle.png')

    # Analysis 2: Distance Covered vs. Fuel Used per Vehicle
    merged_data['Distance Covered'] = merged_data['End Odometer'] - merged_data['Start Odometer']
    total_distance = merged_data.groupby('Vehicle')['Distance Covered'].sum()
    fuel_used_data = merged_data.groupby('Vehicle')['Fuel Used'].sum()
    plot_distance_vs_fuel_used(total_distance, fuel_used_data, 'distance_vs_fuel_used_per_vehicle.png')

    # Analysis 3: Fuel Economy Trend Over Time
    merged_data['Month'] = pd.to_datetime(merged_data['Trip Started']).dt.to_period('M')
    fuel_economy_trend = merged_data.groupby(['Vehicle', 'Month'])['Fuel Economy'].mean()
    plot_fuel_economy_trend(fuel_economy_trend, 'fuel_economy_trend_over_time.png')


if __name__ == "__main__":
    main()

