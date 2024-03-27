import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

data = pd.read_excel('modified_dates.xlsx')

# Convert 'Date' to datetime format
data['Date'] = pd.to_datetime(data['Date'], errors='coerce', format='%Y-%m-%d')

# Convert 'Week' and 'Month' to string for consistent index type
data['Week'] = data['Week'].astype(str)
data['Month'] = data['Month'].astype(str)

# Vehicle columns
vehicle_columns = ['#14', '#38A', '#46', '#47', '#51']

# Adding a helper column for easier calculation of visits and driving time
data['Visit'] = data[vehicle_columns].sum(axis=1)
data['DrivingTime'] = pd.to_timedelta(data['Time for Call'])

# Filtering only rows where a vehicle was used (Visit > 0)
vehicle_data = data[data['Visit'] > 0]
vehicle_data['DateString'] = vehicle_data['Date'].dt.strftime('%Y-%m-%d')

# Prepare the Excel writer for exporting results
with pd.ExcelWriter('Vehicle_Analysis.xlsx') as writer:
    for vehicle in vehicle_columns:
        vehicle_specific_data = vehicle_data[vehicle_data[vehicle] > 0]
        
        # Analysis
        daily_visits = vehicle_specific_data['DateString'].value_counts().sort_index()
        weekly_visits = vehicle_specific_data['Week'].value_counts().sort_index()
        monthly_visits = vehicle_specific_data['Month'].value_counts().sort_index()
        total_driving_time_daily = vehicle_specific_data.groupby('DateString')['DrivingTime'].sum()
        total_driving_time_weekly = vehicle_specific_data.groupby('Week')['DrivingTime'].sum()
        total_driving_time_monthly = vehicle_specific_data.groupby('Month')['DrivingTime'].sum()
        
        # Exclude 'DOCK' and count top 10 locations for each period
        valid_destinations = vehicle_specific_data[vehicle_specific_data['To'] != 'DOCK']

        top_daily_locations = valid_destinations.groupby('DateString')['To'].apply(lambda x: x.value_counts().head(10))
        top_weekly_locations = valid_destinations.groupby('Week')['To'].apply(lambda x: x.value_counts().head(10))
        top_monthly_locations = valid_destinations.groupby('Month')['To'].apply(lambda x: x.value_counts().head(10))
        
        # Unique place visits
        unique_daily_visits = valid_destinations.groupby('DateString')['To'].nunique()
        unique_weekly_visits = valid_destinations.groupby('Week')['To'].nunique()
        unique_monthly_visits = valid_destinations.groupby('Month')['To'].nunique()

        # Save analysis and top locations to Excel
        combined_df = pd.DataFrame({
            'Daily Visits': daily_visits,
            'Unique Daily Visits': unique_daily_visits,
            'Weekly Visits': weekly_visits,
            'Unique Weekly Visits': unique_weekly_visits,
            'Monthly Visits': monthly_visits,
            'Unique Monthly Visits': unique_monthly_visits,
            'Total Driving Time Daily (hours)': total_driving_time_daily / pd.Timedelta(hours=1),
            'Total Driving Time Weekly (hours)': total_driving_time_weekly / pd.Timedelta(hours=1),
            'Total Driving Time Monthly (hours)': total_driving_time_monthly / pd.Timedelta(hours=1),
        }).fillna(0)
        combined_df.to_excel(writer, sheet_name=f'{vehicle}_Analysis')

        top_daily_locations.to_excel(writer, sheet_name=f'{vehicle}_Top_Daily_Locations')
        top_weekly_locations.to_excel(writer, sheet_name=f'{vehicle}_Top_Weekly_Locations')
        top_monthly_locations.to_excel(writer, sheet_name=f'{vehicle}_Top_Monthly_Locations')
        
        valid_destinations = vehicle_specific_data[vehicle_specific_data['To'] != 'DOCK']
        
        # Sort the top locations in decreasing order
        top_daily_locations = valid_destinations.groupby('DateString')['To'].apply(lambda x: x.value_counts().head(3)).sort_values(ascending=False).reset_index(level=0, drop=True)
        top_weekly_locations = valid_destinations.groupby('Week')['To'].apply(lambda x: x.value_counts().head(5)).sort_values(ascending=False).reset_index(level=0, drop=True)
        top_monthly_locations = valid_destinations.groupby('Month')['To'].apply(lambda x: x.value_counts().head(10)).sort_values(ascending=False).reset_index(level=0, drop=True)
    
        # Total Driving Time for Daily, Weekly, Monthly in hours
        total_daily_driving_time_hours = vehicle_specific_data.groupby('DateString')['DrivingTime'].sum() / pd.Timedelta(hours=1)
        total_weekly_driving_time_hours = vehicle_specific_data.groupby('Week')['DrivingTime'].sum() / pd.Timedelta(hours=1)
        total_monthly_driving_time_hours = vehicle_specific_data.groupby('Month')['DrivingTime'].sum() / pd.Timedelta(hours=1)
    
        # Visualization
        # Top 10 Visited Locations for Daily, Weekly, Monthly
        plt.figure(figsize=(15, 9))
        sns.barplot(y=top_daily_locations.index, x=top_daily_locations.values)
        plt.title(f'Daily Top Locations for {vehicle}')
        plt.savefig(f'{vehicle}_Daily_Top_Locations.png')
        plt.close()  # Close the plot to free up memory

        plt.figure(figsize=(15, 9))
        sns.barplot(y=top_weekly_locations.index, x=top_weekly_locations.values)
        plt.title(f'Weekly Top Locations for {vehicle}')
        plt.savefig(f'{vehicle}_Weekly_Top_Locations.png')
        plt.close()  # Close the plot to free up memory

        plt.figure(figsize=(15, 9))
        sns.barplot(y=top_monthly_locations.index, x=top_monthly_locations.values)
        plt.title(f'Monthly Top Locations for {vehicle}')
        plt.savefig(f'{vehicle}_Monthly_Top_Locations.png')
        plt.close()  # Close the plot to free up memory
        
        
        plt.figure(figsize=(15, 9))
        total_daily_driving_time_hours.plot(kind='line')
        plt.title(f'Daily Total Driving Time (hrs) for {vehicle}')
        plt.savefig(f'{vehicle}_Daily_Total_Driving_Time.png')
        plt.close()

        plt.figure(figsize=(15, 9))
        total_weekly_driving_time_hours.plot(kind='line')
        plt.title(f'Weekly Total Driving Time (hrs) for {vehicle}')
        plt.savefig(f'{vehicle}_Weekly_Total_Driving_Time.png')
        plt.close()

        plt.figure(figsize=(15, 9))
        total_monthly_driving_time_hours.plot(kind='line').invert_xaxis()
        plt.title(f'Monthly Total Driving Time (hrs) for {vehicle}')
        plt.savefig(f'{vehicle}_Monthly_Total_Driving_Time.png')
        plt.close()
        
        
        plt.figure(figsize=(15, 9))
        daily_visits.plot(kind='line')
        plt.title(f'Daily Visits for {vehicle}')
        plt.savefig(f'{vehicle}_Daily_Visits.png')
        plt.close()

        plt.figure(figsize=(15, 9))
        weekly_visits.plot(kind='line')
        plt.title(f'Weekly Visits for {vehicle}')
        plt.savefig(f'{vehicle}_Weekly_Visits.png')
        plt.close()

        plt.figure(figsize=(15, 9))
        monthly_visits.plot(kind='line').invert_xaxis()
        plt.title(f'Monthly Visits for {vehicle}')
        plt.savefig(f'{vehicle}_Monthly_Visits.png')
        plt.close()
        

        plt.figure(figsize=(15, 9))
        unique_daily_visits.plot(kind='line')
        plt.title(f'Daily Unique Visits for {vehicle}')
        plt.savefig(f'{vehicle}_Daily_Unique_Visits.png')
        plt.close()

        plt.figure(figsize=(15, 9))
        unique_weekly_visits.plot(kind='line')
        plt.title(f'Weekly Unique Visits for {vehicle}')
        plt.savefig(f'{vehicle}_Weekly_Unique_Visits.png')
        plt.close()

        plt.figure(figsize=(15, 9))
        unique_monthly_visits.plot(kind='line').invert_xaxis()
        plt.title(f'Monthly Unique Visits for {vehicle}')
        plt.savefig(f'{vehicle}_Monthly_Unique_Visits.png')
        plt.close()


print("Analysis data has been saved to 'Vehicle_Analysis.xlsx'")
print("Comprehensive visualizations have been saved.")


    
    


