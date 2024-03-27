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
    
        fig, axes = plt.subplots(nrows=4, ncols=3, figsize=(30, 30)) # Adjust the subplot layout

        # Top 10 Visited Locations for Daily, Weekly, Monthly
        sns.barplot(y=top_daily_locations.index, x=top_daily_locations.values, ax=axes[0, 0])
        sns.barplot(y=top_weekly_locations.index, x=top_weekly_locations.values, ax=axes[0, 1])
        sns.barplot(y=top_monthly_locations.index, x=top_monthly_locations.values, ax=axes[0, 2])
        axes[0, 0].set_title(f'Daily Top 10 Locations for {vehicle}')
        axes[0, 1].set_title(f'Weekly Top 10 Locations for {vehicle}')
        axes[0, 2].set_title(f'Monthly Top 10 Locations for {vehicle}')

        # Total Driving Time for Daily, Weekly, Monthly
        total_daily_driving_time_hours.plot(ax=axes[1, 0], kind='line')
        total_weekly_driving_time_hours.plot(ax=axes[1, 1], kind='line')
        total_monthly_driving_time_hours.plot(ax=axes[1, 2], kind='line').invert_xaxis()
        axes[1, 0].set_title(f'Daily Total Driving Time (hrs) for {vehicle}')
        axes[1, 1].set_title(f'Weekly Total Driving Time (hrs) for {vehicle}')
        axes[1, 2].set_title(f'Monthly Total Driving Time (hrs) for {vehicle}')

        # Number of Visits for Daily, Weekly, Monthly
        daily_visits.plot(ax=axes[2, 0], kind='line')
        weekly_visits.plot(ax=axes[2, 1], kind='line')
        monthly_visits.plot(ax=axes[2, 2], kind='line').invert_xaxis()
        axes[2, 0].set_title(f'Daily Visits for {vehicle}')
        axes[2, 1].set_title(f'Weekly Visits for {vehicle}')
        axes[2, 2].set_title(f'Monthly Visits for {vehicle}')
        
        # Add plots for unique visits
        unique_daily_visits.plot(ax=axes[3, 0], kind='line')
        unique_weekly_visits.plot(ax=axes[3, 1], kind='line')
        unique_monthly_visits.plot(ax=axes[3, 2], kind='line').invert_xaxis()
        axes[3, 0].set_title(f'Daily Unique Visits for {vehicle}')
        axes[3, 1].set_title(f'Weekly Unique Visits for {vehicle}')
        axes[3, 2].set_title(f'Monthly Unique Visits for {vehicle}')

        plt.tight_layout()

        # Save each comprehensive figure with a unique name
        plt.savefig(f'{vehicle}_comprehensive_visualization.png')
        plt.close(fig)


print("Analysis data has been saved to 'Vehicle_Analysis.xlsx'")
print("Comprehensive visualizations have been saved.")


    
    


