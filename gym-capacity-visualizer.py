import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, time
import base64
from io import BytesIO

# Read the Excel file
df = pd.read_excel('multi_gym_capacity_weather.xlsx')

# Convert Timestamp to datetime
df['Timestamp'] = pd.to_datetime(df['Timestamp'])

# Define the correct order of days
day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

# Extract day of week and hour from Timestamp
df['Day'] = pd.Categorical(df['Day'], categories=day_order, ordered=True)
df['Hour'] = df['Timestamp'].dt.hour

# Calculate occupancy percentage
df['Occupancy'] = df['Count'] / df['Capacity'] * 100

# Function to save plot as base64 string
def plot_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    return base64.b64encode(buf.getvalue()).decode('utf-8')

# 1. Breakdown of busyness at times of each day
plt.figure(figsize=(14, 8))
pivot_table = df.pivot_table(values='Occupancy', index='Hour', columns='Day', aggfunc='mean')
pivot_table = pivot_table.reindex(columns=day_order)
sns.heatmap(pivot_table, cmap='YlOrRd', annot=True, fmt='.1f')
plt.title('Average Occupancy by Day and Hour')
plt.xlabel('Day of Week')
plt.ylabel('Hour of Day')
busyness_chart = plot_to_base64(plt.gcf())
plt.close()

# 2. Quietest days/times
quiet_times = df.groupby(['Day', 'Hour'])['Occupancy'].mean().sort_values().head(10)
plt.figure(figsize=(12, 6))
quiet_times.plot(kind='bar')
plt.title('Top 10 Quietest Times')
plt.xlabel('Day and Hour')
plt.ylabel('Average Occupancy (%)')
plt.xticks(rotation=45, ha='right')
quietest_times_chart = plot_to_base64(plt.gcf())
plt.close()

# 3. Days after working hours where it's the quietest
after_work_hours = range(17, 24)  # 5 PM to 11 PM
after_work_data = df[df['Hour'].isin(after_work_hours)]
quiet_after_work = after_work_data.groupby('Day')['Occupancy'].mean()
quiet_after_work = quiet_after_work.reindex(day_order)  # Reorder the days

plt.figure(figsize=(12, 6))
quiet_after_work.plot(kind='bar')
plt.title('Average Occupancy After Work Hours by Day')
plt.xlabel('Day of Week')
plt.ylabel('Average Occupancy (%)')
plt.xticks(rotation=45, ha='right')
after_work_chart = plot_to_base64(plt.gcf())
plt.close()

# Generate HTML
html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gym Capacity Analysis</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 0; padding: 20px; }}
        h1 {{ color: #333; }}
        img {{ max-width: 100%; height: auto; margin-bottom: 20px; }}
    </style>
</head>
<body>
    <h1>Gym Capacity Analysis</h1>
    
    <h2>Breakdown of Busyness</h2>
    <img src="data:image/png;base64,{busyness_chart}" alt="Busyness Heatmap">
    
    <h2>Quietest Times</h2>
    <img src="data:image/png;base64,{quietest_times_chart}" alt="Quietest Times Chart">
    
    <h2>After Work Hours Occupancy</h2>
    <img src="data:image/png;base64,{after_work_chart}" alt="After Work Hours Chart">
</body>
</html>
"""

# Save the HTML file
with open('gym_capacity_analysis.html', 'w') as f:
    f.write(html_content)

print("Analysis complete. Open gym_capacity_analysis.html to view the charts.")