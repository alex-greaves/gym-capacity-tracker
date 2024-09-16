import requests
import re
import pandas as pd
from datetime import datetime
import time
import schedule

def get_capacity():
    url = "https://portal.rockgympro.com/portal/public/a9c7d8e770832d2ce2a1a5b371f95dfb/occupancy?&iframeid=occupancyCounter&fId="  # Make sure this is correct
    print(f"Fetching data from {url}")
    response = requests.get(url)
    print(f"Response status code: {response.status_code}")
    
    # Use regex to find the data variable in the JavaScript
    data_match = re.search(r'var data = (\{[^;]+\});', response.text)
    
    if data_match:
        data_str = data_match.group(1)
        print(f"Found data: {data_str}")
        # Extract the count and capacity using regex
        count_match = re.search(r"'count'\s*:\s*(\d+)", data_str)
        capacity_match = re.search(r"'capacity'\s*:\s*(\d+)", data_str)
        
        if count_match and capacity_match:
            count = int(count_match.group(1))
            capacity = int(capacity_match.group(1))
            print(f"Extracted count: {count}, capacity: {capacity}")
            return count, capacity
    
    print("Could not extract capacity data from the webpage")
    raise ValueError("Could not extract capacity data from the webpage")

def update_excel():
    print("Updating Excel file...")
    try:
        count, capacity = get_capacity()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            df = pd.read_excel('gym_capacity.xlsx')
            print("Existing Excel file found")
        except FileNotFoundError:
            print("Creating new Excel file")
            df = pd.DataFrame(columns=['Timestamp', 'Count', 'Capacity'])
        
        new_row = pd.DataFrame({'Timestamp': [timestamp], 'Count': [count], 'Capacity': [capacity]})
        df = pd.concat([df, new_row], ignore_index=True)
        
        df.to_excel('gym_capacity.xlsx', index=False)
        print(f"Updated at {timestamp}: Count = {count}, Capacity = {capacity}")
    except Exception as e:
        print(f"Error updating data: {e}")

def main():
    print("Starting the gym capacity tracker...")
    schedule.every().hour.do(update_excel)
    print("Scheduled to update every hour")
    
    update_excel()  # Run once immediately
    
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    main()