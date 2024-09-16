import requests
import re
import pandas as pd
from datetime import datetime
import time
import schedule
import os

# Replace with your OpenWeatherMap API key
WEATHER_API_KEY = "YOUR_API_KEY_HERE"
# Replace with your gym's latitude and longitude
GYM_LAT = "37.7749"
GYM_LON = "-122.4194"

def get_capacity():
    url = "https://portal.rockgympro.com/portal/public/a9c7d8e770832d2ce2a1a5b371f95dfb/occupancy?&iframeid=occupancyCounter&fId="  # Make sure this is correct
    print(f"Fetching data from {url}")
    response = requests.get(url)
    print(f"Response status code: {response.status_code}")
    
    data_match = re.search(r'var data = (\{[^;]+\});', response.text)
    
    if data_match:
        data_str = data_match.group(1)
        print(f"Found data: {data_str}")
        count_match = re.search(r"'count'\s*:\s*(\d+)", data_str)
        capacity_match = re.search(r"'capacity'\s*:\s*(\d+)", data_str)

        if count_match and capacity_match:
            count = int(count_match.group(1))
            capacity = int(capacity_match.group(1))
            print(f"Extracted count: {count}, capacity: {capacity}")
            return count, capacity

    print("Could not extract capacity data from the webpage")
    raise ValueError("Could not extract capacity data from the webpage")

def get_weather():
    url = f"http://api.openweathermap.org/data/2.5/weather?lat={GYM_LAT}&lon={GYM_LON}&appid={WEATHER_API_KEY}&units=metric"
    response = requests.get(url)
    data = response.json()

    if response.status_code == 200:
        weather_main = data['weather'][0]['main']
        temperature = data['main']['temp']
        return weather_main, temperature
    else:
        print(f"Error fetching weather data: {response.status_code}")
        return None, None

def update_excel():
    print("Updating Excel file...")
    try:
        count, capacity = get_capacity()
        weather_main, temperature = get_weather()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            df = pd.read_excel('gym_capacity_weather.xlsx')
            print("Existing Excel file found")
        except FileNotFoundError:
            print("Creating new Excel file")
            df = pd.DataFrame(columns=['Timestamp', 'Count', 'Capacity', 'Weather', 'Temperature'])
        
        new_row = pd.DataFrame({
            'Timestamp': [timestamp],
            'Count': [count],
            'Capacity': [capacity],
            'Weather': [weather_main],
            'Temperature': [temperature]
        })
        df = pd.concat([df, new_row], ignore_index=True)
        
        df.to_excel('gym_capacity_weather.xlsx', index=False)
        print(f"Updated at {timestamp}: Count = {count}, Capacity = {capacity}, Weather = {weather_main}, Temp = {temperature}°C")
    except Exception as e:
        print(f"Error updating data: {e}")

def main():
    print("Starting the gym capacity and weather tracker...")
    schedule.every().hour.do(update_excel)
    print("Scheduled to update every hour")
    
    update_excel()  # Run once immediately
    
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    main()