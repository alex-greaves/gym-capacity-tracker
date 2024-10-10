import requests
import re
import pandas as pd
from datetime import datetime, time, date
import time as time_module
import schedule
import os
import configparser

# Read config file
config = configparser.ConfigParser()
config.read('config.ini')

# OpenWeatherMap API key
WEATHER_API_KEY = config['API_KEYS']['OPENWEATHERMAP']

# Define gym configurations
GYMS = [
    {
        'name': 'Ground Up Squamish',
        'url': 'https://portal.rockgympro.com/portal/public/a9c7d8e770832d2ce2a1a5b371f95dfb/occupancy?&iframeid=occupancyCounter&fId=',
        'lat': '49.7240836',
        'lon': '-123.1526125',
        'parser': 'groundup',
        'type': 'rockgympro',
        # 'open_time': time(6, 0),  # 6:00 AM PST
        'close_time': time(22, 0)  # 10:00 PM PST
    },
]

def get_capacity_rockgympro(url):
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
    return None, None

def get_weather(lat, lon):
    url = f"https://api.openweathermap.org/data/2.5/weather?lat={lat}&lon={lon}&appid={WEATHER_API_KEY}&units=metric"
    try:
        print(f"Fetching weather data from: {url}")
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        print(f"Weather API response: {data}")

        weather_main = data['weather'][0]['main']
        temperature = data['main']['temp']
        print(f"Extracted weather: {weather_main}, temperature: {temperature}°C")
        return weather_main, temperature
    except Exception as e:
        print(f"Error getting weather data: {e}")
        return None, None

def is_gym_open(gym):
    if 'open_time' not in gym or 'close_time' not in gym:
        return True  # Always consider the gym open if no times are specified

    now = datetime.now().time()
    if gym['open_time'] < gym['close_time']:
        return gym['open_time'] <= now < gym['close_time']
    else:  # Handles case where gym is open past midnight
        return now >= gym['open_time'] or now < gym['close_time']

def update_excel():
    print("Updating Excel file...")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    day = date.strftime(date.today(), "%A")

    try:
        df = pd.read_excel('multi_gym_capacity_weather.xlsx')
        print("Existing Excel file found")
    except FileNotFoundError:
        print("Creating new Excel file")
        df = pd.DataFrame(columns=['Timestamp', 'Day', 'Gym', 'Count', 'Capacity', 'Weather', 'Temperature'])

    for gym in GYMS:
        if not is_gym_open(gym):
            print(f"{gym['name']} is currently closed. Skipping update.")
            continue

        print(f"Processing gym: {gym['name']}")

        # Get capacity data
        if gym['type'] == 'rockgympro':
            count, capacity = get_capacity_rockgympro(gym['url'])
        # elif gym['parser'] == 'another_parser':
        #     count, capacity = get_capacity_another_parser(gym['url'])
        else:
            print(f"Unknown parser for gym: {gym['name']}")
            continue

        # Get weather data
        weather_main, temperature = get_weather(gym['lat'], gym['lon'])
        
        new_row = pd.DataFrame({
            'Timestamp': [timestamp],
            'Day': [day],
            'Gym': [gym['name']],
            'Count': [count],
            'Capacity': [capacity],
            'Weather': [weather_main],
            'Temperature': [temperature]
        })
        df = pd.concat([df, new_row], ignore_index=True)

        print(f"Updated at {timestamp} for {gym['name']}: Count = {count}, Capacity = {capacity}, Weather = {weather_main}, Temp = {temperature}°C")

    df.to_excel('multi_gym_capacity_weather.xlsx', index=False)
    print("Excel file updated successfully")

def main():
    print("Starting the multi-gym capacity and weather tracker...")
    schedule.every(30).minutes.do(update_excel)
    print("Scheduled to update every 30 minutes")

    update_excel()  # Run once immediately

    while True:
        schedule.run_pending()
        time_module.sleep(1)

if __name__ == "__main__":
    main()
