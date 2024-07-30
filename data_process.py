import json
import pandas as pd

# Load the JSON data
file_path = './weather_data.json'
with open(file_path, 'r') as file:
    data = json.load(file)

# Extract the relevant data
result = data['result']['hourly']

# Initialize an empty dictionary to store the data
processed_data = {
    "datetime": [],
    "经度": [],
    "纬度": [],
    "降水量": [],
    "降水概率": [],
    "温度": [],
    "体感温度": [],
    "地表10米风速": [],
    "风向": [],
    "地表2米相对湿度": [],
    "云量": [],
    "天气现象": [],
    "地面气压": [],
    "地表水平能见度": [],
    "向下短波辐射通量": [],
    "空气质量": [],
    "PM2.5浓度": []
}

# Helper function to extract data based on datetime


# Helper function to extract data based on datetime
def extract_data_by_datetime(key, datetime_list, value_key='value'):
    data_dict = {entry['datetime']: entry[value_key] for entry in result[key]}
    return [data_dict.get(dt, None) for dt in datetime_list]

# Helper function to handle nested dictionaries properly


def extract_data_by_datetime_nested(data, datetime_list, nested_keys):
    data_dict = {}
    for entry in data:
        value = entry
        try:
            for nk in nested_keys:
                value = value[nk]
            data_dict[entry['datetime']] = value
        except (KeyError, TypeError):
            data_dict[entry['datetime']] = None
    return [data_dict.get(dt, None) for dt in datetime_list]


# Extract the datetime list
datetime_list = [entry['datetime'] for entry in result['temperature']]

latitude = data['location'][0]
longitude = data['location'][1]

# Populate the processed data dictionary
processed_data['datetime'] = datetime_list
processed_data['经度'] = [longitude] * len(datetime_list)
processed_data['纬度'] = [latitude] * len(datetime_list)
processed_data['降水量'] = extract_data_by_datetime(
    'precipitation', datetime_list)
processed_data['降水概率'] = extract_data_by_datetime(
    'precipitation', datetime_list, value_key='probability')
processed_data['温度'] = extract_data_by_datetime(
    'temperature', datetime_list)
processed_data['体感温度'] = extract_data_by_datetime(
    'apparent_temperature', datetime_list)
processed_data['地表10米风速'] = extract_data_by_datetime(
    'wind', datetime_list, value_key='speed')  # Assuming wind speed for now
processed_data['风向'] = extract_data_by_datetime(
    'wind', datetime_list, value_key='direction')  # Assuming wind speed for now
processed_data['地表2米相对湿度'] = extract_data_by_datetime(
    'humidity', datetime_list)
processed_data['云量'] = extract_data_by_datetime(
    'cloudrate', datetime_list)
processed_data['天气现象'] = extract_data_by_datetime(
    'skycon', datetime_list, value_key='value')  # Assuming skycon is a value list
processed_data['地面气压'] = extract_data_by_datetime(
    'pressure', datetime_list)
processed_data['地表水平能见度'] = extract_data_by_datetime(
    'visibility', datetime_list)
processed_data['向下短波辐射通量'] = extract_data_by_datetime('dswrf', datetime_list)
processed_data['空气质量'] = extract_data_by_datetime_nested(
    result['air_quality']['aqi'], datetime_list, nested_keys=['value', 'chn'])
processed_data['PM2.5浓度'] = extract_data_by_datetime_nested(
    result['air_quality']['pm25'], datetime_list, nested_keys=['value'])

# Create a DataFrame
df = pd.DataFrame(processed_data)

# Save the DataFrame to a CSV file
output_file_path = './weather_data.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Data has been saved to {output_file_path}")
