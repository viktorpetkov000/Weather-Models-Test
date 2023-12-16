import pandas as pd

def convert_datetime(input_file, output_file):
    df = pd.read_csv(input_file)
    df['timestamp'] = pd.to_datetime(df['timestamp'], format="%y-%m-%d %H:%M:%S")
    df['timestamp'] = df['timestamp'] - pd.to_timedelta(df['timestamp'].dt.minute % 10, unit='m')
    df.to_csv(output_file, index=False)

input_file_path = 'datasets/MS5/MS5_wind_dateTime_year.csv'
output_file_path = 'datasets/MS5/MS5_wind_dateTime_year_converted_rounded.csv'

convert_datetime(input_file_path, output_file_path)