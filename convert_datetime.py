import pandas as pd

def convert_datetime(input_file, output_file):
    df = pd.read_csv(input_file)
    df['timestamp'] = pd.to_datetime(df['timestamp'], format="%y-%m-%d %H:%M:%S")
    df.to_csv(output_file, index=False)

input_file_path = 'datasets/1wind/St_5_WindSpeed.csv'
output_file_path = 'datasets/1wind/st_5_windspeed_converted.csv'

convert_datetime(input_file_path, output_file_path)