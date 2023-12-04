#%%
import pandas as pd
df = pd.read_csv('datasets/1wind/st_5_windspeed_converted.csv')
df['timestamp'] = pd.to_datetime(df['timestamp'])

for name, group in df.groupby(pd.Grouper(key='timestamp', freq='1W')):
    file_name = f'datasets/1wind/1week/data_{str(name).replace(" ", "_").replace(":", "_")}.csv'
    group.to_csv(file_name, index=False)
    
for name, group in df.groupby(pd.Grouper(key='timestamp', freq='2W')):
    file_name = f'datasets/1wind/2weeks/data_{str(name).replace(" ", "_").replace(":", "_")}.csv'
    group.to_csv(file_name, index=False)
    
for name, group in df.groupby(pd.Grouper(key='timestamp', freq='4W')):
    file_name = f'datasets/1wind/4weeks/data_{str(name).replace(" ", "_").replace(":", "_")}.csv'
    group.to_csv(file_name, index=False)
    
for name, group in df.groupby(pd.Grouper(key='timestamp', freq='8W')):
    file_name = f'datasets/1wind/8weeks/data_{str(name).replace(" ", "_").replace(":", "_")}.csv'
    group.to_csv(file_name, index=False)
    
for name, group in df.groupby(pd.Grouper(key='timestamp', freq='16W')):
    file_name = f'datasets/1wind/16weeks/data_{str(name).replace(" ", "_").replace(":", "_")}.csv'
    group.to_csv(file_name, index=False)
    
for name, group in df.groupby(pd.Grouper(key='timestamp', freq='6M')):
    file_name = f'datasets/1wind/6months/data_{str(name).replace(" ", "_").replace(":", "_")}.csv'
    group.to_csv(file_name, index=False)

# %%
