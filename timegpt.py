#%% Imports and setup
from nixtlats import TimeGPT
import pandas as pd
token = "CLfEQIBnngCcRuxmVz8wCe9EbQY87J80en7pLo3T11UcP3ra3vSHEtEb1P61nI9iJTL1klVM7sXxy0RxXHJVtTmxfqJKVcqWYxfgbPDIENq41JdTpSjwW268a4j5sE4QE6RvaxhMYSYdqtxsdP8V0a5EKJ79Nqs2Ficn67ov5eB27OzJuUL62RlKSHzxEADIpC9GixiRdc2uaHnltO4uQp8BU89q9D8ACm7rYTJu0ntHcTtLK3zpUUh97lEuCsd7"
file_name_plot = "plots/timegpt_result.png"
file_name_workbook = "workbooks/results.xlsx"
file_name_dataset = "datasets/1wind/st_5_windspeed_converted.csv"
time_col = "timestamp"
target_col = "wind_speed"
freq = "10T"
timegpt = TimeGPT(
    token = token
)
df = pd.read_csv(file_name_dataset)

#%% Visualize original data
print(df.head())
plot = timegpt.plot(df, time_col=time_col, target_col=target_col)
plot.savefig("plots/original_dataset")
plot

#%% Forecast the next 3 hours with 10 minutes intervals (18 steps)
timegpt_fcst_df = timegpt.forecast(df=df, h=18, time_col=time_col, target_col=target_col, freq=freq)

# %% Visualize TimeGPT predictions
print(timegpt_fcst_df.head())
timegpt_plot = timegpt.plot(df, timegpt_fcst_df, time_col=time_col, target_col=target_col)
timegpt_plot.savefig(file_name_plot)
timegpt_plot

#%% Add data and charts to excel
writer = pd.ExcelWriter(file_name_workbook, engine='xlsxwriter')
workbook = writer.book

#%% Add the original data
df.to_excel(writer, sheet_name='Sheet1')
worksheet = writer.sheets['Sheet1']
chart = workbook.add_chart({'type': 'line'})
chart.add_series({
    'name':       ['Sheet1', 0, 2],
    'categories': ['Sheet1', 1, 1, len(df), 1],
    'values':     ['Sheet1', 1, 2, len(df), 2],
})
chart.set_x_axis({'name': time_col, 'position_axis': 'on_tick'})
chart.set_y_axis({'name': target_col, 'major_gridlines': {'visible': False}})
chart.set_legend({'position': 'top'})
worksheet.insert_chart('D2', chart)

#%% Add the TimeGPT Predictions
timegpt_fcst_df.to_excel(writer, sheet_name='Sheet2')
worksheet = writer.sheets['Sheet2']
chart = workbook.add_chart({'type': 'line'})
chart.add_series({
    'name':       ['Sheet2', 0, 2],
    'categories': ['Sheet2', 1, 1, len(timegpt_fcst_df), 1],
    'values':     ['Sheet2', 1, 2, len(timegpt_fcst_df), 2],
})
chart.set_x_axis({'name': time_col, 'position_axis': 'on_tick'})
chart.set_y_axis({'name': target_col, 'major_gridlines': {'visible': False}})
chart.set_legend({'position': 'top'})
worksheet.insert_chart('D2', chart)

#%% Save the excel
writer.close()
# %%
