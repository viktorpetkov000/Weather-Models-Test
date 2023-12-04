#%% Imports and setup
from nixtlats import TimeGPT
from prophet import Prophet
import pandas as pd
token = "CLfEQIBnngCcRuxmVz8wCe9EbQY87J80en7pLo3T11UcP3ra3vSHEtEb1P61nI9iJTL1klVM7sXxy0RxXHJVtTmxfqJKVcqWYxfgbPDIENq41JdTpSjwW268a4j5sE4QE6RvaxhMYSYdqtxsdP8V0a5EKJ79Nqs2Ficn67ov5eB27OzJuUL62RlKSHzxEADIpC9GixiRdc2uaHnltO4uQp8BU89q9D8ACm7rYTJu0ntHcTtLK3zpUUh97lEuCsd7"
unique_run = "1week"
file_name = "data_2022-10-02_00_00_00"
file_name_dataset = f"datasets/1wind/{unique_run}.csv"
file_name_plot = f"plots/timegpt_result_{unique_run}.png"
file_name_plot_original = f"plots/original_dataset_{unique_run}.png"
file_name_workbook = f"workbooks/results_{unique_run}.xlsx"

time_col = "timestamp"
target_col = "wind_speed"
freq = "10T"
steps = 18
timegpt = TimeGPT(
    token = token
)
df = pd.read_csv(file_name_dataset)

#%% Prophet initialize
file_name_plot_prophet = f"plots/prophet_result_{unique_run}.png"
file_name_dataset_prophet = f"datasets/1wind/{unique_run}_prophet.csv"

#%% Visualize original data
print(df.head())
plot = timegpt.plot(df, time_col=time_col, target_col=target_col)
plot.savefig(file_name_plot_original)
plot

#%% Forecast the next 3 hours with 10 minutes intervals (18 steps) using TimeGPT
timegpt_fcst_df = timegpt.forecast(df=df, h=steps, time_col=time_col, target_col=target_col, freq=freq)

# %% Visualize TimeGPT predictions
print(timegpt_fcst_df.tail(steps))
timegpt_plot = timegpt.plot(df, timegpt_fcst_df, time_col=time_col, target_col=target_col)
timegpt_plot.savefig(file_name_plot)
timegpt_plot

#%% Forecast the next 3 hours with 10 minutes intervals (18 steps) using Prophet
m = Prophet()
df_prophet = pd.read_csv(file_name_dataset_prophet)
m.fit(df_prophet)
future = m.make_future_dataframe(periods=steps, freq=freq)
forecast = m.predict(future)

#%% Visualize Prophet predictions
print(forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].tail(steps))
prophet_plot = m.plot_components(forecast)
prophet_plot.savefig(file_name_plot_prophet)

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

#%% Add the Prophet Predictions
forecast.tail(steps).to_excel(writer, sheet_name='Sheet3')
worksheet = writer.sheets['Sheet3']
chart = workbook.add_chart({'type': 'line'})
chart.add_series({
    'name':       ['Sheet3', 0, 19],
    'categories': ['Sheet3', 0, 1, len(forecast.tail(steps)), 1],
    'values':     ['Sheet3', 0, 19, len(forecast.tail(steps)), 19],
})
chart.set_x_axis({'name': time_col, 'position_axis': 'on_tick'})
chart.set_y_axis({'name': target_col, 'major_gridlines': {'visible': False}})
chart.set_legend({'position': 'top'})
worksheet.insert_chart('D2', chart)

#%% Comparison sheet

# Add the original data and TimeGPT predictions to a new sheet
writer = pd.ExcelWriter(file_name_workbook, engine='xlsxwriter')
workbook = writer.book
df_combined = pd.merge(df, timegpt_fcst_df, on=time_col, how='outer')
print(df_combined)
df_combined.to_excel(writer, sheet_name='Sheet4')
worksheet = writer.sheets['Sheet4']
writer.close()

#%% Save the excel
writer.close()
# %%
