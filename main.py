#%% Imports and setup
from nixtlats import TimeGPT
from prophet import Prophet
import pandas as pd
token = ""
unique_run = "full_dataset"
file_name_dataset = f"datasets/1wind/full_dataset/st_5_windspeed_converted.csv"
next_3_hours_data = f"datasets/1wind/full_dataset/hours/st_5_windspeed_converted.csv"
file_name_plot = f"plots/timegpt_result_{unique_run}.png"
file_name_plot_original = f"plots/original_dataset_{unique_run}.png"
file_name_workbook = f"workbooks/results_{unique_run}.xlsx"
file_name_plot_prophet = f"plots/timegpt_result_{unique_run}_prophet.png"

time_col = "timestamp"
target_col = "wind_speed"
freq = "10T"
steps = 18
timegpt = TimeGPT(
    token = token
)
df = pd.read_csv(file_name_dataset)

#%% Visualize original data
print(df)
plot = timegpt.plot(df, time_col=time_col, target_col=target_col)
plot.savefig(file_name_plot_original)
plot

#%% Forecast the next 3 hours with 10 minutes intervals (18 steps) using TimeGPT
timegpt_fcst_df = timegpt.forecast(df=df, h=steps, time_col=time_col, target_col=target_col, freq=freq)

# %% Visualize TimeGPT predictions
print(timegpt_fcst_df)
timegpt_plot = timegpt.plot(df, timegpt_fcst_df, time_col=time_col, target_col=target_col)
timegpt_plot.savefig(file_name_plot)
timegpt_plot

#%% Forecast the next 3 hours with 10 minutes intervals (18 steps) using Prophet
m = Prophet()
df_prophet = df.rename(columns = {'timestamp':'ds', 'wind_speed':'y',})
m.fit(df_prophet)
future = m.make_future_dataframe(periods=steps, freq=freq)
forecast = m.predict(future)

#%% Visualize Prophet predictions
print(forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].tail(steps))
prophet_plot = m.plot_components(forecast)
prophet_plot.savefig(file_name_plot)
forecast=forecast.rename(columns = {'ds':'timestamp', 'yhat':'Prophet'})

#%% Add data and charts to excel
writer = pd.ExcelWriter(file_name_workbook, engine='xlsxwriter')
workbook = writer.book

#%% Add the original data
df.to_excel(writer, sheet_name='Sheet1')
worksheet = writer.sheets['Sheet1']
df=df.rename(columns = {'wind_speed':'Original'})
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

#%% Add the Prophet predictions
forecast['timestamp'] = pd.to_datetime(forecast['timestamp'])
forecast['timestamp'] = forecast['timestamp'].dt.strftime("%Y-%m-%d %H:%M:%S")
forecast.tail(steps).to_excel(writer, sheet_name='Sheet3')
worksheet = writer.sheets['Sheet3']
chart = workbook.add_chart({'type': 'line'})
chart.add_series({
    'name':       ['Sheet3', 0, forecast.columns.get_loc('Prophet')+1],
    'categories': ['Sheet3', 1, 1, len(forecast.tail(steps)), 1],
    'values':     ['Sheet3', 1, forecast.columns.get_loc('Prophet')+1, len(forecast.tail(steps)), forecast.columns.get_loc('Prophet')+1],
})
chart.set_x_axis({'name': time_col, 'position_axis': 'on_tick'})
chart.set_y_axis({'name': target_col, 'major_gridlines': {'visible': False}})
chart.set_legend({'position': 'top'})
worksheet.insert_chart('D2', chart)

#%% Comparison sheet
next_3_hours = pd.read_csv(next_3_hours_data)
next_3_hours=next_3_hours.rename(columns = {'wind_speed':'Original'})
df_combined = pd.merge(timegpt_fcst_df, next_3_hours, on='timestamp', how='outer')
df_combined['timestamp'] = pd.to_datetime(df_combined['timestamp'])
forecast['timestamp'] = pd.to_datetime(forecast['timestamp'])
df_combined = pd.merge(df_combined, forecast.tail(steps), on='timestamp', how='outer')
df_combined['timestamp'] = df_combined['timestamp'].dt.strftime("%Y-%m-%d %H:%M:%S")

df_combined.to_excel(writer, sheet_name='Sheet4')
worksheet = writer.sheets['Sheet4']
chart = workbook.add_chart({'type': 'line'})
# Add the Original + TimeGPT predictions series
chart.add_series({
    'name':       ['Sheet4', 0, 2],
    'categories': ['Sheet4', 1, 1, len(df_combined), 1],
    'values':     ['Sheet4', 1, 2, len(df_combined), 2],
})

# Add the next week series
chart.add_series({
    'name':       ['Sheet4', 0, 3],
    'categories': ['Sheet4', 1, 1, len(df_combined), 1],
    'values':     ['Sheet4', 1, 3, len(df_combined), 3],
})

# Add the next prophet series
chart.add_series({
    'name':       ['Sheet4', 0, df_combined.columns.get_loc('Prophet')+1],
    'categories': ['Sheet4', 1, 1, len(df_combined), 1],
    'values':     ['Sheet4', 1, df_combined.columns.get_loc('Prophet')+1, len(df_combined), df_combined.columns.get_loc('Prophet')+1],
})
chart.set_x_axis({'name': time_col, 'position_axis': 'on_tick'})
chart.set_y_axis({'name': target_col, 'major_gridlines': {'visible': False}})
chart.set_legend({'position': 'top'})
worksheet.insert_chart('H2', chart)

#%% Save the excel
writer.close()
# %%
