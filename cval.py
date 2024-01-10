#%% Imports and setup
from nixtlats import TimeGPT
from prophet import Prophet
import pandas as pd
import numpy as np
from scipy.stats import pearsonr
token = ""
unique_run = "finetune_MS5_run1"
file_name_dataset = f"datasets/MS5/MS5_wind_dateTime_year_converted_rounded.csv"
next_3_hours_data = f"datasets/MS5/MS5_wind_dateTime_year_next3hours_rounded.csv"
file_name_plot = f"datasets/MS5/results/plots/timegpt_result_{unique_run}.png"
file_name_plot_original = f"datasets/MS5/results/plots/original_dataset_{unique_run}.png"
file_name_workbook = f"workbooks/results_{unique_run}.xlsx"
time_col = "timestamp"
target_col = "wind_speed"
freq = "10T"
steps = 18
finetune_steps = [1,100,150,200,250,300]
#%%
timegpt_fcst_df = {}
timegpt_rmse = {}
timegpt_abse = {}
timegpt_rele = {}
timegpt_corr = {}
#%%
timegpt = TimeGPT(
    token = token
)
df = pd.read_csv(file_name_dataset)
#%% Visualize original data
print(df)
plot = timegpt.plot(df, time_col=time_col, target_col=target_col)
plot.savefig(file_name_plot_original)
plot

#%% Forecast the next 3 hours with 10 minutes intervals (18 steps) using TimeGPT, different forcast times
for i in range(len(finetune_steps)):
  timegpt_fcst_df[i] = timegpt.cross_validation(df=df, h=steps, n_windows=5, finetune_steps=finetune_steps[i], time_col=time_col, target_col=target_col, freq=freq)
  timegpt_fcst_df[i] = timegpt_fcst_df[i].rename(columns = {'TimeGPT':f'TimeGPT{finetune_steps[i]}'})

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
for i in range(len(finetune_steps)):
  sheet_name = f'Finetune {finetune_steps[i]}'
  timegpt_fcst_df[i].to_excel(writer, sheet_name=sheet_name)
  worksheet = writer.sheets[sheet_name]
  chart = workbook.add_chart({'type': 'line'})
  chart.add_series({
      'name':       [sheet_name, 0, 2],
      'categories': [sheet_name, 1, 1, len(timegpt_fcst_df[i]), 1],
      'values':     [sheet_name, 1, 2, len(timegpt_fcst_df[i]), 2],
  })
  chart.set_x_axis({'name': time_col, 'position_axis': 'on_tick'})
  chart.set_y_axis({'name': target_col, 'major_gridlines': {'visible': False}})
  chart.set_legend({'position': 'top'})
  worksheet.insert_chart('D2', chart)

#%% Comparison sheet
next_3_hours = pd.read_csv(next_3_hours_data)
next_3_hours=next_3_hours.rename(columns = {'wind_speed':'Original'})
df_combined = pd.merge(timegpt_fcst_df[0], next_3_hours, on='timestamp', how='outer')
for i in range(len(finetune_steps)-1):
    df_combined = pd.merge(timegpt_fcst_df[i+1], df_combined, on='timestamp', how='outer')
df_combined['timestamp'] = pd.to_datetime(df_combined['timestamp'])
df_combined['timestamp'] = df_combined['timestamp'].dt.strftime("%Y-%m-%d %H:%M:%S")
df_combined.to_excel(writer, sheet_name='Sheet4')
worksheet = writer.sheets['Sheet4']
chart = workbook.add_chart({'type': 'line'})
# Add the TimeGPT predictions series
for i in range(len(finetune_steps)):
  chart.add_series({
      'name':       ['Sheet4', 0, i+2],
      'categories': ['Sheet4', 1, 1, len(df_combined), 1],
      'values':     ['Sheet4', 1, i+2, len(df_combined), i+2],
  })

# Add the next week series
chart.add_series({
    'name':       ['Sheet4', 0, len(finetune_steps)+2],
    'categories': ['Sheet4', 1, 1, len(df_combined), 1],
    'values':     ['Sheet4', 1, len(finetune_steps)+2, len(df_combined), len(finetune_steps)+2],
})

chart.set_x_axis({'name': time_col, 'position_axis': 'on_tick'})
chart.set_y_axis({'name': target_col, 'major_gridlines': {'visible': False}})
chart.set_legend({'position': 'top'})
worksheet.insert_chart('H2', chart)

#%%
# Add the root mean squared
worksheet = writer.sheets['Sheet4']
worksheet.write(f'B{steps+3}', 'Model')
worksheet.write(f'C{steps+3}', 'RMSE')
worksheet.write(f'D{steps+3}', 'Absolute Error')
worksheet.write(f'E{steps+3}', 'Relative Error')
worksheet.write(f'F{steps+3}', 'Correlation')
for i in range(len(finetune_steps)):
    timegpt_rmse[i] = np.sqrt(np.mean((df_combined['Original'] - df_combined[f'TimeGPT{finetune_steps[i]}'])**2))
    worksheet.write(f'A{i+steps+4}', i)
    worksheet.write(f'B{i+steps+4}', f'TimeGPT{finetune_steps[i]}')
    worksheet.write(f'C{i+steps+4}', timegpt_rmse[i])
    timegpt_abse[i] = abs(df_combined['Original'] - df_combined[f'TimeGPT{finetune_steps[i]}']).mean()
    worksheet.write(f'D{i+steps+4}', timegpt_abse[i])
    timegpt_rele[i] = (abs((df_combined['Original'] - df_combined[f'TimeGPT{finetune_steps[i]}']) / df_combined['Original']) * 100).mean()
    worksheet.write(f'E{i+steps+4}', timegpt_rele[i])
    timegpt_corr[i] = pearsonr(df_combined['Original'], df_combined[f'TimeGPT{finetune_steps[i]}'])[0]
    worksheet.write(f'F{i+steps+4}', timegpt_corr[i])
    #%% Save the excel
writer.close()

# %%
