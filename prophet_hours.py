import pandas as pd
from prophet import Prophet

# Load the wind speed data into a Pandas DataFrame
df = pd.read_csv('datasets/1wind/St_5_WindSpeed.csv')

# Create a Prophet model
m = Prophet()

# Fit the Prophet model to the data in the DataFrame
m.fit(df)

# Create a future dataframe for the next 3 hours
future = m.make_future_dataframe(periods=3, freq='H')

# Predict the wind speed for the next 3 hours using the Prophet model
forecast = m.predict(future)

# Print the predicted wind speed values to the console
print(forecast[['ds', 'yhat']])