#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from matplotlib import pyplot
import numpy as np
from statsmodels.tsa.arima_model import ARIMA
import matplotlib.pylab as plt
from sklearn.metrics import mean_squared_error


#Reading the input excel file for Crude Oil Prices
data = pd.read_excel('File.xls', sheetname="Data 1", skiprows=2)

#Setting index as Date
data.set_index(['Date'], inplace=True)

#Extracting the WTI prices as list
ts = data['Cushing, OK WTI Spot Price FOB (Dollars per Barrel)'] 

#Checking if the series data is stationary and create visual plot
def test_stationarity(series):
    
    #Determing rolling stats - Mean and Standard deviation
    rolmean = pd.rolling_mean(series, window=12)
    rolstd = pd.rolling_std(series, window=12)

    ax = plt.subplot(111)
    
    #Plot rolling stats here 
    ax.plot(series, color='blue',label='Original Values')
    ax.plot(rolmean, color='red', label='Rolling Mean')
    ax.plot(rolstd, color='green', label = 'Rolling Standard Deviation')
    ax.legend(loc='best')
    plt.title('Rolling Mean & Standard Deviation')
    plt.show()
    
#Checking stationarity of given timeseries data
plt.plot(ts)
test_stationarity(ts)

#Checking stationarity of log transformation of timeseries data    
ts_log = np.log(ts)
test_stationarity(ts_log)

#Plotting the moving average for window size 12
moving_avg = pd.rolling_mean(ts_log,12)
plt.plot(ts_log)
plt.plot(moving_avg, color='red')

#Checking stationarity of log transformation difference of timeseries data  
ts_log_diff = ts_log - ts_log.shift()
plt.plot(ts_log_diff)
test_stationarity(ts_log_diff)


#Building Model and Comparison for better performance
model = ARIMA(np.asarray(ts_log), order=(5,1,4))
model_fit = model.fit(disp=0)
print(model_fit.summary())

#Printing the residuals
residuals = pd.DataFrame(model_fit.resid)
residuals.plot()
pyplot.show()
residuals.plot(kind='kde')
pyplot.show()
print(residuals.describe())


#Predcition based on the input values
X = ts.values
size = int(len(X) * 0.66)
train, test = X[0:size], X[size:len(X)]
history = [x for x in train]
predictions = list()

for t in range(len(test)):
    model = ARIMA(history, order=(5,1,0))
    model_fit = model.fit(disp=0)
    model = model_fit
    output = model_fit.forecast()
    yhat = output[0]
    predictions.append(yhat)
    if(t >= len(test)):
        obs = yhat # used for future prediction
    else:
        obs = test[t]
    
    history.append(obs)
    print('predicted=%f, Actual=%f' % (yhat, obs))

#Calculating the Mean Squared Error as Performance metric
error = mean_squared_error(test, predictions)
print('Test MSE: %.3f' % error)

# plot the final results of ARIMA model prediction
pyplot.plot(test)
pyplot.plot(predictions, color='red')
pyplot.show()
