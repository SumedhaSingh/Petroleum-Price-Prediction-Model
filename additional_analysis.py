#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@author: SumedhaSingh
"""

import pandas as pd
import numpy as np

allData = []

#Aggregating the information into dataframes from all the sheets
for i in range(7):
    data = pd.read_excel('Jr._Data_Analyst_Project_File.xls', sheetname="Data "+str(i+1), skiprows=2)
    data.set_index(['Date'], inplace=True)
    allData.append(data)

# Create a Excel Writer Object to Write Yearly analysis Price Change
writer = pd.ExcelWriter('output_yearly.xlsx')
writer_n = pd.ExcelWriter('output_monthly.xlsx')

#Types of Data Available
types = ['Crude Oil', 'Conventional Gasoline', 'RBOB Regular Gasoline', 'Heating Oil', 'Diesel Fuel', 'Kerosene Type Jet Fuel', 'Propane']

#Resultant array to store yearly and monthly data
resultAll = []
resultAll_month = []

#This method calcualtes the Yearly Mean Price of Oil Products   
def findYearlyPriceChange(data, ind):
    #Extracting the Prices
    ts = data[data.columns[0]] 

    #List of unique years
    yearlist = set(data.index.year)                
    #create a map to store yearly means
    yearMap = dict.fromkeys(yearlist)
    
    for i in yearlist:
        #Yearly aggregated data
        currentList = ts[str(i)]
        #find the mean for yearly aggregated data
        yearMap[i] = np.mean(currentList)
    
    #Find separate list for year and means
    yearls = list(yearMap.keys())
    meanls = list(yearMap.values())
    
    result = pd.DataFrame(np.column_stack([yearls, meanls]), 
                               columns=['Year', 'Mean Price of '+str(types[ind])])
    resultAll.append(result)
          

#This method calcualtes the Monthly Mean Price of Oil Products    
def findMonthyPriceChange(data, ind):
    ts = data[data.columns[0]] 

    monthlist = set(data.index.month)                
    #create a map to store Monthly means
    monthMap = dict.fromkeys(monthlist)
    
    for i in monthlist:
        currentList = ts[(ts.index.month==i)] 
        monthMap[i] = np.mean(currentList)
     
    yearls = list(monthMap.keys())
    meanls = list(monthMap.values())
    
    result = pd.DataFrame(np.column_stack([yearls, meanls]), 
                               columns=['Month', 'Mean Price of '+str(types[ind])])
    resultAll_month.append(result)    
    
for i in range(len(allData)):    
    findYearlyPriceChange(allData[i], i)    
    findMonthyPriceChange(allData[i], i) 
    
#Storing resultant monthly analysis to excel
for i in range(len(resultAll_month)):
    resultAll_month[i].to_excel(writer_n,'Data'+str(i+1))
writer_n.save()

#Storing resultant yearly analysis to excel 
for i in range(len(resultAll)):
    resultAll[i].to_excel(writer,'Data'+str(i+1))
writer.save()
