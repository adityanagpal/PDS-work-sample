#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Oct 23 11:48:22 2018

@author: aditya
"""
########## IMPORTING THE LIBRARIES ##############
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

############### READING THE DATASET ###################

dataset = pd.ExcelFile('BOMBAY_AIR90JAN16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df1 = dataset.parse('Results')

dataset = pd.ExcelFile('BOMBAY_AIR90-JUL17IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df2 = dataset.parse('Report')

dataset = pd.ExcelFile('CALCUTTA_AIR90NOV16-DEC16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df3 = dataset.parse('Report')

dataset = pd.ExcelFile('Delhi_air_90_-JAN16IMP.xls')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df4 = dataset.parse('Results')

dataset = pd.ExcelFile('HYDERABAD_AIR90-DEC16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df5 = dataset.parse('Report')

dataset = pd.ExcelFile('HYDERABAD_AIR90-NOV16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df6 = dataset.parse('Report')

dataset = pd.ExcelFile('MADRAS_AIR90-DEC16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df7 = dataset.parse('Report')

dataset = pd.ExcelFile('Report_BangaloreAir_Hs90_Imp_Jan16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df8 = dataset.parse('Report')

dataset = pd.ExcelFile('Report_DelhiAir_Hs90_Imp_Jan16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df9 = dataset.parse('Report')

dataset = pd.ExcelFile('4_Report_HyderabadAir_Hs90_Imp_Oct16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df10 = dataset.parse('Report')

dataset = pd.ExcelFile('Report_MadrasAir_Hs90_Imp_Jan16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df11 = dataset.parse('Report')

############ MODYFYING COLUMN NAMES ###########################

df1 = df1.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin'})

df4 = df4.rename(columns={'Actual Duty (Total BE Wise)':'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin'})

########################### MERGING THE DATA FRAMES ############################


results = pd.concat([df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11])

############ EXTRACTING THE IMPORTANT COLUMS - Product Description and Unit Rate ####

X=results["Unit Rate INR"].tolist()
y1=results["Product Description"].tolist()

############ EXTRACTING THE IMPORTANT COLUMS - Product Description and Unit Rate ####

X1 = results["A. Value INR"].tolist()
X2 = results["Actual Duty Paid INR"].tolist()

#################### INITIALISING THE COLUMNS ###################

j = len(X)

results['Deviations_Unit_Rate']=[0 for i in range(j)]
results['Duty Percentage']=[0 for i in range(j)]
results['Duty_Percent_Deviation']=[0 for i in range(j)]
results['Mean_Unit_Rate']=[0 for i in range(j)]
results['Mean_Duty_Percent'] = [0 for i in range(j)]

############## DUTY PERCENTAGE CALCULATION #################
X2 = np.array(X2); X1 = np.array(X1)
X2[X2 == "No Duty Paid"] = 0 # making such entries zero
X2 = X2.astype("float32") # from strings to floats
X1[X1 == "No Duty Paid"] = 0 # making such entries zero
X1 = X1.astype("float32") # from strings to floats
cl = len(results.columns)-4
results.iloc[:, cl] = (X2/X1) * 100
        
########################### CALCULATING THE ABSOLUTE VALUE OF PRODUCTS #########

unit_rate_mean = results.groupby(['Product Description'])['Unit Rate INR'].mean()

duty_percent_mean = results.groupby(['Product Description'])['Duty Percentage'].mean()  

######################## DETECTING THE DEVIATIONS ##################
dul = len(results.columns)-5
dpl = len(results.columns)-3
dpmv = len(results.columns)-1
urmv = len(results.columns)-2

results = results.fillna(0)
results=results.reset_index()
del results['index']

X = np.array(X)
X = X.astype("float32") # from strings to floats

# for unit rate:
means = np.array(unit_rate_mean[y1[:]])
diff = abs(X - means)
fraction = (means / 20 ) 
mask = diff > fraction
k = np.array(["Deviation_Unit_Rate" if j else "" for j in mask]) # better this way
results.iloc[:,dul] = k
results.iloc[:,urmv] = means # write the means

# for duty percentage
means_duty = np.array(duty_percent_mean[y1[:]])
diff_duty = abs(results.iloc[:, cl] - means_duty)
fraction_duty = (means_duty / 20 ) 
mask_duty = diff_duty > fraction_duty
k = np.array(["Deviation_Duty_Percent" if j else "" for j in mask_duty]) # better this way
results.iloc[:,dpl] = k
results.iloc[:,dpmv] = means_duty # write the means
        
#################### REARRANGING THE COLUMNS #######################
cols= results.columns.tolist()
print(cols)
results = results[['Date','HS Code','Product Description','Importer Name','Duty Percentage',
           'Deviations_Unit_Rate', 'Duty_Percent_Deviation', 'Mean_Unit_Rate', 'Mean_Duty_Percent',
           'Unit Rate INR','A. Value INR','Actual Duty Paid INR','BE Number','IEC',
           'Applicable Duty INR','Port Of Destination','Country Of Origin',
           'CHA Name', 'CHA Pan No', 'City',
           'Foreign Country', 'HS Description','Inv No', 'Port Of Origin',
           'QTY', 'State', 'UNIT',]]
 

##################### REARRANGING INDEX GROUP BY PRODUCT DESCRIPTION ########

results=results.reset_index()
df = results.groupby('Product Description')['index'].apply(list)
List = []
del results['index']
for i in range(len(df)):
    List+=df[i]
    
results = results.reindex(List)    
 
############## CONVERTING INTO EXCEL #############################


results.to_csv('output90(1).csv')




     
  
