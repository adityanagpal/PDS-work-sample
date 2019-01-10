#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct 24 22:56:50 2018

@author: aditya
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

############### READING THE DATASET ###################

dataset = pd.ExcelFile('1_Report_BangaloreAir_Hs85_Imp_Jun16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df1 = dataset.parse('Report')

dataset = pd.ExcelFile('1_Report_BombayAir_Hs85_Imp_Jun16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df2 = dataset.parse('Report')

dataset = pd.ExcelFile('1_Report_DelhiAir_Hs85_Imp_Jul16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df3 = dataset.parse('Report')

dataset = pd.ExcelFile('1_Report_DelhiAir_Hs85_Imp_Jun16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df4 = dataset.parse('Report')

dataset = pd.ExcelFile('2_Report_BangaloreAir_Hs85_Imp_Jun16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df5 = dataset.parse('Report')

dataset = pd.ExcelFile('2_Report_BombayAir_Hs85_Imp_Jun16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df6 = dataset.parse('Report')

dataset = pd.ExcelFile('2_Report_DelhiAir_Hs85_Imp_Jul16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df7 = dataset.parse('Report')

dataset = pd.ExcelFile('2_Report_DelhiAir_Hs85_Imp_Jun16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df8 = dataset.parse('Report')

dataset = pd.ExcelFile('3_Report_DelhiAir_Hs85_Imp_Jul16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df9 = dataset.parse('Report')

dataset = pd.ExcelFile('3_Report_DelhiAir_Hs85_Imp_Jun16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df10 = dataset.parse('Report')

dataset = pd.ExcelFile('3_Report_MadrasAir_Hs85_Imp_Jun16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df11 = dataset.parse('Report')

dataset = pd.ExcelFile('Bombay_air_85_FEB16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df12 = dataset.parse('Results')

dataset = pd.ExcelFile('Bombay_air_85_JUL14IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df13 = dataset.parse('Results')

dataset = pd.ExcelFile('Bombay_air_85_JUL15IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df14 = dataset.parse('Results')

dataset = pd.ExcelFile('Bombay_air_85_OCT14IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df15 = dataset.parse('Results')

dataset = pd.ExcelFile('Bombay_air_85_SEP14IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df16 = dataset.parse('Results')

dataset = pd.ExcelFile('Bombay_air_85_SEP15-OCT15IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df17 = dataset.parse('Results')

dataset = pd.ExcelFile('BOMBAY_AIR30JAN16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df18 = dataset.parse('Results')

dataset = pd.ExcelFile('BOMBAY_AIR85OCT15IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df19 = dataset.parse('Results')

dataset = pd.ExcelFile('BOMBAY_AIR85SEP15IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df20 = dataset.parse('Results')

dataset = pd.ExcelFile('CALCUTTA_SEA85MAR15IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df21 = dataset.parse('Results')

dataset = pd.ExcelFile('Delhi_air_90_-JAN16IMP.xls')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df22 = dataset.parse('Results')

dataset = pd.ExcelFile('DELHI_AIR85JUL14IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df23 = dataset.parse('Results')

dataset = pd.ExcelFile('DELHI_AIR85MAY14IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df24 = dataset.parse('Results')

dataset = pd.ExcelFile('JNPT98JAN16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df25 = dataset.parse('Results')

dataset = pd.ExcelFile('MADRAS_AIR85JAN16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df26 = dataset.parse('Results')

dataset = pd.ExcelFile('MADRAS_SEA85OCT14IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df27 = dataset.parse('Results')

dataset = pd.ExcelFile('TKD85JAN16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df28 = dataset.parse('Results')


############ MODYFYING COLUMN NAMES ###########################

df12 = df12.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})


df13 = df13.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df14 = df14.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df15 = df15.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df16 = df16.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df17 = df17.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df18 = df18.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df19 = df19.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df20 = df20.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df21 = df21.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df22 = df22.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df23 = df23.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df24 = df24.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df25 = df25.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df26 = df27.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df27 = df27.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})

df28 = df28.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate'})



########################### MERGING THE DATA FRAMES ############################


results = pd.concat([df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13,
                     df14,df15,df16,df17,df18,df19,df20,df21,df22,df23,df24,
                     df25,df26,df27,df28])

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

###################### NULL VALUE REMOVING ###############

results = results.fillna(0)
results=results.reset_index()
del results['index']

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
           'Applicable Duty INR','Port Of Destination','Country Of Origin','Address1',
           'CHA Name', 'CHA No', 'City', 'Contact Person','E-Mail',
           'Foreign Country', 'Phone', 'Pin', 'Port Of Origin','QTY', 'State', 'UNIT',
           'Unit Rate USD','FAX','Shipment Mode','Standard Qty','Standard Unit',
           'Standard Unit Rate', 'Unit Rate USD',
          'Value + Duty in INR', 'Value + Duty in USD']]

##################### REARRANGING INDEX GROUP BY PRODUCT DESCRIPTION ########

results=results.reset_index()
df = results.groupby('Product Description')['index'].apply(list)
List = []
del results['index']
for i in range(len(df)):
    List+=df[i]
    
results = results.reindex(List) 

for i in range(j):
    if results.iloc[i,2]==0:
        continue
    else:
        cv=i
        break


for i in range(cv):
    del results.iloc[i,:]   


 
############## CONVERTING INTO EXCEL #############################


results.to_csv('output85(1).csv')

