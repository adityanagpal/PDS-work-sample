#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct 24 21:03:58 2018

@author: aditya
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

############### READING THE DATASET ###################

dataset = pd.ExcelFile('BOMBAY_AIR30JAN16IMP.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df1 = dataset.parse('Results')

dataset = pd.ExcelFile('Report_DelhiAir_Hs30_Imp_Jan16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df2 = dataset.parse('Report')

dataset = pd.ExcelFile('Report_HyderabadAir_Hs30_Imp_Jan16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df3 = dataset.parse('Report')

dataset = pd.ExcelFile('Report_MadrasAir_Hs30_Imp_Jan16.xlsx')
print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
df4 = dataset.parse('Report')

############ MODYFYING COLUMN NAMES ###########################

df1 = df1.rename(columns={'Actual Duty (Total BE Wise)': 'Actual Duty Paid INR',
                          'Applicable Duty in INR':'Applicable Duty INR',
                          'Ass. Value in INR':'A. Value INR',
                          'Unit Rate in INR':'Unit Rate INR',
                          'Unit Rate in USD':'Unit Rate USD',
                          'Indian Importer':'Importer Name',
                          'Indian Port':'Port Of Destination',
                          'Foreign Port':'Port Of Origin',
                          'Standard Unit Rate(INR)':'Standard Unit Rate INR'})


########################### MERGING THE DATA FRAMES ############################


results = pd.concat([df1,df2,df3,df4])

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
           'A. Value $','CHA Name', 'CHA Pan No', 'City', 'Contact Person','E-Mail',
           'Foreign Country', 'Phone', 'Pin', 'Port Of Origin',
           'QTY', 'State', 'UNIT', 'Unit Rate USD', 'A.Value + Duty $', 'A.Value + Duty INR',
             'Address2', 'Ass. Value in USD','FAX','Shipment Mode', 'Standard Qty',
           'Standard Unit', 'Standard Unit Rate INR', 
           'State', 'Unit Rate USD', 'Value + Duty in INR', 'Value + Duty in USD']]

##################### REARRANGING INDEX GROUP BY PRODUCT DESCRIPTION ########

results=results.reset_index()
df = results.groupby('Product Description')['index'].apply(list)
List = []
del results['index']
for i in range(len(df)):
    List+=df[i]
    
results = results.reindex(List)    
 
############## CONVERTING INTO EXCEL #############################


results.to_csv('output30(1).csv')


