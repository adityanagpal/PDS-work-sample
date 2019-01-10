#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Dec 18 12:22:22 2018

@author: aditya
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pyxlsb import open_workbook as open_xlsb


############### READING THE DATASET ###################

df = []

with open_xlsb('85-OCT18IMP.xlsb') as wb:
    with wb.get_sheet(2) as sheet:
        for row in sheet.rows():
            df.append([item.v for item in row])

df = pd.DataFrame(df[1:], columns=df[0])

df1 = []

with open_xlsb('85-SEP18IMP.xlsb') as wb:
    with wb.get_sheet(2) as sheet:
        for row in sheet.rows():
            df1.append([item.v for item in row])

df1 = pd.DataFrame(df1[1:], columns=df1[0])

df2 = []

with open_xlsb('JNPT_85-SEP18IMP.xlsb') as wb:
    with wb.get_sheet(2) as sheet:
        for row in sheet.rows():
            df2.append([item.v for item in row])

df2 = pd.DataFrame(df2[1:], columns=df2[0])


############ MODYFYING COLUMN NAMES ###########################


df2 = df2.rename(columns={
                          'Importer Add1':'Address1',
                          'Importer Add2':'Address2',
                          'Importer City':'City',
                          'Importer PinCode':'Pin',
                          'Importer State':'State' ,
                          'Importer Phone':'Phone',
                          'Importer Email':'E-mail',
                          'Unit':'UNIT' ,
                          'Invoice No':'Inv No' })

results = pd.concat([df,df1,df2])

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
           'CHA Name','City', 'Contact Person','E-Mail', 'Phone', 'Pin', 'Port Of Origin','QTY', 'State', 'UNIT',
           'Unit Rate USD','FAX','Shipment Mode','Standard Qty','Standard Unit', 'Unit Rate USD']]


##########################Modifying Date from float to Real format ###########

from pyxlsb import convert_date
for i in range(len(results.iloc[:,0])):
    results.iloc[i,0]= convert_date(results.iloc[i,0])


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


results.to_csv('outputChapter85(1).csv')  
    
