#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 26 14:26:42 2018

@author: aditya
"""


import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

############### READING THE DATASET ###################
df1 = pd.read_excel('1_Report_BangaloreAir_Hs85_Imp_Jun16.xlsx',sheet_name="Report")
df2 = pd.read_excel('1_Report_BombayAir_Hs85_Imp_Jun16.xlsx',sheet_name="Report")
df3 = pd.read_excel('1_Report_DelhiAir_Hs85_Imp_Jul16.xlsx',sheet_name="Report")
df4 = pd.read_excel('1_Report_DelhiAir_Hs85_Imp_Jun16.xlsx',sheet_name="Report")
df5 = pd.read_excel('2_Report_BangaloreAir_Hs85_Imp_Jun16.xlsx',sheet_name="Report")
df6 = pd.read_excel('2_Report_BombayAir_Hs85_Imp_Jun16.xlsx',sheet_name="Report")
df7 = pd.read_excel('2_Report_DelhiAir_Hs85_Imp_Jul16.xlsx',sheet_name="Report")
df8 = pd.read_excel('2_Report_DelhiAir_Hs85_Imp_Jun16.xlsx',sheet_name="Report")
df9 = pd.read_excel('3_Report_DelhiAir_Hs85_Imp_Jul16.xlsx',sheet_name="Report")
df10 = pd.read_excel('3_Report_DelhiAir_Hs85_Imp_Jun16.xlsx',sheet_name="Report")
df11 = pd.read_excel('3_Report_MadrasAir_Hs85_Imp_Jun16.xlsx',sheet_name="Report")
df12 = pd.read_excel('Bombay_air_85_FEB16IMP.xlsx',sheet_name="Results")
df13 = pd.read_excel('Bombay_air_85_JUL14IMP.xlsx',sheet_name="Results")
df14 = pd.read_excel('Bombay_air_85_JUL15IMP.xlsx',sheet_name="Results")
df15 = pd.read_excel('Bombay_air_85_OCT14IMP.xlsx',sheet_name="Results")
df16 = pd.read_excel('Bombay_air_85_SEP14IMP.xlsx',sheet_name="Results")
df17 = pd.read_excel('Bombay_air_85_SEP15-OCT15IMP.xlsx',sheet_name="Results")
df18 = pd.read_excel('BOMBAY_AIR30JAN16IMP.xlsx',sheet_name="Results")
df19 = pd.read_excel('BOMBAY_AIR85OCT15IMP.xlsx',sheet_name="Results")
df20 = pd.read_excel('BOMBAY_AIR85SEP15IMP.xlsx',sheet_name="Results")
df21 = pd.read_excel('CALCUTTA_SEA85MAR15IMP.xlsx',sheet_name="Results")
df22 = pd.read_excel('Delhi_air_90_-JAN16IMP.xls',sheet_name="Results")
df23 = pd.read_excel('DELHI_AIR85JUL14IMP.xlsx',sheet_name="Results")
df24 = pd.read_excel('DELHI_AIR85MAY14IMP.xlsx',sheet_name="Results")
df25 = pd.read_excel('JNPT98JAN16IMP.xlsx',sheet_name="Results")
df26 = pd.read_excel('MADRAS_AIR85JAN16IMP.xlsx',sheet_name="Results")
df27 = pd.read_excel('MADRAS_SEA85OCT14IMP.xlsx',sheet_name="Results")
df28 = pd.read_excel('TKD85JAN16IMP.xlsx',sheet_name="Results")




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

cl = len(results.columns)-4
for i in range(j):
    if ((X2[i])!='No Duty Paid' or X2[i]!=0.0):
        continue
    else:
        results.iloc[i,cl]= (float(X2[i])/float(X1[i]))*100
########################### CALCULATING THE ABSOLUTE VALUE OF PRODUCTS #########

unit_rate_mean = results.groupby(['Product Description'])['Unit Rate INR'].mean()

duty_percent_mean = results.groupby(['Product Description'])['Duty Percentage'].mean()  

######################## DETECTING THE DEVIATIONS ##################
dul = len(results.columns)-5
dpl = len(results.columns)-3
dpmv = len(results.columns)-1
urmv = len(results.columns)-2


for i in range(j):
    if abs(X[i]-unit_rate_mean[y1[i]])>(unit_rate_mean[y1[i]]/20):
        results.iloc[i,dul]='Deviation_Unit_Rate'
    if abs(float(results.iloc[i,cl])-duty_percent_mean[y1[i]])>(duty_percent_mean[y1[i]]/20):
        results.iloc[i,dpl]='Deviation_Duty_Percent'
    results.iloc[i,dpmv] = duty_percent_mean[y1[i]]        
    results.iloc[i,urmv] = unit_rate_mean[y1[i]]   
        
#################### REARRANGING THE COLUMNS #######################
cols= results.columns.tolist()
print(cols)
results = results[['Date','HS Code','Product Description','Importer Name','Duty Percentage',
           'Deviations_Unit_Rate', 'Duty_Percent_Deviation', 'Mean_Unit_Rate', 'Mean_Duty_Percent',
           'Unit Rate INR','A. Value INR','Actual Duty Paid INR','BE Number','IEC',
           'Applicable Duty INR','Port Of Destination','Country Of Origin','Address1',
           'CHA Name', 'CHA No', 'City', 'Contact Person','E-Mail',
           'Foreign Country', 'Phone', 'Pin', 'Port Of Origin','QTY', 'State', 'UNIT',
           'Unit Rate USD','Address2','FAX','Shipment Mode','Standard Qty','Standard Unit',
           'Standard Unit Rate', 'Unit Rate USD','Unnamed: 14', 'Unnamed: 4', 'Unnamed: 6',
          'Value + Duty in INR', 'Value + Duty in USD','Unnamed: 14', 'Unnamed: 4', 
          'Unnamed: 6','Value + Duty in INR', 'Value + Duty in USD']]
 
 
############## CONVERTING INTO EXCEL #############################

writer = pd.ExcelWriter('output85.xlsx')
results.to_excel(writer,'Sheet1')
writer.save()
