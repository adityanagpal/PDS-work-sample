# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

dataset = pd.ExcelFile('3_Report_DelhiAir_Hs85_Imp_Jul16.xlsx')
print(dataset.sheet_names)
df1 = dataset.parse('Report')

X=df1.iloc[:,12].values
y1=df1["Product Description"].tolist()
y = df1.iloc[:,2].values

X1 = df1["A. Value INR"].tolist()
X2 = df1["Actual Duty Paid INR"].tolist()

for i in range(len(X)):
    if(X2[i])!='No Duty Paid':
        df1.iloc[i,19]= (float(X2[i])/float(X1[i]))*100
   

d=df1.groupby(['Product Description'])['Unit Rate INR'].mean()

duty_percent_mean = df1.groupby(['Product Description'])['Duty Percentage'].mean()


df1['Deviations']=[0 for i in range(len(X))]
df1['Duty Percentage Deviation'] = [0 for i in range(len(X))]
df1['Mean_Unit_Rate']=[0 for i in range(len(X))]
df1['Mean_Duty_Percent'] = [0 for i in range(len(X))]

e=[]  #contains the indexes of product unit rate more than 5 percent deviated
f=[]  #contains the indexes of duty percentage more than 5 percent deviated

for i in range(len(X)):
    if abs(df1.iloc[i,12]-d[y1[i]])>(d[y1[i]]/20):
        df1.iloc[i,18]='Deviation_Unit_Rate'
        e.append(i)
    df1.iloc[i,21] = d[y1[i]]   
        
for i in range(len(X)):
    if abs(df1.iloc[i,19]-duty_percent_mean[y1[i]])>(duty_percent_mean[y1[i]]/20):
        df1.iloc[i,20]='Deviation_Duty_Percent'
        f.append(i)  
    df1.iloc[i,22] = duty_percent_mean[y1[i]]    
    


writer = pd.ExcelWriter('output2.xlsx')
df1.to_excel(writer,'Sheet1')
writer.save()







     
