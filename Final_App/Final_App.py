#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jan  6 23:38:20 2019

@author: aditya
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pyxlsb import open_workbook as open_xlsb

import tkinter as tk
from tkinter import filedialog as fd
from tkinter import *
def openfile():
    global filename
    filename = fd.askopenfilename(title="Open file")

def printfile():
    if filename[len(filename)-1]=='b':
        results = []

        with open_xlsb(filename) as wb:
            with wb.get_sheet(2) as sheet:
                for row in sheet.rows():
                    results.append([item.v for item in row])

        results = pd.DataFrame(results[1:], columns=results[0])

        xd = (results.columns).tolist()
    

        root = Tk()
        root.geometry('400x400')                      #ceating list of cloumns in scroll bar
        scrollbar = Scrollbar(root)
        scrollbar.pack( side = RIGHT, fill = Y )

        mylist = Listbox(root, yscrollcommand = scrollbar.set )
        for line in range(len(xd)):
           mylist.insert(END, xd[line])

        mylist.pack( side = LEFT, fill = BOTH )
        scrollbar.config( command = mylist.yview )

        mainloop()

    elif filename[len(filename)-1]=='x' or filename[len(filename)-1]=='s':  
        dataset = pd.ExcelFile(filename)
        print(dataset.sheet_names) # After getting the name we found the sheet Report to be used
        results = dataset.parse('Report')                           #reading of files
        
        xd = (results.columns).tolist()
    

        root = Tk()
        root.geometry('400x400')                      #ceating list of cloumns in scroll bar
        scrollbar = Scrollbar(root)
        scrollbar.pack( side = RIGHT, fill = Y )

        mylist = Listbox(root, yscrollcommand = scrollbar.set )
        for line in range(len(xd)):
           mylist.insert(END, xd[line])

        mylist.pack( side = LEFT, fill = BOTH )
        scrollbar.config( command = mylist.yview )

        mainloop()
        
    elif filename[len(filename)-1]=='v':  
        
        results = pd.read_csv(filename)                           #reading of files
        
        xd = (results.columns).tolist()
    

        root = Tk()
        root.geometry('400x400')                      #ceating list of cloumns in scroll bar
        scrollbar = Scrollbar(root)
        scrollbar.pack( side = RIGHT, fill = Y )

        mylist = Listbox(root, yscrollcommand = scrollbar.set )
        for line in range(len(xd)):
           mylist.insert(END, xd[line])

        mylist.pack( side = LEFT, fill = BOTH )
        scrollbar.config( command = mylist.yview )

        mainloop()    
    
def printspinbox():
    print(spinbox.get())

window = tk.Tk()
window.geometry("300x300")
filename = ''# global variable

tk.Button(window, text='Browse', command=openfile).pack()
tk.Button(window, text='Print filename', command=printfile).pack()

def add_text():
   label1 = Label(root, text="You have entered the information for calculating the Deviation ")
   label1.pack()
       
def retrieve_input():
   global Unit_Rate
   global Product_Description
   global Accessible_Duty
   global Actual_Duty
   global Importer_Name
   Unit_Rate = veh_reg_text_box.get()
   Product_Description = time_text_box.get()
   Accessible_Duty = distance_text_box.get()
   Actual_Duty = Actual_Duty_box.get()
   Importer_Name = Importer_Name_box.get()
   
   results = []

   with open_xlsb(filename) as wb:
       with wb.get_sheet(2) as sheet:
           for row in sheet.rows():
               results.append([item.v for item in row])

   results = pd.DataFrame(results[1:], columns=results[0])

   xd = (results.columns).tolist()
   ############ EXTRACTING THE IMPORTANT COLUMS - Product Description and Unit Rate ####

   X=results[Unit_Rate].tolist()
   y1=results[Product_Description].tolist()

   ############ EXTRACTING THE IMPORTANT COLUMS - Product Description and Unit Rate ####

   X1 = results[Accessible_Duty].tolist()
   X2 = results[Actual_Duty].tolist()

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

   unit_rate_mean = results.groupby([Product_Description])[Unit_Rate].mean()

   duty_percent_mean = results.groupby([Product_Description])['Duty Percentage'].mean()  

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

   ##################### REARRANGING INDEX GROUP BY PRODUCT DESCRIPTION ########

   results=results.reset_index()
   df = results.groupby([Importer_Name,Product_Description])['index'].apply(list)
   List = []
   del results['index']
   for i in range(len(df)):
       List+=df[i]
    
   results = results.reindex(List) 
   
   results = results.reset_index()
   del results['index']
   results = results.reset_index()

   customer_sum = results.groupby([Importer_Name])['A. Value INR'].sum()

   customer_sum = customer_sum.sort_values(ascending = False)

   customer_sum = customer_sum.reset_index()

   dfc = results.groupby([Importer_Name])['index'].apply(list)

   Ll = []
   for i in range(len(dfc)):
       Ll+=dfc[customer_sum[Importer_Name][i]]
    
    
   results = results.reindex(Ll)

   dfd = np.array([((results.iloc[dfc[i],:].values).tolist()) for i in range(len(dfc))])

   fghg = [[]]

   for i in range(len(dfd)):
       for j in range(len(dfd[i])):
           fghg.append(dfd[i][j])
       fghg.append([0 for i in range(len(results.columns))])
       fghg.append([0,0,0,0,'Accessible Sum',customer_sum['A. Value INR'][i]])
       fghg.append([0 for i in range(len(results.columns))])
    
   dfcfd = pd.DataFrame(fghg,columns=['index']+xd+['Deviations_Unit_Rate','Duty Percentage',
                                         'Duty_Percent_Deviation','Mean_Unit_Rate',
                                         'Mean_Duty_Percent'])# This will give error due to shape mismatch but do some adjustment



 
   ############## CONVERTING INTO EXCEL #############################


   dfcfd.to_csv(filename+'output(1).csv')

       

veh_reg_label = tk.Label(window, text="Unit Rate:")
veh_reg_label.pack()

veh_reg_text_box = tk.Entry(window, bd=1)
veh_reg_text_box.pack()

distance_label = tk.Label(window, text="Accessible Duty:")
distance_label.pack()

distance_text_box = tk.Entry(window, bd=1)
distance_text_box.pack()

time_label = tk.Label(window, text="Product Description:")
time_label.pack()

time_text_box = tk.Entry(window, bd=1)
time_text_box.pack()

Actual_Duty_label = tk.Label(window, text="Actual Duty:")
Actual_Duty_label.pack()

Actual_Duty_box = tk.Entry(window, bd=1)
Actual_Duty_box.pack()

Importer_Name_label = tk.Label(window, text="Importer Name:")
Importer_Name_label.pack()

Importer_Name_box = tk.Entry(window, bd=1)
Importer_Name_box.pack()

enter_button = tk.Button(window, text="Enter", command=retrieve_input)
enter_button.pack()

'''spinbox = tk.Spinbox(window, from_=0, to=10)
spinbox.pack(pady=10)
tk.Button(window, text='Print spinbox value', command=printspinbox).pack()'''

window.mainloop()