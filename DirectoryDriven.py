#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 19 15:33:30 2018

@author: aditya
"""
from os import walk

f = []
for (dirpath, dirnames, filenames) in walk('/home/aditya/Videos'):
    f.extend(filenames)
    break

for dire in dirnames:
    i=0
    f[i] = []
    for (dirpaths, dirname, filename) in walk('/home/aditya/Videos/'+dire):
        f[i].extend(filenames)
        break
    i=i+1
    
f = [[]] * len(dirnames)
    
    
import glob
for dire in dirnames:
    i=0
    mylist = [f[i] for f[i] in glob.glob(dire+"/"+"*.pdf")]
    i=i+1
     