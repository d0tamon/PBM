# -*- coding: utf-8 -*-
"""
Created on Fri Mar 27 13:29:30 2020

@author: SahilPatil
"""


import os
import glob
import pandas as pd

mypath = r"C:\Update Files"
os.chdir(mypath)
extension = 'xlsm'
mylist = glob.glob('*.{}'.format(extension))

for f in mylist:
    os.rename(f,f+' '+ 'New' + '.xlsm')