# -*- coding: utf-8 -*-
"""
Created on Mon Mar 30 09:07:35 2020

@author: SahilPatil
"""
import openpyxl
from openpyxl import load_workbook
import pandas as pd
import numpy as np

#CRM Workbook Import
crm = pd.read_excel(r"C:\Users\SahilPatil\Partners MGU\Reports - Documents\Private\CRM Imports\DY to CRM Import - 2020-03-30T18_58_22.4350523Z.xlsx")

#Import DY Info and clean data
dy = pd.read_excel(r"C:\Users\SahilPatil\Partners MGU\Reports - Documents\Private\DY Reports\All Cases.xlsx")
columns = ['Effective','Prospect','City','St','Region','Producer','Carrier','Uwr','Mkt','Ren','Status','Received','Due','Orig Quoted','Last Quoted','Administrator','Administrator Category','Office','Producer Category']
dy.columns = columns
dy = dy[3:]
dy = dy.reset_index(drop = True)

dy['Effective'] = pd.to_datetime(dy['Effective'])
dy['Effective'] = dy['Effective'].dt.strftime("%m/%d/%Y")
dy['Received'] = pd.to_datetime(dy['Received'])
dy['Received'] = dy['Received'].dt.strftime("%m/%d/%Y")
dy['Due'] = pd.to_datetime(dy['Effective'])
dy['Due'] = dy['Due'].dt.strftime("%m/%d/%Y")
dy['Orig Quoted'] = pd.to_datetime(dy['Orig Quoted'])
dy['Orig Quoted'] = dy['Orig Quoted'].dt.strftime("%m/%d/%Y")
dy['Last Quoted'] = pd.to_datetime(dy['Last Quoted'])
dy['Last Quoted'] = dy['Last Quoted'].dt.strftime("%m/%d/%Y")

#Export file
export = pd.DataFrame(columns = ['Account Name'])
export['Account Name'] = dy['Producer'].unique()
export = export.sort_values(['Account Name'], ascending = [True])
export = export.reset_index(drop = True)


#Create current and new producers
current_producers = set(crm['Name in DY'])

#Extract missing producers list
export['Is it in Current Export?'] = list(export['Account Name'].apply(lambda x: "Yes" if x in current_producers else "No"))
missing_producers = pd.DataFrame(export[export['Is it in Current Export?'] == 'No'])
missing_producers = missing_producers['Account Name']
missing_producers = missing_producers[~missing_producers.str.contains("Test Producer")]
missing_producers.to_excel(r"C:\Users\SahilPatil\Partners MGU\Reports - Documents\Private\CRM Imports\Missing Producers List.xlsx", index = None)