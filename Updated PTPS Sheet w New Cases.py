# -*- coding: utf-8 -*-
"""
Created on Mon Mar 23 09:17:45 2020

@author: SahilPatil
"""


import pandas as pd
import os
import glob
from oauth2client.service_account import ServiceAccountCredentials
import pygsheets
import gspread
import numpy as np

#Authorization for PTPS Sheet
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds2 = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\SahilPatil\OneDrive - Partners MGU\Desktop\Python Files\ptps-270913-61e509b076ef.json", scope)
client1 = gspread.authorize(creds2)



mypath = r"C:\Users\SahilPatil\Partners MGU\PBM Team - Documents\Initial Partners Repricing"
os.chdir(mypath)
extension = 'xlsm'
mylist = glob.glob('*.{}'.format(extension))


effective = []
prospect = []
group_Size = []
members = []
total_Claims = []
gross_Premium = []
claims_Months = []
total_Cost = []
ptpssavings = []
premium_reduction = []
revenue = []
mgu_fee = []
spec_deduct = []


for f in mylist:
    df = pd.read_excel(f, header = None)
    effective.append(df[1][1])
    prospect.append(df[1][0])
    group_Size.append(df[4][1])
    members.append(round(df[5][1]))
    total_Claims.append(round(df[6][1]))
    gross_Premium.append(df[8][1])
    claims_Months.append(df[1][24])
    total_Cost.append(df[1][25])
    ptpssavings.append(df[6][32])
    premium_reduction.append((format((abs(df[5][85])),'.2%')))
    revenue.append(round(df[6][1]*3.50))
    mgu_fee.append(format((abs(df[4][79])),'.2%'))
    spec_deduct.append(df[5][69])
    
        
main = {'Effective':effective,'Prospect':prospect,'Group Size':group_Size,'Members':members,'Total Claims':total_Claims,'Gross Premium':gross_Premium,'Claims Months':claims_Months,'Total Cost':total_Cost,'PTPS Savings w PH':ptpssavings,'Premium Reduction': premium_reduction,'PBM Revenue':revenue,'Max Reduction in MGU Fee':mgu_fee,'Spec Deduct Quoted':spec_deduct}

data = pd.DataFrame(main)

#Effective date format
data['Effective'] = pd.to_datetime(data['Effective'])
data['Effective'] = data['Effective'].dt.strftime("%m/%d/%Y")

#Sort
data = data.sort_values(['Effective','Prospect'], ascending = [True,True])
data = data.reset_index(drop = True)

#Import all cases
cases = pd.read_excel(r"C:\Users\SahilPatil\Partners MGU\PBM Team - Documents\Prospective Cases\All Cases.xlsx")
cases = cases.rename(columns = cases.iloc[2])
cases = cases[3:]
cases = cases.reset_index(drop = True)

#Change effective format and add key
cases['Effective'] = pd.to_datetime(cases['Effective'])
cases['Effective'] = cases['Effective'].dt.strftime("%m/%d/%Y")
cases['Key'] = cases['Effective'] + ' ' + cases['Prospect']


#Create Key
data['Key'] = data['Effective'] + ' ' + data['Prospect']
data['Uwr'] = ''

#Join for Uwr and Mkt
final = pd.merge(data,cases, on = 'Key', how = 'left')
final = final.drop(['Uwr_x','Effective_y','Prospect_y','City','St','Region', 'Carrier', 'Ren','Status', 'Due', 'Quoted','Quoted','Producer Category','Administrator'], axis = 1)
final = final.rename(columns = {'Effective_x':'Effective','Prospect_x':'Prospect','Uwr_y':'Uwr'})
final = final.drop_duplicates(subset = ['Prospect'])

#Format Received
final['Received'] = pd.to_datetime(final['Received'])
final['Received'] = final['Received'].dt.strftime("%m/%d/%Y")

#Join for Spec level
spec = pd.read_excel(r"C:\Users\SahilPatil\Partners MGU\PBM Team - Documents\Data\Spec Level Extraction.xlsx")
spec['Date'] = pd.to_datetime(spec['Date'])
spec['Date'] = spec['Date'].dt.strftime("%m/%d/%Y")
spec['Key'] = spec['Date'] + ' ' + spec['Policyholder']
final = pd.merge(final, spec, on = 'Key', how = 'left')
final = final.drop(['Date','Policyholder'], axis = 1)
final = final.drop_duplicates(subset = ['Prospect'])
final['Spec Deduct'] = final['Spec Deduct'].fillna(0)
final['Spec Deduct'] = np.where(final['Spec Deduct'] == 0, final['Spec Deduct Quoted'], final['Spec Deduct'])
final = final.drop(['Spec Deduct Quoted'], axis = 1)
final['Spec Deduct'] = final['Spec Deduct'].fillna(0)




final = final[['Effective','Received','Prospect','Group Size','Members','Total Claims','Gross Premium','Claims Months','Total Cost','PTPS Savings w PH','Spec Deduct','Premium Reduction','Max Reduction in MGU Fee','PBM Revenue','Key','Producer','Uwr','Mkt']]

#Format numbers as currency
final['Gross Premium'] = final['Gross Premium'].apply(lambda x: '$' + str('{:,}'.format(round(x))))
final['Total Cost'] = final['Total Cost'].apply(lambda x: '$' + str('{:,}'.format(round(x))))
final['PTPS Savings w PH'] = final['PTPS Savings w PH'].apply(lambda x: '$' + str('{:,}'.format(round(x))))
final['PBM Revenue'] = final['PBM Revenue'].apply(lambda x: '$' + str('{:,}'.format(round(x))))
final['Spec Deduct'] = final['Spec Deduct'].apply(lambda x: '$' + str('{:,}'.format(round(x))))



#
gc = pygsheets.authorize(service_file = r"C:\Users\SahilPatil\OneDrive - Partners MGU\Desktop\Python Files\ptps-270913-61e509b076ef.json")
sh = gc.open('Partners Total Pharmacy Solution')

case_details = sh[2]

case_details.delete_rows(index = 3, number = 500)
case_details.set_dataframe(final,(1,1))

os.remove(r"C:\Users\SahilPatil\Partners MGU\PBM Team - Documents\Control Sheet\Partners Total Pharmacy Solution.xlsx")
final.to_excel(r"C:\Users\SahilPatil\Partners MGU\PBM Team - Documents\Control Sheet\Partners Total Pharmacy Solution.xlsx", index = False)
