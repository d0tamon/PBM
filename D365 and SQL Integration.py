# -*- coding: utf-8 -*-
"""
Created on Wed Jul 15 08:46:51 2020

@author: SahilPatil
"""


# -*- coding: utf-8 -*-
"""
Created on Mon Jun 29 07:30:39 2020

@author: SahilPatil
"""

import adal
from dynamics365crm.client import Client 
import pandas as pd
import psycopg2
import numpy as np

pd.set_option('precision',0)
RESOURCE_URI = 'RESOURCE URL'
# O365 credentials for authentication w/o login prompt
USERNAME = 'USERNAME@COMPANYNAME.COM'
PASSWORD = 'PASSWORD'
# Azure Directory OAUTH 2.0 AUTHORIZATION ENDPOINT
AUTHORIZATION_URL = 'AUTH URL'

token_response = adal.acquire_token_with_username_password(
        AUTHORIZATION_URL,
        USERNAME,
        PASSWORD,
        resource=RESOURCE_URI
    )

token = token_response['accessToken']
refresh = token_response['refreshToken']


client = Client(RESOURCE_URI, token)

stoken = client.set_token(token)

get_accounts = client.get_accounts()
accounts = get_accounts['value']

df = pd.DataFrame(accounts)
df = df[df['new_nameindy'].notnull()]
columns = list(df.columns)

df1 = df[['accountid','name','new_nameindy','new_lastrfpreceived','new_receivedthisuwyear','new_alltimereceived','new_lastquoteddate','new_quotedthisuwyear','new_alltimequoted','new_lastsolddate','new_soldthisuwyear','new_alltimesold','new_firstquoteddate','new_rfpinfoupdatedate']]


con = psycopg2.connect(user = "TEST",
                       password = "PASS",
                       host = "HOST",
                       port = "PORT",
                       database = "DATABASE")

cursor = con.cursor()

##Query
postgreSQL_select_Query = ''' SET SESSION myvariables.uwyearstart = '2020-02-01 00:00:00';
SET SESSION myvariables.uwyearend = '2021-01-31 23:59:59';
SELECT T1.accountid, 
			T1.name, 
			T1.new_nameindy, 
			MAX(T2.received) AS "new_lastrfpreceived", 
			MAX(T2.last_quoted) AS "new_lastquoteddate", 
			MIN(T2.orig_quoted) AS "new_firstquoteddate", 
			COUNT(T2.received) AS "new_alltimereceived",
			COUNT(T2.orig_quoted) AS "new_alltimequoted",
			COUNT(CASE WHEN T2.orig_quoted BETWEEN (SELECT current_setting('myvariables.uwyearstart')::date) AND (SELECT current_setting('myvariables.uwyearend')::date) THEN T2.orig_quoted
					ELSE NULL END) AS new_quotedthisuwyear,
			COUNT(CASE WHEN T2.received BETWEEN (SELECT current_setting('myvariables.uwyearstart')::date) AND (SELECT current_setting('myvariables.uwyearend')::date) THEN T2.received
					ELSE NULL END) AS new_receivedthisuwyear,
			COUNT(CASE WHEN T2.status = 'Sold' THEN T2.prospect ELSE NULL END) AS new_alltimesold,
			MAX(CASE WHEN T2.status = 'Sold' THEN T2.effective ELSE NULL END) AS new_lastsolddate,
			COUNT(CASE WHEN T2.status = 'Sold' AND T2.received BETWEEN (SELECT current_setting('myvariables.uwyearstart')::date) AND (SELECT current_setting('myvariables.uwyearend')::date) THEN T2.received ELSE NULL END) AS new_soldthisuwyear,
			NOW() AS "new_rfpinfoupdateddate"
	FROM crm_tbl AS T1
	LEFT JOIN underwriting_tbl AS T2
	ON T1.new_nameindy = T2.producer
	GROUP BY T1.new_nameindy, T1.accountid
	ORDER BY T1.new_nameindy ASC,
	T1.accountid ASC,
	T1.name;
 '''
cursor.execute(postgreSQL_select_Query)

mobile_records = cursor.fetchall() 
colnames = [desc[0] for desc in cursor.description]

##Get df
sql_query = pd.DataFrame(mobile_records, columns = colnames)


sql_query = sql_query.astype({'new_receivedthisuwyear':'str'}, errors = 'ignore')
sql_query['new_receivedthisuwyear'] = sql_query['new_receivedthisuwyear'].replace('\.0','',regex = True)
sql_query['new_receivedthisuwyear'] = sql_query['new_receivedthisuwyear'].replace('None',np.nan)

sql_query = sql_query.astype({'new_alltimereceived':'str'}, errors = 'ignore')
sql_query['new_alltimereceived'] = sql_query['new_alltimereceived'].replace('\.0','',regex = True)
sql_query['new_alltimereceived'] = sql_query['new_alltimereceived'].replace('None',np.nan)

sql_query = sql_query.astype({'new_alltimequoted':'str'}, errors = 'ignore')
sql_query['new_alltimequoted'] = sql_query['new_alltimequoted'].replace('\.0','',regex = True)
sql_query['new_alltimequoted'] = sql_query['new_alltimequoted'].replace('None',np.nan)

sql_query = sql_query.astype({'new_quotedthisuwyear':'str'}, errors = 'ignore')
sql_query['new_quotedthisuwyear'] = sql_query['new_quotedthisuwyear'].replace('\.0','',regex = True)
sql_query['new_quotedthisuwyear'] = sql_query['new_quotedthisuwyear'].replace('None',np.nan)


sql_query = sql_query.astype({'new_alltimesold':'str'}, errors = 'ignore')
sql_query['new_alltimesold'] = sql_query['new_alltimesold'].replace('\.0','',regex = True)
sql_query['new_alltimesold'] = sql_query['new_alltimesold'].replace('None',np.nan)

sql_query = sql_query.astype({'new_soldthisuwyear':'str'}, errors = 'ignore')
sql_query['new_soldthisuwyear'] = sql_query['new_soldthisuwyear'].replace('\.0','',regex = True)
sql_query['new_soldthisuwyear'] = sql_query['new_soldthisuwyear'].replace('None',np.nan)


sql_query['new_lastrfpreceived'] = pd.to_datetime(sql_query['new_lastrfpreceived'])
sql_query['new_lastrfpreceived'] = sql_query['new_lastrfpreceived'].dt.strftime('%Y-%m-%dT07:%M:%SZ')
sql_query['new_lastrfpreceived'] = np.where(sql_query['new_lastrfpreceived'].isnull(), None, sql_query['new_lastrfpreceived'])

sql_query['new_lastquoteddate'] = pd.to_datetime(sql_query['new_lastquoteddate'])
sql_query['new_lastquoteddate'] = sql_query['new_lastquoteddate'].dt.strftime('%Y-%m-%dT07:%M:%SZ')
sql_query['new_lastquoteddate'] = np.where(sql_query['new_lastquoteddate'].isnull(), None, sql_query['new_lastquoteddate'])

sql_query['new_firstquoteddate'] = pd.to_datetime(sql_query['new_firstquoteddate'])
sql_query['new_firstquoteddate'] = sql_query['new_firstquoteddate'].dt.strftime('%Y-%m-%dT07:%M:%SZ')
sql_query['new_firstquoteddate'] = np.where(sql_query['new_firstquoteddate'].isnull(), None, sql_query['new_firstquoteddate'])

sql_query['new_lastsolddate'] = pd.to_datetime(sql_query['new_lastsolddate'])
sql_query['new_lastsolddate'] = sql_query['new_lastsolddate'].dt.strftime('%Y-%m-%dT07:%M:%SZ')
sql_query['new_lastsolddate'] = np.where(sql_query['new_lastsolddate'].isnull(), None, sql_query['new_lastsolddate'])

sql_query['new_rfpinfoupdateddate'] = pd.to_datetime(sql_query['new_rfpinfoupdateddate'])
sql_query['new_rfpinfoupdateddate'] = sql_query['new_rfpinfoupdateddate'].dt.strftime('%Y-%m-%dT07:%M:%SZ')
sql_query['new_rfpinfoupdateddate'] = np.where(sql_query['new_rfpinfoupdateddate'].isnull(), None, sql_query['new_rfpinfoupdateddate'])



for a,b in zip(sql_query['accountid'],sql_query['new_receivedthisuwyear']):
    client.update_account(a, new_receivedthisuwyear = b)
    
for a,b in zip(sql_query['accountid'],sql_query['new_alltimereceived']):
    client.update_account(a, new_alltimereceived = b)
    
for a,b in zip(sql_query['accountid'],sql_query['new_alltimequoted']):
    client.update_account(a, new_alltimequoted = b)
    
for a,b in zip(sql_query['accountid'],sql_query['new_quotedthisuwyear']):
    client.update_account(a, new_quotedthisuwyear = b)
    
for a,b in zip(sql_query['accountid'],sql_query['new_alltimesold']):
    client.update_account(a, new_alltimesold = b)
    
for a,b in zip(sql_query['accountid'],sql_query['new_soldthisuwyear']):
    client.update_account(a, new_soldthisuwyear = b)

for account_id,new_lastrfpreceived in zip(sql_query.accountid,sql_query.new_lastrfpreceived):
    client.update_account(account_id, new_lastrfpreceived = new_lastrfpreceived)
    
for account_id,new_lastquoteddate in zip(sql_query.accountid,sql_query.new_lastquoteddate):
    client.update_account(account_id, new_lastquoteddate = new_lastquoteddate)

for account_id,new_firstquoteddate in zip(sql_query.accountid,sql_query.new_firstquoteddate):
    client.update_account(account_id, new_firstquoteddate = new_firstquoteddate)

for account_id,new_lastsolddate in zip(sql_query.accountid,sql_query.new_lastsolddate):
    client.update_account(account_id, new_lastsolddate = new_lastsolddate)

for account_id,new_rfpinfoupdateddate in zip(sql_query.accountid,sql_query.new_rfpinfoupdateddate):
    client.update_account(account_id, new_rfpinfoupdateddate = new_rfpinfoupdateddate)


