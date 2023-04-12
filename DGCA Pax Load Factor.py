# This is a sample Python script.
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import pymssql
import numpy as np
conn = pymssql.connect('database-1.cmtpadv1tggf.ap-south-1.rds.amazonaws.com','OPERATIONSDBOWNER','OPERATIONSDBOWNER@4321','ForOperationsData')
s ='''SELECT [Report_Month]
      ,[Operating_Airline]
      ,[Airline_Operations]
      ,[Airline_Service]
      ,[Passengers_kms_in_Thousands]
      ,[Available_Seat_kms_in_Thousands]
      ,Passengers_kms_in_Thousands/Available_Seat_kms_in_Thousands as PAX_Load_Factor
  FROM [ForOperationsData].[dbo].[IN_Monthly_DGCA_Airline_Traffic]
  WHERE Available_Seat_kms_in_Thousands>0 and Available_Seat_kms_in_Thousands is not null'''
df = pd.read_sql(s, conn)
df1 = df.pivot_table(index='Report_Month', columns='Airline_Service', values='PAX_Load_Factor', aggfunc=np.sum)
df1.sort_values(by = 'Report_Month', ascending=False, inplace=True)
df1.rename(columns = {'Report_Month':'Months','Domestic':'PAX_Load_Factor(Domestic)','International':'PAX_Load_Factor(International)'}, inplace = True)
df1['index'] = df1.index
first_column = df1.pop('index')
df1.insert(0, 'index', first_column)
cols = ['PAX_Load_Factor(Domestic)','PAX_Load_Factor(International)']
df1['index'] = pd.to_datetime(df1['index'], dayfirst=True)
df3 = df1.set_index('index')[cols]
d = {'PAX_Load_Factor(Domestic)':'(Domestic)','PAX_Load_Factor(International)':'(International)'}
df1 = df3.shift(1, freq='MS')
df2 = df3.shift(12, freq='MS')
df4 = df3.shift(36, freq='MS')
df11 = df3.sub(df1).div(df1).rename(columns=d).add_prefix('MOM Change')
df22 = df3.sub(df2).div(df2).rename(columns=d).add_prefix('Change_SMLY')
df44 = df3.sub(df4).div(df4).rename(columns=d).add_prefix('Change Pre-Covid')
df3 = pd.concat([df3, df11.reindex(df3.index), df22.reindex(df3.index), df44.reindex(df3.index)], axis=1)
print(df3)
df3.to_excel("Ratio(PAX_Load_Factor).xlsx", sheet_name='Ratio(PAX_Load_Factor)', index=True)