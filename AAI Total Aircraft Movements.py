# This is a sample Python script.
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import pymssql
import numpy as np
conn = pymssql.connect('database-1.cmtpadv1tggf.ap-south-1.rds.amazonaws.com','OPERATIONSDBOWNER','OPERATIONSDBOWNER@4321','ForOperationsData')
s ='''SELECT * FROM [ForOperationsData].[dbo].[IN_Monthly_Aircraft_Movements]'''
df = pd.read_sql(s, conn)
df1 = df.pivot_table(index='Report_Month', columns='Operation', values='value_rm', aggfunc=np.sum)
df1.sort_values(by = 'Report_Month', ascending=False, inplace=True)
df1.rename(columns = {'Report_Month':'Months','Domestic':'Report_Month_Value(Domestic)','International':'Report_Month_Value(International)'}, inplace = True)
df1['index'] = df1.index
first_column = df1.pop('index')
df1.insert(0, 'index', first_column)
cols = ['Report_Month_Value(Domestic)','Report_Month_Value(International)']
df1['index'] = pd.to_datetime(df1['index'], dayfirst=True)
df3 = df1.set_index('index')[cols]
d = {'Report_Month_Value(Domestic)':'(Domestic)','Report_Month_Value(International)':'(International)'}
df1 = df3.shift(1, freq='MS')
df2 = df3.shift(12, freq='MS')
df4 = df3.shift(36, freq='MS')
df11 = df3.sub(df1).div(df1).rename(columns=d).add_prefix('MOM Change')
df22 = df3.sub(df2).div(df2).rename(columns=d).add_prefix('Change_SMLY')
df44 = df3.sub(df4).div(df4).rename(columns=d).add_prefix('Change Pre-Covid')
df3 = pd.concat([df3, df11.reindex(df3.index), df22.reindex(df3.index), df44.reindex(df3.index)], axis=1)
print(df3)
df3.to_excel("Total Aircraft Movements.xlsx", sheet_name='Ratio(Total Aircraft Movements)', index=True)