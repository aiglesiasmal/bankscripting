import pandas as pd                       #to perform data manipulation and analysis
import numpy as np                        #to cleanse data
from datetime import datetime             #to manipulate dates
import plotly.express as px               #to create interactive charts
import plotly.graph_objects as go         #to create interactive charts
from jupyter_dash import JupyterDash      #to build Dash apps from Jupyter environments
import dash_core_components as dcc        #to get components for interactive user interfaces
import dash_html_components as html       #to compose the dash layout using Python structures
import requests
import xlsxwriter as wrt
import xlrd as rd
import openpyxl as opxl

excel1 = 'movimientos.xls'
excel2 = 'GASTOS.xlsx'
excel3 = 'Movimientos (1).xls'

df1 = pd.read_excel (excel1,skiprows=4)
df2 = pd.read_excel(excel2, sheet_name='Hoja de gastos') #acelerar para que no devuelva errores
df3 = pd.read_excel(excel3,skiprows=3)

nrows1=df1.index[df1.FECHA == 'Total Crédito']- 1
tc = 6+nrows1 #60 es el número de cargos descontando los de espera, hacer count?
df1 = df1.drop(index = tc)
df1 = df1.dropna()
df3 = df3.iloc[3:,] #errases first 3 rows

df1 = df1.rename(columns={'FECHA': 'Date','COMERCIO/CAJERO': 'Description','IMPORTE': 'Debits'})
df3 = df3.rename(columns={'FECHA VALOR': 'Date','DESCRIPCION': 'Description','IMPORTE': 'Debits'})

values1 = df1[['Date','Description','Debits']]
values2 = df2[['Date','Description','Debits']]
values3 = df3[['Date','Description','Debits']]

dataframes = [values1,values3]
join = pd.concat(dataframes)
join.to_excel("output.xlsx")

import os, os.path
import win32com.client

if os.path.exists("GASTOS.xlsm"):
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(os.path.abspath("GASTOS.xlsm"), ReadOnly=1)
    xl.Application.Run("GASTOS.xlsm!Module1.Copypaste")
##    xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl

#copying and pasting file
import shutil

original = r'C:\Users\varit\Python Projects\GASTOS.xlsm'
target = r'C:\Users\varit\OneDrive\GASTOS\GASTOS.xlsm'

shutil.copyfile(original,target)
