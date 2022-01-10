#Moshe Khorshidi Code Jungle

#must install for the code  
get_ipython().system('pip install TableauScraper')

#import scarper for webpage with tableau as TS 
from tableauscraper import TableauScraper as TS

url = "https://public.tableau.com/views/WYCOVID-19Dashboard/WyomingCOVID-19CaseDashboard"

ts = TS()
ts.loads(url)
workbook = ts.getWorkbook() #get the workbook that loaded

for sheet in workbook.worksheets:
    print(f"worksheet name : {sheet.name}") #show worksheet name with string 
    print(sheet.data) #show dataframe for this worksheet

# if we want specific worksheet 

ts = TS()
ts.loads(url)

ws = ts.getWorksheet('Case Age')
print(ws.data)

# get csv from the data 

url = 'https://public.tableau.com/views/WYCOVID-19Dashboard/WyomingCOVID-19CaseDashboard'
ts = TS()
ts.loads(url)
wb = ts.getWorkbook()
data = wb.getCsvData(sheetName='Case Age')
print(data)

# use pandas to move the file on .xlsx

import pandas as pd
datadf = pd.DataFrame( data = data )
print(datadf)

# target: is like C:
# data ready for shipment

from pandas import ExcelWriter
writer = ExcelWriter('Target:\\xxxxxxx\\xxxxxxx\\xxxxxxx\\xxxxxx.xlsx')
datadf.to_excel(writer,'Data L-W')
writer.save()

# dataframe to csv 
datadf.to_csv('filetoload.csv', sep=',')

#END 
