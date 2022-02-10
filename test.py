from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
wb2 = load_workbook('13F-data.xlsx')


ws = wb2["Stocks By Quarters"]


print(ws['B1'].value)