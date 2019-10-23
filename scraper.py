import urllib
import pandas as pd
import win32com.client as win32
from openpyxl import load_workbook
import os

#url, filename & path variables
url = 'insert_url_here.com'
filename = 'test_{0}.xls'
path = r'file_path'

asx = pd.read_csv('codes.csv')

#interate through each asx code and download .xls file
for i in asx['ASX Code']:
    web_path = url.format(i)
    try:
        urllib.request.urlretrieve(path, filename.format(i))
    except:
        pass

#convert each file to .xlsx and remove .xls file    
for i in asx['ASX Code']:
    fname = path.format(i)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        wb = excel.Workbooks.Open(fname)
        wb.SaveAs(fname+"x", FileFormat = 51)
        wb.Close()
        excel.Application.Quit()
        os.remove(filename.format(i))
    except:
        pass
   
sheet = 'Sheet{0}'
filename_x = 'test_{0}.xlsx'

#set-up writer
book = load_workbook('Book1.xlsx')
writer = pd.ExcelWriter('Book1.xlsx', engine='openpyxl')
writer.book = book

#separate according to financial statement type
for i in asx['ASX Code']:
    file = filename_x.format(i)
    try:
        df = pd.read_excel(file)
        sh = sheet.format(i)
        df.to_excel(writer, sh)
        writer.save()
    except:
        pass
