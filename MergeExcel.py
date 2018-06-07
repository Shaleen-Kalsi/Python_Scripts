import re
import csv
from xlrd import open_workbook
from os import listdir
import pandas as pd
import xlsxwriter
# Empty dictionary where we will store each row, before parsing it to the CSV
row = {}
##To merge excel files
# This is where the xls files are
basedir = 'data/'

# This is an os function that returns a list of filenames in a folder
files = listdir('data')

# Empty list to store only XLS files found in the folder
books = [filename for filename in files if filename.endswith("xls")]
filename = "en01_13.xls"
worksheet = open_workbook(basedir + filename)
df = pd.read_excel(basedir+filename, None)
n = input("enter no. of sheets to be merged")
str = input("enter comma separated list of sheet names")
##str = 'Family health plus,Managed Long Term Care,Medicaid Advantage'
sh = str.split(',')
header_is_written = False
# Iterating over the files in folder
all_data = {}
for sheetindex in range(0,int(n)):
    all_data[sheetindex] = pd.DataFrame()
    for filename in books:
        worksheet = open_workbook(basedir + filename).sheet_by_index(sheetindex)
        sheet= pd.read_excel(basedir+filename, sheet_name = sheetindex)
        all_data[sheetindex]=pd.concat([all_data[sheetindex],sheet], verify_integrity=True, ignore_index=True)        
writer = pd.ExcelWriter('verfinal14.xlsx',engine='xlsxwriter')
k=0
for k in range(0,int(n)): 
    all_data[k].to_excel(writer, sheet_name = sh[k])
writer.save()
