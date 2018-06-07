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
        #print('Parsing {0}{1}\r'.format(basedir, filename))
        # Opens the xls file
        worksheet = open_workbook(basedir + filename).sheet_by_index(sheetindex)
        sheet= pd.read_excel(basedir+filename, sheet_name = sheetindex)
        all_data[sheetindex]=pd.concat([all_data[sheetindex],sheet], verify_integrity=True, ignore_index=True)
        #all_data[sheetindex]=all_data[sheetindex].append(sheet,ignore_index=True)
        #all_data[sheetindex].head()

writer = pd.ExcelWriter('verfinal14.xlsx',engine='xlsxwriter')
    #all_data[sheetindex].to_excel(writer, )
    #writer.save()
#i=['0','1','2','3']
# for j in range(0,4):
#     k = i[j]
#     print(k)
# j=0
# for key in all_data.keys():
#     print(key)
#     if j==key:
#         print("equal")
#         print(j)
#-----------------------------------------------------------
# j=0
# writer = pd.ExcelWriter('verfinal6.xlsx',engine='xlsxwriter')
# for key in df.keys():
#     all_data[j].to_excel(writer, sheet_name = key)
#     j=j+1
#     print(j)
#     print(key)
#     writer.save()
#---------------------------------------------------------------
k=0
for k in range(0,int(n)): 
    all_data[k].to_excel(writer, sheet_name = sh[k])
writer.save()
#---------------------------------------------------------------
# all_data[0].to_excel(writer, sheet_name = 'Family health plus')
# all_data[1].to_excel(writer, sheet_name = 'Managed Long Term Care')
# all_data[2].to_excel(writer, sheet_name = 'Medicaid Advantage') 
# writer.save()              
# ------------------------------------------------------------------
# for key in all_data.keys():
#     print(key)
#     print(all_data[key])