import xml.etree.ElementTree as ET
import csv
import pandas as pd
import xlsxwriter
from openpyxl import Workbook
##To convert .xml to excel file
tree = ET.parse("en01_13_Try.xml")
root = tree.getroot()
#Example .xml format
# <ITEM><County>Albany</County><Plan Name>TOTALS:</Plan Name>
# <A Enrolled>21862</A Enrolled><B Enrolled>4997</B Enrolled>
# <C Enrolled>26859</C Enrolled><D Enrolled>3757</D Enrolled>
# <TOTAL ENROLLED>30616</TOTAL ENROLLED></ITEM>
# open a file for writing
#str='MedicaidMangagedCare,County,PlanName,AEnrolled,BEnrolled,CEnrolled,DEnrolled,TOTALENROLLED,TRY'
str = input("Enter the sheet title and tags in comma separated form")
tags=str.split(',')
print(tags)
book=Workbook()
Sheet = book.active
Sheet.title = tags[0]

item_head = []
count = 0
N = tags.__len__()
for member in root.findall('ITEM'):
	item = []
	if count == 0:
		i=1
		while i<N :
			Tag = member.find(tags[i]).tag
			item_head.append(Tag)
			i=i+1
		df1 = pd.DataFrame(item_head) 	
		Sheet.append(item_head)
		book.save('FinalTry4.xlsx')
		count = count + 1
		print(item_head)
	i=1
	while i<N:
		Data = member.find(tags[i]).text
		item.append(Data)
		i=i+1
	Sheet.append(item)
	book.save('xml2excel.xlsx')
