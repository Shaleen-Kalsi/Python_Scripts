import xml.etree.ElementTree as ET
import csv
import pandas as pd
import xlsxwriter
from openpyxl import Workbook

tree = ET.parse("en01_13_test.xml")
root = tree.getroot()
# <ITEM><County>Albany</County><Plan Name>TOTALS:</Plan Name>
# <A Enrolled>21862</A Enrolled><B Enrolled>4997</B Enrolled>
# <C Enrolled>26859</C Enrolled><D Enrolled>3757</D Enrolled>
# <TOTAL ENROLLED>30616</TOTAL ENROLLED></ITEM>
# open a file for writing

Item_data = open('Ans2.csv', 'w')

# create the csv writer object
book=Workbook()
#book.create_sheet("MedicaidManagedCare")
MedicaidManagedCare = book.active
MedicaidManagedCare.title = 'MedicaidManagedCare'
writer = pd.ExcelWriter('Ans2.xlsx',engine='xlsxwriter')
csvwriter = csv.writer(Item_data)
item_head = []


count = 0
for member in root.findall('ITEM'):
	item = []
	if count == 0:
		County = member.find('County').tag
		item_head.append(County)
		PlanName = member.find('PlanName').tag
		item_head.append(PlanName)
		AEnrolled = member.find('AEnrolled').tag
		item_head.append(AEnrolled)
		BEnrolled = member.find('BEnrolled').tag
		item_head.append(BEnrolled)
		CEnrolled = member.find('CEnrolled').tag
		item_head.append(CEnrolled)
		DEnrolled = member.find('DEnrolled').tag
		item_head.append(DEnrolled)
		TOTALENROLLED = member.find('TOTALENROLLED').tag
		item_head.append(TOTALENROLLED)	
		df1 = pd.DataFrame(item_head)
		#df1.T.to_excel(writer, sheet_name = 'Medicaid Managed Care')
		#writer.save() 	
		MedicaidManagedCare.append(item_head)
		book.save('Anstry4.xlsx')
		csvwriter.writerow(item_head)
		count = count + 1

	County = member.find('County').text
	item.append(County)
	PlanName = member.find('PlanName').text
	item.append(PlanName)
	AEnrolled = member.find('AEnrolled').text
	item.append(AEnrolled)
	BEnrolled = member.find('BEnrolled').text
	item.append(BEnrolled)
	CEnrolled = member.find('CEnrolled').text
	item.append(CEnrolled)
	DEnrolled = member.find('DEnrolled').text
	item.append(DEnrolled)	
	TOTALENROLLED = member.find('TOTALENROLLED').text
	item.append(TOTALENROLLED)	
	#df2 = pd.DataFrame(item)
	#df2.T.to_excel(writer, sheet_name = 'Medicaid Managed Care')
	#writer.save()
	MedicaidManagedCare.append(item)
	book.save('Anstry4.xlsx')
	csvwriter.writerow(item)
Item_data.close()
