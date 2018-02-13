import openpyxl
import win32ui
import os

print("\n")
print('Please select a TXT file.\nStart selecting after hitting enter.')
input()
dlg = win32ui.CreateFileDialog(1) 
dlg.SetOFNInitialDir('c:/') 
dlg.DoModal() 
filename = dlg.GetPathName() 
print ("Please confirm the filename:\n"+filename+"\n")

file=open(filename)
content=file.read()
print("The contents of the TXT file are:")
print(content)
rows=content.split('\n')
num=len(rows)

print("\n")
print('Please select a excel file.\nStart selecting after hitting enter.')
input()
dlg = win32ui.CreateFileDialog(1) 
dlg.SetOFNInitialDir('c:/') 
dlg.DoModal() 
filename = dlg.GetPathName() 
print ("Please confirm the filename:\n"+filename+"\n")

workbook=openpyxl.load_workbook(filename)
tlist=workbook.get_sheet_names()
sheet=workbook.get_active_sheet()

a=int(input("Please enter the starting row: "))
b=int(input("Please enter the starting column: "))

for i in range(num):
	temp=rows[i].split()
	#print(temp)
	for j in range(len(temp)):
		sheet.cell(row=i+a,column=j+b).value=temp[j]
workbook.save(filename)
print("\nCompleted\n")
input("Press any key to exit.")
