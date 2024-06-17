# -*- coding: utf-8 -*-

import pandas as pd
import classGrades
#from win32com.client.dynamic import Dispatch

sem = 'Fall'
year = '2019'
#className = 'MTH113'
className = 'MTH120'
classTime = '400'

inPath = "C:/Users/Seaver-AK/Desktop/PII_AMRDEC/Calhoun/"
inFileName = "hwCh3.csv"
hwPath = inPath+sem+year+'/'+className+'-'+classTime+'/private/hw/'

sectionNum = '400'
weekDays = 'TTH'
rosterDateStr = '10172019'
xlPath = inPath+sem+year+'/'+className+'-'+classTime+'/private/rosters/'
rosterFile = classTime+'roster'+rosterDateStr+'.xlsx'
xlFile = xlPath+rosterFile
classData = classGrades.ClassGrades(className,sectionNum,classTime,weekDays,\
                                    year,sem,xlFile)

hwFile = hwPath+inFileName
outFileName = "hwCh3.xlsx"

## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
## ~~~~~~~~~~ Add appropriate borders to used cells in given XL sheet ~~~~~~~~
## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

def addBorders(sheet):
  usedCells = sheet.UsedRange.Value
  rows = len(usedCells)
  columns = len(usedCells[0])

  # Get xl address of 1st and last cell within the range of cells used
  first = sheet.Cells(1,1).Address
  last = sheet.Cells(rows,columns).Address
  # Set range of xl cells
  myRange = sheet.Range(first+":"+last)
  myRange.Font.Size = 10

  # Define the borders to be used within the range
  borders = [7,8,9,10,11,12]

  # Set each border for the given range
  for border in borders:
    b = myRange.Borders(border)
    # Define the border attributes
    b.LineStyle = 1
    b.Weight = 2
    b.ColorIndex = 1

## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
def autoFitColumns(sheet,data):
    for idx, col in enumerate(data):
        series = data[col]
        maxLen = max((series.astype(str).map(len).max(),len(str(series.name))))+1
        sheet.set_column(idx,idx,maxLen)

## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

def printInXL(data,hwPath,outFileName):
  writer = pd.ExcelWriter(hwPath+outFileName, engine='xlsxwriter')
  data.to_excel(writer,sheet_name='HW',index=None,\
                float_format='%.1f')
  wb = writer.book
  sheet = writer.sheets['HW']
  
  nameFormat = wb.add_format({'bold':True, 'border':True})
  sheet.set_column('A:A',len(data),nameFormat)
  
  autoFitColumns(sheet,data)
  writer.save()
  #sheet.Columns.AutoFit()
  #addBorders(sheet)
  #sheet.Rows.AutoFit()

## ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
  
data = pd.read_csv(hwFile,skiprows=[1,2,3,4])
data = data.iloc[:-7,:-1]
data = data.drop(columns=[data.columns[2],data.columns[3],data.columns[4]])
data = data.rename(columns={data.columns[0]:'Last Name',\
                            data.columns[1]:'First Name'})

data['Last Name'] = data['Last Name'].str.title()
data['First Name'] = data['First Name'].str.title()
data = data.fillna(0.0)
data['Name'] = data['Last Name'] + ', ' + data['First Name']
cols = data.columns.to_list()
cols = cols[-1:] + cols[:-1]
data = data[cols]
for student in classData.students:
  matchFlag = False
  rosterLastName = student.name.split(', ')[0]
  rosterFirstName = student.name.split(', ')[1].split(' ')[0]
  #print(rosterLastName,end=' ')
  if rosterLastName in list(data['Last Name']):
    matches = [i for i,e in enumerate(list(data['Last Name'])) if \
                e == rosterLastName]
    matchFlag = True
    if rosterFirstName in list(data['First Name'][matches]):
       idx = list(data['First Name']).index(rosterFirstName)
    else:
       data['First Name'].loc[matches[0]] = rosterFirstName
       
  if not matchFlag:
      print(student.name)
        #print([student.name,data['Last Name'][idx]],data['First Name'][idx])

for idx,colName in enumerate(data.columns):
    if 'Homework' in colName:
        data.rename(columns={colName:colName.split(' ')[0]},inplace=True)
        
#for student in classData.students:
#  for student2 in data['Last Name']:
#    if student2==student.name.split(', ')[0]:
#        print([student.name.split(', ')[0],\
#           student.name.split(', ')[1].split(' ')[0],\
#           student2])
#del data['Last Name']
#del data['First Name']

ch = '3'

ch = ch+'.'
cols = [x for x in data.columns if '5.' in x]
data['Bonus'] = -3

colMean = data[cols].mean(axis=1)
colMin = data[cols].min(axis=1)
for idx,line in enumerate(colMin):
    if line > 98:
        data['Bonus'][idx] = 5
    elif line > 40:
        data['Bonus'][idx] = 0
    elif colMean[idx] > 50:
        data['Bonus'][idx] = 0
        #print([data['Name'][idx:idx+1],data[cols][idx:idx+1]])
#data['Bonus']
data['Mean'] = data.mean(axis=1)
printInXL(data,hwPath,outFileName)




#data = pd.read_csv(inFile,skiprows=4)


#classInfo = dict()
#  ## Open and retrieve list if exists
#  if os.path.isfile(inFile):
#    with open(inFile,'r') as fin:
#      reader = csv.reader(fin,delimiter=',')
#      #next(reader,None)
#      for row in reader:
#        classInfo[row[0]]=row[1:]