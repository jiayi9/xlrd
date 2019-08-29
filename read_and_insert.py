
import xlrd
import pandas as pd
import numpy as np
from datetime import datetime

loc = ("//bosch.com/dfsrb/DfsCN/loc/Wx/Dept/TEF/60_MFE_Manufacturing_Engineering/06_Data_Analytics/01_Project/RBCD/cleanliness/DNOX/2.2 -SM-2.xlsx") 
#loc = ("/Users/tef-itm/Documents/cleanliness/xlrd/2.2 -SM-2.xlsx") 

wb = xlrd.open_workbook(loc) 

sheet = wb.sheet_by_index(0) 
  
sheet.cell_value(0,0)

def col_to_index(col):
    return sum((ord(c) - 64) * 26**i for i, c in enumerate(reversed(col))) - 1

def format_datetime_1(x):
    x_as_datetime = datetime(*xlrd.xldate_as_tuple(x, wb.datemode))
    return(x_as_datetime.strftime("%Y_%m_%d_%H_%M_%S"))

def format_datetime_2(x):
    x_as_datetime = datetime(*xlrd.xldate_as_tuple(x, wb.datemode))
    return(x_as_datetime.strftime("%Y-%m-%d %H:%M:%S"))



Dictionary = [
# General information
 ["Analysis_No", 7, "N"],
 ["Extraction_Device", 8, "N"],
 ["Test_Specification", 9, "N"],
# Test speciman Left
 ["Standard", 13, "G"],
 ["Component_No", 14, "G"],
 ["Batch_No", 15, "G"],
 ["Media_Sample", 16, "G"],
 ["Others", 17, "G"],
# Test speciman Right 
 ["Sample_No", 13, "T"],
 ["Shipping_Package", 14, "T"],
 ["Extraction_Date", 15, "T"],
 ["Filtering_Drying", 16, "T"],
# Light microscope analysis
 ["Microscope", 21, "I"], 
 ["Software_Version", 21, "U"],
 ["Calibration", 22, "U"],
 ["Filter_Size", 23, "U"],
 ["Analyze_Size", 24, "U"],
 ["Threshold_Level", 25, "U"],
 ["Population_Density", 26,"U"],
# Result
 ["Result", 44, "H"],
 ["Inspector", 44, "M"],
 ["Datetime", 44, "V"],
 ["Comment", 47, "H"]
]

for index, item in enumerate(Dictionary):
    NAME, i, EXCEL_COL_INDEX = item
    j = col_to_index(EXCEL_COL_INDEX)
    CONTENT = sheet.cell_value(i-1, j)
    Dictionary[index].append(CONTENT)
#    print(NAME, i,j, CONTENT)

pd.DataFrame(np.array(Dictionary),columns=['ProductType',"x","u",'xx'])

Dictionary

Dictionary[0][3] = format_datetime_1(Dictionary[0][3])

Dictionary[10][3] = format_datetime_2(Dictionary[10][3])

Dictionary[-2][3] = format_datetime_2(Dictionary[-2][3])



import pyodbc
conn = pyodbc.connect(
        'Driver={SQL Server};'
        'Server=SGHZ001020351\SQLEXPRESS;'
        'Database=CLEANLINESS;'
        'Trusted_Connection=yes;'
        )


cursor = conn.cursor()
cursor.execute("select * from persons")

for row in cursor:
    print(row)

cursor = conn.cursor()

VALUES = [str(x[3]) for x in Dictionary]

sep = "','"
VALUE_sep = sep.join(VALUES)
STRING = "insert into REPORTS values ('" + VALUE_sep +"')"

cursor.execute(STRING)

conn.commit()





def get_nrow(sheet, i = 33, j = 6):
    N = 0
    while True:
        i = 33 + N
        value = str(sheet.cell_value(i, 6))
        print(value)
        if value != '':
            N = N + 1
        else:
            break
    return(N)

N = get_nrow(sheet)


#for i in range(N):
#    print(sheet.cell_value(31, col_to_index('G')+i*2))

COL_LABELS = [sheet.cell_value(31, col_to_index('G')+i*2) for i in range(N)]

ROW_LABELS = ['Max', 'All', 'Length_Fibre', 'Length_Reflecting', 'Length_Others', 'Width_Fibre', 'Width_Reflecting', 'Width_Others']
ROW_INDEX = [34, 35, 36, 37, 38, 39, 40, 41]




L = list()
for row_label, i in zip(ROW_LABELS, ROW_INDEX):
    for col_label, j in zip(COL_LABELS, range(N)):
        
        value = sheet.cell_value(i - 1, col_to_index('G')+j*2)
        print(i, j,row_label, col_label, value)
        L.append([row_label, col_label, value])
        
ANALYSIS_NO = Dictionary[0][3]
STANDARD = Dictionary[3][3]

STRING = "insert into results values (?, ?, ?, ?, ?);"

for index, item in enumerate(L):
    cursor.execute(STRING, ANALYSIS_NO, STANDARD, item[0], item[1], item[2])
    conn.commit()



