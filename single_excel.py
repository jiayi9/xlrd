
import xlrd
import pandas as pd
import numpy as np
from datetime import datetime

loc = ("/Users/tef-itm/Documents/cleanliness/xlrd/2.2 -SM-2.xlsx") 

wb = xlrd.open_workbook(loc) 

sheet = wb.sheet_by_index(0) 
  
sheet.cell_value(0,0)

def col_to_index(col):
    return sum((ord(c) - 64) * 26**i for i, c in enumerate(reversed(col))) - 1

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

for item in Dictionary:
    NAME, i, EXCEL_COL_INDEX = item
    j = col_to_index(EXCEL_COL_INDEX)
    CONTENT = sheet.cell_value(i-1, j)
    print(NAME, i,j, CONTENT)

pd.DataFrame(np.array(Dictionary),columns=['ProductType',"x","u"])





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


for i in range(N):
    print(sheet.cell_value(31, col_to_index('G')+i*2))

COL_LABELS = [sheet.cell_value(31, col_to_index('G')+i*2) for i in range(N)]

ROW_LABELS = ['Max', 'All', 'Length_Fibre', 'Length_Reflecting', 'Length_Others', 'Width_Fibre', 'Width_Reflecting', 'Width_Others']
ROW_INDEX = [34, 35, 36, 37, 38, 39, 40, 41]

for i, j in zip(LABELS, ROW_INDEX):
    print(i,j)


for row_label, i in zip(ROW_LABELS, ROW_INDEX):
    for col_label, j in zip(COL_LABELS, range(N)):
        
        value = sheet.cell_value(i - 1, col_to_index('G')+j*2)
        print(i, j,row_label, col_label, value)

def format_datetime(x):
    x_as_datetime = datetime(*xlrd.xldate_as_tuple(x, wb.datemode))
    return(x_as_datetime.strftime("%Y_%m_%d_%H_%M_%S"))


format_datetime(x)







sheet.cell_value(6, col_to_index('N'))

sheet.cell_value(14, col_to_index('N'))

sheet.cell_value(14, col_to_index('T'))




format_date()
    
    

from datetime import datetime
x = sheet.cell_value(14, col_to_index('T'))
a1_as_datetime = datetime(*xlrd.xldate_as_tuple(x, wb.datemode))
a1_as_datetime.strftime("%Y_%m_%d_%H_%M_%S")






py_date = datetime.datetime(*xlrd.xldate_as_tuple(a[0].value,
                                                  book.datemode))



a1_as_datetime.strftime("%m/%d/%Y, %H:%M:%S")


import datetime




print(sheet.nrows)

for i in range(sheet.ncols):
    print(i)

for i in range(sheet.nrows):
    print(i)
    
    
    
    
    
    
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
 
df = pd.read_excel(loc)
 
print("Column headings:")
print(df.columns)



xl = pd.ExcelFile(loc)
xl.sheet_names

df_Report = pd.read_excel(loc, sheetname = 'Report')

df_Images =  pd.read_excel(loc, sheetname = 'Images')







from PIL import ImageGrab
import win32com.client as win32
import numpy as np

loc = ("//bosch.com/dfsrb/DfsCN/loc/Wx/Dept/TEF/60_MFE_Manufacturing_Engineering/06_Data_Analytics/01_Project/RBCD/cleanliness/MOE6_Cleanliness report with picture/20180511_170114_492-Cleanalyzer2P-1#Line(NOK).xlsx") 

pic_loc = "//bosch.com/dfsrb/DfsCN/loc/Wx/Dept/TEF/60_MFE_Manufacturing_Engineering/06_Data_Analytics/01_Project/RBCD/cleanliness/MOE6_Cleanliness report with picture/"

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(loc)

SHAPE = None

for sheet in workbook.Worksheets:
    for i, shape in enumerate(sheet.Shapes):
        if not shape.Name.startswith('Picture'):
            shape.Copy()
            image = ImageGrab.grabclipboard()
            PATH  = pic_loc+ '/' + str(i)+'.jpeg'
            
#            image.save('{}.jpg'.format(i+1), 'jpeg')
            if not np.all(np.array(image).flatten() == 255):
                image.save(PATH, 'jpeg')


workbook.Close(False)

import numpy as np
pix = np.array(SHAPE)
pix.flatten()
pix.flatten() == 255
np.all(pix.flatten() == 255)


for sheet in workbook.Worksheets:
    for i, shape in enumerate(sheet.Shapes):
        print(shape.Name)


# https://www.penwatch.net/cms/images_from_excel/

import win32com.client       # Need pywin32 from pip
from PIL import ImageGrab    # Need PIL as well
import os

excel = win32com.client.Dispatch("Excel.Application")
workbook = excel.ActiveWorkbook


wb_folder = 'C:/daten/'
wb_name = workbook.Name
wb_path = os.path.join(wb_folder, wb_name)

wb_path = loc

print("Extracting images from ", wb_path)

image_no = 0
SHAPE = None

for sheet in workbook.Worksheets:
    print(sheet)
    
    for n, shape in enumerate(sheet.Shapes):
        if shape.Name.startswith("Picture"):
            # Some debug output for console
            image_no += 1
            print( "---- Image No. %07i ----")

            # Sequence number the pictures, if there's more than one
            num = "" if n == 0 else "_%03i" % n

            filename = sheet.Name + num + ".jpg"
            file_path = os.path.join (wb_folder, filename)

            print ("Saving as", file_path)    # Debug output

            shape.Copy() # Copies from Excel to Windows clipboard

            # Use PIL (python imaging library) to save from Windows clipboard
            # to a file
            image = ImageGrab.grabclipboard()
            image.save(file_path,'jpeg')
            if i == 5:
                SHAPE = image
