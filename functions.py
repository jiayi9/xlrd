import xlrd
import pyodbc
import os
from datetime import datetime

# clear log file
def clear_log_file(PATH, MAX = 50000):
    status = 0
    try:
        N = len(open(PATH).readlines())
        #print('Row count:', N, '< Limit')
    except:
        print('Reading log file error.')
        status = -1
        return(status)
    if N > MAX:
        print('Detected row count > Limit. Clean the log at ', PATH)
        file = open(PATH ,"w")
        file.truncate(0)
        file.close()
        NEW_N = len(open(PATH).readlines())
        if NEW_N < 100:
            status = 1
            print('Successfully cleared ', PATH)
        else:
            status = 2
            print('Cleared log file but the latest row count is > 100')
    else:
        status = 9
    return(status)

# write a line to a log
def append_log(CONTENT, LOG_FILE):
    with open(LOG_FILE, 'a') as file:
        file.write('\n' + CONTENT)

# output indexes of current files
def write_list_into_rows(LIST, PATH):
    with open(PATH, "w") as f:
        for s in LIST:
            f.write(str(s) +"\n")

# retrieve the list of current files
def read_rows_into_list(PATH):
    L = []
    with open(PATH, "r") as f:
        for line in f:
            L.append(line.strip())
    return(L)

# list all files under a folder recursively
def list_files_recur(path, format = '.xlsx'):
    file_paths = []
    file_names = []
    for r, d, f in os.walk(path):
        for file in f:
            if format in file:
                file_paths.append(os.path.join(r, file))
                file_names.append(file)
    return([file_paths, file_names])

# excel operation: A->1, B->2
def col_to_index(col):
    return sum((ord(c) - 64) * 26**i for i, c in enumerate(reversed(col))) - 1

# format for reprot ID
def format_datetime_1(x, wb):
    x_as_datetime = datetime(*xlrd.xldate_as_tuple(x, wb.datemode))
    return(x_as_datetime.strftime("%Y_%m_%d_%H_%M_%S"))

# format for datetime string
def format_datetime_2(x, wb):
    x_as_datetime = datetime(*xlrd.xldate_as_tuple(x, wb.datemode))
    return(x_as_datetime.strftime("%Y-%m-%d %H:%M:%S"))

# count the number of columns in result table
def get_nrow(sheet, i = 31, j = 6):
    #33
    N = 0
    NCOL = 0
    while True:
        tmp = i + N
        value = str(sheet.cell_value(tmp, j))
        if value != '':
            N = N + 2
            NCOL = NCOL + 1
        else:
            break
    return(NCOL)

# extract information from the reprot and save in a list of report, results and file path
def read_to_list(FILE_PATH, AREA = "MOE3"):
    wb = xlrd.open_workbook(FILE_PATH) 
    sheet = wb.sheet_by_index(0) 

    ## define extraction cells
    L_Reports = [
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

    for index, item in enumerate(L_Reports):
        NAME, i, EXCEL_COL_INDEX = item
        j = col_to_index(EXCEL_COL_INDEX)
        CONTENT = sheet.cell_value(i-1, j)
        L_Reports[index].append(CONTENT)
    
    ## format datetimes
    # Analysis_No
    L_Reports[0][3] = format_datetime_1(L_Reports[0][3], wb)
    # Extraction_Date
    L_Reports[10][3] = format_datetime_2(L_Reports[10][3], wb)
    # Datetime
    L_Reports[-2][3] = format_datetime_2(L_Reports[-2][3], wb)
    
    # add year, no need comment out
    #L_Reports.append(['', '', '', L_Reports[0][3][0:4]])
    #L_Reports.append(['', '', '', FILE_PATH])
    L_Reports.append(['AREA', '', '', AREA])
    L_Reports.append(['REMARK', '', '', ''])
    L_Reports.append(['FILE_NAME', '', '', FILE_PATH.split('/')[-1]])
    L_Reports.append(['FILE_PATH', '', '', FILE_PATH])



    N = get_nrow(sheet)
    COL_LABELS = [sheet.cell_value(31, col_to_index('G')+i*2) for i in range(N)]
    ROW_LABELS = ['Max', 'All', 'Length_Fibre', 'Length_Reflecting', 'Length_Others', 'Width_Fibre', 'Width_Reflecting', 'Width_Others']
    ROW_INDEX = [34, 35, 36, 37, 38, 39, 40, 41]
    L_Results = list()
    for row_label, i in zip(ROW_LABELS, ROW_INDEX):
        for col_label, j in zip(COL_LABELS, range(N)):
            value = sheet.cell_value(i - 1, col_to_index('G')+j*2)
            L_Results.append([row_label, col_label, value])
    
    L = [L_Reports, L_Results, FILE_PATH]
    return(L)

# insert the List into database
def insert_to_database(L, CONNECTION_STRING = ('Driver={SQL Server};''Server=WX1SRV41\SQLEXPRESS;''Database=CLEANLINESS;''Trusted_Connection=yes;')):
    L_Reports, L_Results, FILE_PATH = L
    conn = pyodbc.connect(CONNECTION_STRING)
    cursor = conn.cursor()
    
    # Report
    VALUES = [str(x[3]) for x in L_Reports]
    STRING = "insert into Reports values (?,?,?,?,?,?,?,?,?,?,    ?,?,?,?,?,?,?,?,?,?,    ?,convert(datetime,?),?,?,?,?,?)"
    cursor.execute(STRING, tuple(VALUES))
    conn.commit()
    
    # Result
    wb = xlrd.open_workbook(FILE_PATH) 
    sheet = wb.sheet_by_index(0) 
    N = get_nrow(sheet)
    COL_LABELS = [sheet.cell_value(31, col_to_index('G')+i*2) for i in range(N)]
    ROW_LABELS = ['Max', 'All', 'Length_Fibre', 'Length_Reflecting', 'Length_Others', 'Width_Fibre', 'Width_Reflecting', 'Width_Others']
    ROW_INDEX = [34, 35, 36, 37, 38, 39, 40, 41]
    
    L = list()
    for row_label, i in zip(ROW_LABELS, ROW_INDEX):
        for col_label, j in zip(COL_LABELS, range(N)):
            value = sheet.cell_value(i - 1, col_to_index('G')+j*2)
            L.append([row_label, col_label, value])
            
    ANALYSIS_NO = L_Reports[0][3]
    STANDARD = L_Reports[3][3]
    
    STRING = "insert into results values (?, ?, ?, ?, ?);"
    
    for index, item in enumerate(L):
        cursor.execute(STRING, ANALYSIS_NO, STANDARD, item[0], item[1], item[2])
        conn.commit()
        
    conn.close()

#FILE_PATH = "//bosch.com/dfsrb/DfsCN/loc/Wx/Dept/TEF/60_MFE_Manufacturing_Engineering/06_Data_Analytics/01_Project/RBCD/cleanliness/DNOX/2.2 -SM-2.xlsx"
#L = read_to_list(FILE_PATH)
#insert_to_database(L)

#FILE_PATH = "D:/Share_Tableau_R/Cleanliness/data/MOE3/2017.10.16-2/20171016_161618_852-Cleanalyzer2P.xlsx"
#L = read_to_list(FILE_PATH)
#insert_to_database(L)
