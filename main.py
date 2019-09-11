
from functions import read_to_list, insert_to_database, list_files_recur, write_list_into_rows, read_rows_into_list, append_log, clear_log_file
import os
import datetime

ENV = "D:/Share_Tableau_R/Cleanliness/"
DATA_SOURCE = "common/data_source.txt"
os.chdir(ENV)
SOURCE = eval(open(DATA_SOURCE, 'r').read())
NOW = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
DO_LIST = ['MOE3']

for i, dept in enumerate(DO_LIST):
    #print(i, dept)
    
    LOG_FILE_PATH = "log/log_" + dept + ".txt"
    
    clear_log_file(LOG_FILE_PATH, MAX = 50000)
    
    INDEXES_FILE_PATH = "indexes/" + dept + "_indexes.txt"
    
    DATA_FILE_PATH = "data/" + dept + "/"
    
    EXISTING_REPORTS = set(read_rows_into_list(INDEXES_FILE_PATH))
    
    CURRENT_REPORTS = set(list_files_recur(DATA_FILE_PATH, format = ".xlsx")[0])
    
    #print(len(EXISTING_REPORTS), len(CURRENT_REPORTS))
    append_log(NOW, LOG_FILE_PATH)
    append_log("  Existing files in indexes: " + str(len(EXISTING_REPORTS)) + "   Current files in Folder: " + str(len(CURRENT_REPORTS)), LOG_FILE_PATH)
    
    if len(EXISTING_REPORTS) != len(CURRENT_REPORTS):
        
        # x.difference(y)   elements in x but not in y        
        NEW_REPORTS = CURRENT_REPORTS.difference(EXISTING_REPORTS)        
#        print("There are ", len(NEW_REPORTS))
        append_log("  There are " + str(len(NEW_REPORTS)) + " new files", LOG_FILE_PATH)
        
        tmp = 0
        
        for j, report_path in enumerate(list(NEW_REPORTS)):
            
            if j % 2000 == 0 & j > 1999:
                print(dept,j)
            
            # write to database
            try:
                L = read_to_list(report_path)
                insert_to_database(L)
            except:
                append_log("    " + report_path, LOG_FILE_PATH) 
                tmp = tmp + 1
                #print(report_path)
            # update the indexes
        append_log("  Problematic files: " + str(tmp) ,LOG_FILE_PATH)
        
        write_list_into_rows(CURRENT_REPORTS, INDEXES_FILE_PATH)
        append_log("----------------------------------------" ,LOG_FILE_PATH)
        
        
    else:
        #print("  No update at", datetime.datetime.now())
        append_log("  No update", LOG_FILE_PATH)
        append_log("----------------------------------------" ,LOG_FILE_PATH)

