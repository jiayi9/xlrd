from functions import read_to_list, insert_to_database, list_files_recur

FILE_PATH = "//bosch.com/dfsrb/DfsCN/loc/Wx/Dept/TEF/60_MFE_Manufacturing_Engineering/06_Data_Analytics/01_Project/RBCD/cleanliness/DNOX/2.2 -SM-2.xlsx"

L = read_to_list(FILE_PATH)

insert_to_database(L)


    
file_paths, file_names = list_files_recur("C:/daten", format = '.xlsx')
