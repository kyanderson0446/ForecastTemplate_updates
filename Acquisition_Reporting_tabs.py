import re
import os
import duckdb
import pandas as pd
from glob import glob

acquisition_group = input("Enter the acquisition group name folder: ")
general_path = rf"P:\Finance\Acquisitions\{acquisition_group}\REPORTING_COMPILE"

def create_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Folder '{folder_path}' created.")
    else:
        print(f"Folder '{folder_path}' already exists.")

# create_folder(general_path)

master_xl_file = fr"{general_path}\{acquisition_group}.xlsx"

# Connect to DuckDB
conn = duckdb.connect(database=':memory:', read_only=False)
conn.execute('INSTALL spatial;')
conn.execute('LOAD spatial;')

# Accumulate DataFrames to be written to Excel
dfs_to_write = []

for file in glob(fr"{general_path}\*.xlsx"):
    file_name = os.path.splitext(os.path.basename(file))[0]
    query_str = "SELECT * FROM st_read(?, layer='REPORTING');"
    df = conn.execute(query_str, [file]).fetchdf()
    facility_name = df.iloc[0, 0]
    facility_name = re.sub(r'[^\w\s]', '', facility_name)
    df.columns = [col.lower() for col in df.columns]  # Ensure lowercase column names
    dfs_to_write.append((facility_name, df))
    print(file, "added to Excel file.")

# Write all accumulated DataFrames to Excel
with pd.ExcelWriter(master_xl_file, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
    for facility_name, df in dfs_to_write:
        df.to_excel(writer, sheet_name=facility_name, index=False)

print("All files added to Excel file.")




#####################################################################################

# import re
# import os
# import duckdb
# import pandas as pd
# from glob import glob
#
#
# acquisition_group = str(input("Enter the acquisition group name folder: "))
#
# general_path = fr"P:\Finance\Acquisitions\{acquisition_group}\REPORTING_COMPILE")
#
# def create_folder(folder_path):
#     if not os.path.exists(folder_path):
#         os.makedirs(folder_path)
#         print(f"Folder '{folder_path}' created.")
#     else:
#         print(f"Folder '{folder_path}' already exists.")
#
# # Usage
# # create_folder(fr'{general_path}')
#
# master_xl_file = fr'{general_path}\{acquisition_group}.xlsx'
#
#
# conn = duckdb.connect(database=':memory:', read_only=False)
# conn.execute('INSTALL spatial;')
# conn.execute('LOAD spatial;')
#
# for file in glob(fr"{general_path}\*.xlsx"):
#     file = file.strip('.xlsx')
#     query_str = fr"""SELECT * FROM st_read('{file}', layer='REPORTING');"""
#     df = conn.query(query_str).df()
#     facility_name = df.iloc[0, 0]
#     facility_name = re.sub(r'[^\w\s]', '', facility_name)
#     with pd.ExcelWriter(master_xl_file, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
#         df.to_excel(writer, sheet_name=facility_name, index=False)
#     print(file, "added to Excel file.")
