import re
import os
import time

import duckdb
import pandas as pd
from glob import glob
import openpyxl

print("-"*20)
print("Case sensitive, Please Match the Exact Folder Name")
print("-"*20)
acquisition_group = str(input("Enter the acquisition group name folder: "))
acquisition_type = str(input("Enter the acquisition type\n Active, Closing or Completed Buildings: "))
print("-"*20)
general_path = fr"P:\Finance\Acquistions & New Build\{acquisition_type}\{acquisition_group}\Proformas\*.xlsx"
save_path = fr"P:\Finance\Acquistions & New Build\REPORTING_COMPILE"

#######################################################################

def create_excel_if_not_exists(save_path, acquisition_group):
    master_xl_file = fr"{save_path}\{acquisition_group}-Consolidated_REPORTING.xlsx"
    if not os.path.isfile(master_xl_file):
        wb = openpyxl.Workbook()
        wb.save(master_xl_file)
        wb.close()
        time.sleep(1)
        print(f"File {master_xl_file} created successfully.")
    else:
        print(f"File {master_xl_file} already exists.")


create_excel_if_not_exists(save_path, acquisition_group)


master_xl_file = fr"{save_path}\{acquisition_group}-Consolidated_REPORTING.xlsx"

conn = duckdb.connect(database=':memory:', read_only=False)
conn.execute('INSTALL spatial;')
conn.execute('LOAD spatial;')


for file in glob(general_path):
    print(file)
    query_str = fr"""SELECT * FROM st_read('{file}', layer='REPORTING');"""
    df = conn.query(query_str).df()
    facility_name = df.iloc[0, 0]
    facility_name = re.sub(r'[^\w\s]', '', facility_name)
    with pd.ExcelWriter(master_xl_file, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        df.to_excel(writer, sheet_name=facility_name, index=False)

print()
print(fr'{acquisition_group} has been loaded')


