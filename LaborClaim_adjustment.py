import os
import time
import xlwings as xw
from glob import glob
import pandas as pd


folder = str(input("Enter \"YYYY Quarter#\": "))

quarter_entry = folder.split(" ")
quarter_entry = quarter_entry[1]
year_entry = folder.split(" ")
year_entry = year_entry[0]
year_entry = int(year_entry)

claims = fr"LaborClaims_amounts.csv"
df = pd.read_csv(claims, index_col=False)

path = fr"P:\PACS\Finance\Budgets\{folder}\Received\*.xlsx"

app = xw.App(visible=False)

for file in glob(path):

    file_name = os.path.basename(file)
    # Remove file extension to get the Facility name
    file_name = os.path.splitext(file_name)[0]

    # Split Facility name based on hyphen ('-')
    file_name = file_name.split('-')

    # Use the first part as the complete Facility name
    file_name = file_name[0]

    if not df.loc[df['Facility'] == file_name, 'Pro_fees'].empty:
        new_amount = df.loc[df['Facility'] == file_name, 'Pro_fees'].values[0]
        print(f"Match for {file_name}: {new_amount}")
        time.sleep(1)
        wb = xw.Book(file, update_links=False)
        time.sleep(8)
        main_page = wb.sheets("FORECAST WORKSHEET")
        insurance_line = main_page.range("C809").value
        main_page.range("C809").value = new_amount
        time.sleep(3)
        wb.save(fr'P:\PACS\Finance\Budgets\{year_entry} {quarter_entry}\Received - Adjusted\{file_name}-2024 Q1 Forecast.xlsx')
        time.sleep(2)
        wb.close()
        time.sleep(2)

    else:
        print(f"No match found for Facility: {file_name}")
