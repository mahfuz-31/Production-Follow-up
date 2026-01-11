from bs4 import BeautifulSoup
import math
from openpyxl import load_workbook # type: ignore
from openpyxl.styles import Border, PatternFill, Side # type: ignore
import os
from openpyxl.styles import Font # type: ignore
from openpyxl.styles import Alignment
from datetime import datetime, timedelta
import pandas as pd

def get_confirmation():
    choice = input("Do you want to create yesterday's report? (Press 'N' or 'n' for No, any other key for Yes): ").strip()
    return choice.lower() != 'n'

def setPrecision(value, decimal_places):
    factor = 10 ** decimal_places
    return math.floor(value * factor) / factor
    
# data = pd.read_csv('Report.csv', encoding='ISO-8859-1')
with open("Sewing Diagnostic Summary Report.html", "r", encoding="utf-8") as file:
    html_file = file.read()
html_content = BeautifulSoup(html_file, "html.parser")
rows = []
for row in html_content.find_all('tr'):
    row_data = [cell.get_text(strip=True) for cell in row.find_all('td')]
    if len(row_data) != 20:
        continue
    rows.append(row_data)
today = ''
if get_confirmation():
    today = datetime.now() - timedelta(days=1)
    today = today.strftime("%d-%b-%y")  # Example: 10-Mar-25
else:
    date = int(input("Enter how many days ago from today's date you want to get: "))
    date = datetime.now() - timedelta(days=date)
    today = date.strftime("%d-%b-%y")
file_name = "01. Jan/" + today + ".xlsx"
file2_name = '//192.168.1.231/Planning Internal/Md. Mahfuzur Rahman/Production follow up/01. Jan/' + str(today) + '.xlsx'

units = ['SubTotal: JAL', 'SubTotal: JFL', 'SubTotal: JKL', 'SubTotal: MFL', 'SubTotal: FFL2', 'SubTotal: JKL-U2', 'SubTotal: LIN', 'SubTotal: GTAL']

file_count = sum(1 for file in os.listdir('01. Jan/') if os.path.isfile(os.path.join('01. Jan/', file)))

completed_days = int(file_count) + 1
print("completed_days: ", completed_days)

wb = load_workbook("template.xlsx")
ws = wb["Sheet1"]
i = 0
for row in rows:
    if row[1] in units:
        if row[1] == 'SubTotal: LIN':
            ws['B12'] = int(row[8].replace(',', ''))
            ws['D12'] = int(row[14].replace(',', ''))
            ws['K12'] = int(row[9].replace(',', ''))
            ws['O12'] = float(row[19])
            ws['Q12'] = float(row[15])
            ws['S12'] = float(row[16].replace('%', ''))/100
            ws['W12'] = completed_days
        else:
            ws['B' + str(i + 4)] = int(row[8].replace(',', ''))
            ws['D' + str(i + 4)] = int(row[14].replace(',', ''))
            ws['K' + str(i + 4)] = int(row[9].replace(',', ''))
            ws['O' + str(i + 4)] = float(row[19])
            ws['Q' + str(i + 4)] = float(row[15])
            ws['S' + str(i + 4)] = float(row[16].replace('%', ''))/100
            ws['W' + str(i + 4)] = completed_days
            i += 1

ws['B2'] = today
wb.save(file_name)
wb.save(file2_name)
wb.close()
if os.path.exists(file_name):
    print("File created successfully.")