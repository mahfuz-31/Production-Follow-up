from bs4 import BeautifulSoup
import pandas as pd # type: ignore
import math
from openpyxl import load_workbook # type: ignore
from openpyxl.styles import Border, PatternFill, Side # type: ignore
import os
from openpyxl.styles import Font # type: ignore
from openpyxl.styles import Alignment
from datetime import datetime, timedelta
def get_confirmation():
    choice = input("Do you want to create yesterday's report? (Press 'N' or 'n' for No, any other key for Yes): ").strip()
    return choice.lower() != 'n'

def setPrecision(value, decimal_places):
    factor = 10 ** decimal_places
    return math.floor(value * factor) / factor
    
# data = pd.read_csv('data.csv')
file_name = 'diagnostic_summary.html'
with open(file_name, "r", encoding="latin1") as file:
    html_file = file.read()

html_content = BeautifulSoup(html_file, "html.parser")
rows = []
for row in html_content.find_all('tr'):
    row_data = [cell.get_text(strip=True) for cell in row.find_all('td')]
    rows.append(row_data)

today = ''
if get_confirmation():
    today = datetime.now() - timedelta(days=1)
    today = today.strftime("%d-%b-%y")  # Example: 10-Mar-25
else:
    date = int(input("Enter how many days ago from today's date you want to get: "))
    date = datetime.now() - timedelta(days=date)
    today = date.strftime("%d-%b-%y")
file_name = "07. Jul/" + today + ".xlsx"
file2_name = '//192.168.1.231/Planning Internal/Md. Mahfuzur Rahman/Production follow up/07. Jul/' + str(today) + '.xlsx'

units = ['JAL', 'JAL3', 'JFL', 'JKL', 'MFL', 'FFL2', 'JKL-U2', 'GTAL', 'GMT TOTAL:', 'LINGERIE']

file_count = sum(1 for file in os.listdir('07. Jul/') if os.path.isfile(os.path.join('07. Jul/', file)))

completed_days = int(file_count) + 1
print("completed_days: ", completed_days)

wb = load_workbook("template.xlsx")
ws = wb["Sheet1"]
i = 0
for row in rows:
    if len(row) == 20 and row[0] in units:
        ws['B' + str(i + 4)] = float(row[7].replace(',', ''))
        ws['D' + str(i + 4)] = float(row[13].replace(',', ''))
        ws['K' + str(i + 4)] = float(row[8].replace(',', ''))
        ws['O' + str(i + 4)] = float(row[18])
        ws['Q' + str(i + 4)] = float(row[14])
        ws['S' + str(i + 4)] = float(row[15].replace('%', ''))/100
        ws['W' + str(i + 4)] = completed_days
        i += 1
ws['B2'] = today  # date_cell
wb.save(file_name)
wb.save(file2_name)
wb.close()
if os.path.exists(file_name):
    print("File created successfully.")