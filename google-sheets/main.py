import pandas as pd

import gspread
from oauth2client.service_account import ServiceAccountCredentials

#15099
# Подключение к Google Sheets
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('sheets-python-406315-bfb38375c84e.json', scope)
client = gspread.authorize(creds)
spreadsheet_id = '14oObq8VbRAPGWjgmNvQ0CsAOO2LhRlmFd8GR30c-Gyo'
sheet = client.open_by_key(spreadsheet_id).sheet1

# Чтение данных из файла Excel в DataFrame
excel_file_path = './Free_Test_Data_500KB_XLSX.xlsx'
df = pd.read_excel(excel_file_path)
print(df)
# Перевод данных в формат списка списков 
values = df.values.tolist()

# Определение последней заполненной строки
last_filled_row = len(sheet.get_all_values()) + 1

# Вставка данных в Google Sheets
sheet.insert_rows(values, last_filled_row)

print("Успех, данные вставлены в Google Sheets")
