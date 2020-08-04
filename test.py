import os
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

wb = load_workbook('./08-2020 Расписание онлайн вещания (август).xlsx')
sheet = wb.get_sheet_by_name('1.08')
print(sheet)