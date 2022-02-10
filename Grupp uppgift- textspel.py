import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('C:/Users/MichaelNwaijah/Documents/Textspel.xlsx')
ws = wb.active
for row in range(1, 4):
    for col in range(1, 4):
        char = get_column_letter(col)
        print(ws[char + str(row)].value)

wb.save('C:/Users/MichaelNwaijah/Documents/Textspel.xlsx')
