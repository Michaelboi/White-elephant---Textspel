import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# Laddar in spreadsheet
wb = load_workbook('G:/Min enhet/1Textspel.xlsx')
ws = wb.active

Start = 'B2'
player = Start

while player != 'B8':
    def player():


#for row in range(2, 5):
    #for col in range(1, 5):
        #char = get_column_letter(col)
        #print(ws[char + str(row)].value)

wb.save('/1Textspel.xlsx')
