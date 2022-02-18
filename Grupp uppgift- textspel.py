import openpyxl
from openpyxl import Workbook, load_workbook
import openpyxl.utils
from openpyxl.utils import get_column_letter

# Laddar in spreadsheet
wb = load_workbook('C:/Users/MichaelNwaijah/Downloads/2Textspel.xlsx')
ws = wb.active

# room1 är Start
room1 = ws['B2'].value
room2 = ws['B3'].value
Exit = ws['B8'].value
current_position = room1
directions = ['north', 'south', 'east', 'west']


def textspel():
    if 'look' in playerInput:
        look()



def look():
    print(current_position)

while current_position != ws['B8'].value:
    playerInput = str(input(''))
    if playerInput == 'look':
        look()



current_position = room1



#def north():
    #if 'north' in playerinput:


#def south():
#def east():
#def south():

#for row in range(2, 5):
    #for col in range(1, 5):
        #char = get_column_letter(col)
        #print(ws[char + str(row)].value)

wb.save('C:/Users/MichaelNwaijah/Downloads/2Textspel.xlsx')

