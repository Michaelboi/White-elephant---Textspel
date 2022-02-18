import openpyxl
from openpyxl import Workbook, load_workbook
import openpyxl.utils
from openpyxl.utils import get_column_letter

# Laddar in spreadsheet
wb = load_workbook('C:/Users/MichaelNwaijah/Downloads/3Textspel.xlsx')
ws = wb.active

wsRooms = wb['Rooms']
wsMap = wb['Map']
wsSave = wb['Save']

# room1 är Start
room1 = ws['C2'].value
room2 = ws['C3'].value
room3 = ws['C4'].value
room4 = ws['C5'].value
room5 = ws['C6'].value
room6 = ws['C7'].value
Exit = ws['B10'].value
current_position = [room1, room2, room3, room4, room5, room6]
directions = ['north', 'south', 'east', 'west']

def current_position():
    print(current_position[0])




def look():
    print(current_position[0, 6])

def north():
    print()

def south():
    print()

def east():
    print()

def west():
    print()

while current_position != ws['B8'].value:
    playerInput = str(input(''))
    if 'look' in playerInput:
        look()
    if 'north' in playerInput:
        north()

current_position = room1



#for row in range(2, 5):
    #for col in range(1, 5):
        #char = get_column_letter(col)
        #print(ws[char + str(row)].value)

wb.save('C:/Users/MichaelNwaijah/Downloads/3Textspel.xlsx')

