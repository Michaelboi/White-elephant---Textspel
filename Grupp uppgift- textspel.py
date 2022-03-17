import openpyxl
from openpyxl import Workbook, load_workbook
import openpyxl.utils
from openpyxl.utils import get_column_letter

# Laddar in spreadsheet
wb = load_workbook('C:/Users/MichaelNwaijah/Downloads/Textspel1.xlsx')
ws = wb.active

# Worksheets
wsRooms = wb['Rooms']
wsMap = wb['Map']
wsSave = wb['Save']

#välkommnar spelaren och ger instruktioner på hur man går framåt, bakåt, höger, vänster
print("\nWelcome to our game, hope you enjoy it")
print("The Move Commands are: North, South, West, East")

# start värden till start punkten
current_room_Y = 1
current_room_X = 1
# Ändrings värden till X och Y
current_room_change = 1
X_positions = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E'}
start = True


# Kollar nuvarande postion
def look():
    print(wsMap[f"{X_positions[current_room_X]}{current_room_Y}"].value)



# förflytar spelaren
# Ner
def south():
    global current_room_Y
    if current_room_Y >= 1:
        current_room_Y += current_room_change
        print(wsMap[f'{X_positions[current_room_X]}{current_room_Y}'].value)

# Upp
def north():
    global current_room_Y
    if current_room_Y >= 1:
        current_room_Y -= current_room_change
        print(wsMap[f'{X_positions[current_room_X]}{current_room_Y}'].value)

# Höger
def east():
    global current_room_X
    if current_room_X >= 1:
        current_room_X += current_room_change
        print(wsMap[f'{X_positions[current_room_X]}{current_room_Y}'].value)

# Vänster
def west():
    global current_room_X
    if current_room_X >= 1:
        current_room_X -= current_room_change
        print(wsMap[f'{X_positions[current_room_X]}{current_room_Y}'].value)


# En loop som håller kvar spelaren i input med ett stop villkor
while start:
    player_input = str(input(''))
    if 'quit' in player_input:
        start = False
    if player_input == 'look':
        look()
    if player_input == 'north':
        north()
    if player_input == 'south':
        south()
    if player_input == 'west':
        west()
    if player_input == 'east':
        east()
# Exit
    if player_input == 'B4':
        start = False

# Vägar i spelet
    if current_room_Y == 2 and current_room_X == 5:
        current_room_X = 4
    if current_room_X == 5 and current_room_Y == 3:
        current_room_X = 4
    if current_room_X == 5 and current_room_Y == 1:
        current_room_X = 4
    if current_room_Y == 4 and current_room_X == 3:
        current_room_Y = 3
    if current_room_Y == 4 and current_room_X == 4:
        current_room_Y = 3
    if current_room_X == 1 and current_room_Y == 3:
        current_room_Y = 2
    if current_room_X == 4 and current_room_Y == 1:
        current_room_X = 3

wb.save('C:/Users/MichaelNwaijah/Downloads/Textspel1.xlsx')
