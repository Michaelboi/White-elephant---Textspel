import openpyxl
from openpyxl import Workbook, load_workbook


# Laddar in Excel-filen och konverterar det till en variabel
wb = load_workbook('C:/Users/MichaelNwaijah/Downloads/Textspel1.xlsx')
ws = wb.active

# Använder variabeln för Excel-filen och tar namnen av sheeten i filen för att konvertera även de till variabler
wsRooms = wb['Rooms']
wsMap = wb['Map']
wsSave = wb['Save']

# Välkommnar spelaren och berättar vad move-kommandona är; framåt, bakåt, höger, vänster
print("\nWelcome to our game, hope you enjoy it")
print("The Move Commands are: North, South, West, East")

# Definierar de två variablerna för de två koordinaterna, de två som står där från början är startkoordinaterna
current_room_Y = 2
current_room_X = 2
# Variabel för hur mycket ett koordinatvärde förändras. Även en dictionary där Excels kolumnkoordinater omvandlas till siffror för att göra det enklare
current_room_change = 1
X_positions = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G'}
# Variabeln som berättar om while-loopen för spelet är av eller på
start = True


# Definierar kommandot look, printar värdet i cellen där man befinner sig, alltså ett rum ID
def look():
    print(wsMap[f"{X_positions[current_room_X]}{current_room_Y}"].value)


# De fyra olika move kommandona för de fyra olika hållen man kan gå
# Syd- och norrkommandona tar variabeln för y koordinaten och adderar eller subtraherar 1 från värdet för att få det nya värdet, printar sedan ID av rummet du flyttat till, alltså som look
# Norr och söder är tvärtom vad gäller + och - på grund av att man går ner i excel om värdet på y blir större
# Söder
def south():
    global current_room_Y
    if current_room_Y >= 1:
        current_room_Y += current_room_change
        print(wsMap[f'{X_positions[current_room_X]}{current_room_Y}'].value)

# Norr
def north():
    global current_room_Y
    if current_room_Y >= 1:
        current_room_Y -= current_room_change
        print(wsMap[f'{X_positions[current_room_X]}{current_room_Y}'].value)

# Öst- och västkommandona tar variabeln för x koordinaten och adderar eller subtraherar 1 från värdet för att få det nya värdet, printar sedan ID av rummet du flyttat till, alltså som look
# Öst
def east():
    global current_room_X
    if current_room_X >= 1:
        current_room_X += current_room_change
        print(wsMap[f'{X_positions[current_room_X]}{current_room_Y}'].value)

# Väst
def west():
    global current_room_X
    if current_room_X >= 1:
        current_room_X -= current_room_change
        print(wsMap[f'{X_positions[current_room_X]}{current_room_Y}'].value)


# En while-loop som håller kvar spelaren i input med ett stop villkor, utför ett av de 5 kommandona beroende på vad spelaren skriver, bryter loopen om spelaren skriver 'quit'
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
# Berättar för spelaren att de vunnit och bryter loopen om de kommer till koordinaterna för slutet
    if current_room_X == 3 and current_room_Y == 5:
        print('\nYou win!')
        start = False


# Väggar i spelet
# Puttar tillbaka spelaren in i mappen om de försöker komma utanför den genom att sätta koordinaterna tillbaka till det de var innan spelaren försökte komma utanför
    if current_room_X == 1:
        current_room_X = 2
    if current_room_Y == 1:
        current_room_Y = 2
    if current_room_Y == 2 and current_room_X == 6:
        current_room_X = 5
    if current_room_X == 6 and current_room_Y == 3:
        current_room_X = 5
    if current_room_X == 6 and current_room_Y == 2:
        current_room_X = 5
    if current_room_Y == 5 and current_room_X == 4:
        current_room_Y = 4
    if current_room_Y == 5 and current_room_X == 5:
        current_room_Y = 4
    if current_room_X == 2 and current_room_Y == 4:
        current_room_Y = 3
    if current_room_X == 5 and current_room_Y == 2:
        current_room_X = 4

wb.save('C:/Users/MichaelNwaijah/Downloads/Textspel1.xlsx')
