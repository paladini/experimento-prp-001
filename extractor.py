#/usr/bin/env python3
import pyowm
import gspread
import time
import json
from oauth2client.service_account import ServiceAccountCredentials

with open('credentials.json') as credentials:
    data = json.load(credentials)

spreadsheet_key = data["spreadsheet_key"]
owm_key = data["owm_key"]
date_column = 'A' # date column
humidity_column = 'B' # humidity column
cur_temp_column = 'C' # current temperature column
min_temp_column = 'D' # minimum temperature column
max_temp_column = 'E' # maximum temperature column
start_row = '2' # the first row with values
end_row = '85' # the last row with values
current_date = time.strftime('%d/%m/%Y')

def compare_dates(date_ws):
    print("Current date: " + current_date + " | Worksheet date: " + date_ws)
    if (date_ws == current_date):
        print("Current date and worksheet date are correct.")
        return True
    else:
        print("Dates are not equal, an error will be thrown.")
        return False

def check_empty_cell(worksheet, empty_row):
    ct = worksheet.acell(cur_temp_column+str(empty_row)).value
    mint = worksheet.acell(min_temp_column+str(empty_row)).value
    maxt = worksheet.acell(max_temp_column+str(empty_row)).value
    ctemp = (ct == '') or (ct == None)
    mintemp = (mint == '') or (mint == None)
    maxtemp = (maxt == '') or (maxt == None)

    if (ctemp and mintemp and maxtemp):
        print('Cells are empty and ready to be inserted.')
        return True
    else:
        print('Cells not empty, no data will be written to the document.')
        return False

# API call to OpenWeatherMap
owm = pyowm.OWM(owm_key)

# OpenWeatherMap - Pantanal [lat=-27.60985, lon=-48.516479]
# http://openweathermap.org/city/7874482
observation = owm.weather_at_id(7874478)
w = observation.get_weather()
humidity = w.get_humidity()
cur_temp = w.get_temperature('celsius')['temp']
min_temp = w.get_temperature('celsius')['temp_min']
max_temp = w.get_temperature('celsius')['temp_max']

# Log-in into Google Spreadsheets
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('experimento-prp-001-4984fbe9985e.json', scope)
gc = gspread.authorize(credentials)
spread = gc.open_by_key(spreadsheet_key)

# Updating worksheets with new values
for i in range(0, 3):
    print("\n-------------------------------------\n")
    print("Start working on worksheet number %d" % i)
    ws = spread.get_worksheet(i)

    # Checking for the first empty row
    print("Trying to find an empty row.")
    cell_list = ws.range(humidity_column+start_row + ":" + humidity_column+end_row)
    empty_row = -1
    for i, cell in enumerate(cell_list):
        if ((cell.value == None) or (cell.value == '')):
            empty_row = i + int(start_row)
            print("An empty row was found.")
            break

    if (empty_row == -1):
        print("An empty row was not found. Skipping to the next worksheet.")
        next

    # Conditional to certify wheter the data should or shouldn't be inserted.
    if ((compare_dates(ws.acell(date_column+str(empty_row)).value)) and (check_empty_cell(ws, empty_row))):
        print("Everything goes fine, worksheet being updated now.")
        ws.update_acell(humidity_column+str(empty_row), humidity)
        ws.update_acell(cur_temp_column+str(empty_row), cur_temp)
        ws.update_acell(min_temp_column+str(empty_row), min_temp)
        ws.update_acell(max_temp_column+str(empty_row), max_temp)
        print('Worksheet was updated.')
    else:
        print("Something went wrong, worksheet was not updated.")

print("\n=================================\n\tEnd of script\n=================================")



