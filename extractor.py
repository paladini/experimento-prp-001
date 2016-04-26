#/usr/bin/env python3
import pyowm
import gspread
import time
import json
from datetime import datetime
from dateutil import tz
from oauth2client.service_account import ServiceAccountCredentials

with open('credentials.json') as credentials:
    data = json.load(credentials)

spreadsheet_key = data["spreadsheet_key"]
owm_key = data["owm_key"]
date_column = 1 # date column
humidity_column = 2 # humidity column
cur_temp_column = 3 # current temperature column
min_temp_column = 4 # minimum temperature column
max_temp_column = 5 # maximum temperature column
measured_time_column = 6 # time in which the date was measured
current_date = time.strftime('%d/%m/%Y')

def compare_dates(date_ws):
    '''
        Checks if the empty line found is today. Additional step to prevent damages to the table.
    '''
    if (date_ws == current_date):
        print("Current date and worksheet date are correct.")
        return True
    else:
        return False

def check_empty_cells(line):
    '''
        Checks if the worksheet has empty cells at the given line.
        Using Python resource to check if the string is None or '' using conditionals.
    '''
    if (not(line[humidity_column-1] or line[cur_temp_column-1] or line[min_temp_column-1] or line[max_temp_column-1] or line[measured_time_column-1])):
        print('An empty line was found.')
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
time_utc = datetime.utcfromtimestamp(int(observation.get_reception_time())).replace(tzinfo=tz.tzutc())
measured_time = time_utc.astimezone(tz.tzlocal()).strftime("%H:%M:%S")

# Log-in into Google Spreadsheets
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)
spread = gc.open_by_key(spreadsheet_key)

# Opening all worksheets and operating over it locally (less API calls)
ws = [spread.get_worksheet(0), spread.get_worksheet(1), spread.get_worksheet(2)]

# Updating worksheets with new values
for i, current_ws in enumerate(ws):
    print("\n-------------------------------------\n")
    print("Start working on worksheet number %d." % i)
    print("Trying to find an empty row...")
    for j, line in enumerate(current_ws.get_all_values()):
        if (compare_dates(line[date_column-1]) and check_empty_cells(line)):
            print("Worksheet being updated now...")
            current_ws.update_cell(j+1, humidity_column, int(humidity))
            current_ws.update_cell(j+1, cur_temp_column, cur_temp)
            current_ws.update_cell(j+1, max_temp_column, max_temp)
            current_ws.update_cell(j+1, min_temp_column, min_temp)
            current_ws.update_cell(j+1, measured_time_column, measured_time)
            print('Worksheet was updated.')
            break
        else:
            next

print("\n=================================\n\tEnd of script\n=================================\n\n")