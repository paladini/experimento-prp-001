#!/usr/bin/env python3

import pyowm

from datetime import datetime
from gspread import authorize
from json import load
from os import path
from oauth2client.service_account import ServiceAccountCredentials


def get_resources(api_key, wtr_code):
    owm = pyowm.OWM(api_key).weather_at_id(wtr_code)
    wtr = owm.get_weather()

    today = datetime.today()
    cur_date = today.strftime("%d/%m/%Y")
    cur_time = today.strftime("%H:%M:%S")
    temps = wtr.get_temperature('celsius')

    # order inherent to spreadsheet
    return [
        cur_date, wtr.get_humidity(), temps['temp'], temps['temp_min'],
        temps['temp_max'], cur_time
    ]


def upload_ss(values, keyfile, ss_key):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        keyfile, ['https://spreadsheets.google.com/feeds'])
    gc = authorize(credentials)

    worksheet = gc.open_by_key(ss_key).sheet1

    last = worksheet.row_count + 1
    worksheet.resize(last)
    cell_list = worksheet.range('A{0}:F{0}'.format(last))

    for cell, value in zip(cell_list, values):
        cell.value = value
    worksheet.update_cells(cell_list)


if __name__ == '__main__':

    private_data = load(open(path.join(path.dirname(__file__), 'config.json')))

    api_key = private_data['owm_api_key']
    ss_key = private_data['spreadsheet_key']
    keyfile = path.join(path.dirname(__file__),
                        private_data['json_keyfile_path'])

    upload_ss(get_resources(api_key, 7874478), keyfile, ss_key)
