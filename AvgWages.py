import openpyxl
import requests
import xlrd
import time
from datetime import date
import sys
from pathlib import Path

# Template for file URL at CBS site
URL = 'https://old.cbs.gov.il/archive/{}/y_labor/e1_06.xls'
USERSELECTION = ['Y', 'N']


def get_response(url):
    tries = 1
    while tries <= 4:
        try:
            print(f'Getting from {url}')
            response = requests.get(url)
        except requests.exceptions.ConnectionError as e:
            print(f'{e}: could not contact CBS server at {url} (attempt #{tries})\nRetrying in 15 seconds...')
            tries += 1
            time.sleep(15)
            continue
        else:
            return response
    else:
        sys.exit("Program timed out. Check your internet connection and try again")


def download_files():
    overwrite_all = False
    ask_overwrite = False
    user_selection = input("Overwrite all existing files? [Y/N] or [A] to ask for each file.\n").upper()
    valid_inputs = USERSELECTION + ['A']
    while user_selection not in valid_inputs:
        user_selection = input("Invalid input\n").upper()
    if user_selection == 'Y':
        print("Program will overwrite all existing files.")
        overwrite_all = True
    elif user_selection == 'N':
        print("Program will only download missing files")
    else:
        print("Program will ask overwrite permission per file")
        ask_overwrite = True
    dates_to_remove = []
    for date in dates:
        filename = date + '.xls'
        if Path(filename).exists():
            if ask_overwrite:
                overwrite = input(f"File {filename} already exists. Overwrite? [Y/N]\n").upper()
                while overwrite not in USERSELECTION:
                    overwrite = input("Invalid input\n").upper()
                if overwrite == 'N':
                    continue
            elif not overwrite_all:
                print(f"File {filename} already exists. Skipping...\n")
                continue
        url = URL.format(date)
        response = get_response(url)
        if response.apparent_encoding == 'ascii':
            print(f"File for date {date} does not exist on the server. Skipping...")
            dates_to_remove.append(date)
        else:
            if response.content:
                with open(filename, 'wb') as f:
                    f.write(response.content)
    for date in dates_to_remove:
        dates.remove(date)


def save_data(outfile='Results'):
    _datacolumn=22
    outfile += '.xlsx'
    data_to_write = []
    wb = openpyxl.Workbook()
    ws = wb.active
    for date in dates:
        try:
            dataws = xlrd.open_workbook(date + '.xls').sheet_by_index(0)
        except:
            print(f"Cannot read file {date}.xls. Skipping...")
            continue
        datacell = list(filter(lambda x: x[1].value == 'מדדים ', enumerate(dataws.col(_datacolumn))))[0]
        rowindex = datacell[0]-1
        while not dataws.col(_datacolumn)[rowindex].value:
            rowindex -= 1
        data = dataws.col(_datacolumn)[rowindex].value
        data_to_write.append(data)
    ws.append(dates)
    ws.append(data_to_write)
    wb.save(outfile)


def get_dates():
    year_now = int(date.today().strftime('%Y'))
    month_now = int(date.today().strftime('%m'))
    while True:
        try:
            start_year, start_month = [int(_) for _ in input("Enter starting year and month (in the format: 2018 05)\n")
                                       .split()]
            end_year, end_month = [int(_) for _ in input("Enter year and month to end on (in the format: 2020 05)\n")
                                   .split()]
            if start_year > end_year or (start_year == end_year and start_month > end_month):
                print("Starting date cannot be after ending date")
                continue
            if start_year > year_now or (start_year == year_now and start_month > month_now):
                print("Starting date cannot be after today")
                continue
        except ValueError:
            print("Please enter valid numbers")
            continue
        else:
            break
    if end_year > year_now or (end_year == year_now and end_month > month_now):
        print("CBS does not hold future data. Changing ending date to today")
        end_year = year_now
        end_month = month_now
    years = list(range(start_year, end_year+1))
    if len(years) == 1:
        dates = list(map(lambda x: str(start_year) + str(x).zfill(2), range(start_month, end_month + 1)))
    else:
        for year in years:
            if year == start_year:
                dates = list(map(lambda x: str(year) + str(x).zfill(2), range(start_month, 13)))
            elif year == end_year:
                dates.extend(list(map(lambda x: str(year) + str(x).zfill(2), range(1, end_month+1))))
            else:
                dates.extend(list(map(lambda x: str(year) + str(x).zfill(2), range(1, 13))))
    return dates


if __name__ == '__main__':
    dates = get_dates()
    download_files()
    save_data()
