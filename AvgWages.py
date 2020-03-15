import openpyxl
import requests
import xlrd

URL_S = 'https://old.cbs.gov.il/archive/'
URL_E = '/y_labor/e1_06.xls'
YEARS = ['2017', '2018', '2019']
MONTHS = [str(month).zfill(2) for month in range(1, 13)]


def download_files():
    for year in YEARS:
        for month in MONTHS:
            filename = "./" + year + month + '.xls'
            url = URL_S + year + month + URL_E
            try:
                response = requests.get(url)
            except requests.exceptions.ConnectionError as e:
                print(f'cannot connect. {e}')
            if response.content:
                print(f'getting from {url}')
                with open(filename, 'wb') as f:
                    f.write(response.content)


def save_data(filename):
    if not filename.endswith('.xlsx'):
        print('Unsupported file type')
        return
    DATACOL = 22
    data_to_write = []
    headers = []
    wb = openpyxl.Workbook()
    ws = wb.active
    for year in YEARS:
        for month in MONTHS:
            try:
                dataws = xlrd.open_workbook(year+month+'.xls').sheet_by_index(0)
            except:
                continue
            headers.append(month + '-' + year)
            for rowindex, data in enumerate(dataws.col(DATACOL)):
                data = data.value
                data
                if data != 'מדדים ':
                    continue
                else:
                    while not dataws.col(DATACOL)[rowindex-1].value:
                        rowindex -= 1
                    data_to_write.append(dataws.col(DATACOL)[rowindex-1].value)
    ws.append(headers)
    ws.append(data_to_write)
    wb.save(filename)


if __name__ == '__main__':
    download_files()
    save_data('results.xlsx')