import os
import json
import requests
from openpyxl import load_workbook


class PyXLPro(object):
    def __init__(self):
        self.xl = None
        self.xl_name = ''
        self.url = 'https://public.rts.iebc.or.ke/jsons/round1/results/'

    def load_xl(self):
        try:
            for file in os.listdir('input'):
                file_path = 'input/' + file
                # print(file)
                self.xl_name = file
                self.xl = load_workbook(file_path)
                # print(self.xl.get_sheet_names())
                self.process_xl()
        except:
            print('Something wrong happended!')

    def process_xl(self):
        print('****')
        ws = self.xl.active
        max_rows = ws.max_row + 1
        for row in range(2, 4):
            self.construct_url(str(row), ws)
            self.fetch_results()
        self.save_xl()

    def construct_url(self, row, worksheet):
        county = '1' + worksheet['A' + row].value
        constituency = county + worksheet['C' + row].value
        ward = constituency + worksheet['E' + row].value
        polling_centre = ward + worksheet['G' + row].value
        polling_station = worksheet['J' + row].value
        self.url += 'Kenya_Elections_Presidential%2F1%2F'
        self.url += county + '%2F' + constituency + '%2F' + ward + '%2F'
        self.url += polling_centre + '%2F1' + polling_station + '%2Finfo.json'
        print(self.url)

    def fetch_results(self):
        print('Fetching...')
        response = requests.get(self.url)
        data = response.text
        json_data = json.loads(data)
        print(json_data)
        self.url = 'https://public.rts.iebc.or.ke/jsons/round1/results/'

    def save_xl(self):
        if not os.path.exists('output'):
            os.makedirs('output')
        output_path = 'output/' + self.xl_name
        self.xl.save(output_path)


run = PyXLPro()

run.load_xl()
