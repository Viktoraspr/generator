"""
The task is to read data from .xlsx file
The selected library - python_calamine. The choice was made after evaluating the information found here:
https://hakibenita.com/fast-excel-python

"""

import itertools
import json
import os
import time

from python_calamine import CalamineWorkbook

EXCEL_PATH = r'C:\Pycharm projects\generator\Generator.xlsx'


class Generator:
    """
    Class creates object and allows from Excel file parse to json file.

    excel_path - Excel file path
    json_path: optional if it needs to declare the path
    """
    def __init__(self, excel_path: str = EXCEL_PATH, json_path: str | None = None):
        self.excel_path = excel_path
        self.json_path = json_path

        self.calamine_data: CalamineWorkbook | None = None
        self.data_others_sheets: dict = {}

    def generate_pn_codes_from_pn_sheet(self, sheet_name: str = 'PN') -> tuple[list, itertools.product]:
        """
        Generate PN code from "PN" sheet
        :return: PN codes in list
        """
        data_from_PN_sheet = self.calamine_data.get_sheet_by_name(sheet_name).to_python()

        # Reading columns names and enumerate them
        fields_values, columns = {}, []
        for index, column in enumerate(data_from_PN_sheet[0]):
            fields_values[index] = []
            columns.append(column)

        # Reading values from PN data fields
        for row in data_from_PN_sheet[1:]:
            for index, value in enumerate(row):
                if value:
                    if isinstance(value, float):
                        value = int(value)
                    fields_values[index].append(value)

        generated_all_pn_values = itertools.product(*fields_values.values())

        return columns, generated_all_pn_values

    def read_general_sheet(self, sheet_name: str = 'GENERAL'):
        """
        Reads data form General sheet
        :param sheet_name: sheet name. By default, value is "GENERAL"
        :return: data from general sheet
        """
        data = self.calamine_data.get_sheet_by_name(sheet_name).to_python()
        result = {}
        for column, value in zip(data[0], data[1]):
            if isinstance(value, float):
                value = int(value)
            result[column] = value
        return result

    def generate_data_other_sheets(self) -> None:
        """
        Generates data in data sheets (excluding 'PN' and 'General' sheets
        """

        sheets = self.calamine_data.sheet_names
        sheets.remove('PN')
        sheets.remove('GENERAL')

        for sheet in sheets:
            sheet_data = self.calamine_data.get_sheet_by_name(sheet).to_python()
            columns = sheet_data[0]
            result = {}
            for row in sheet_data[1:]:
                key = row[0]
                fields_values = {}
                for key_1, value in zip(columns[1:], row[1:]):
                    if isinstance(value, float):
                        if value >= 1:
                            value = int(value)
                        else:
                            value = f'{value * 100}%'
                    fields_values[key_1] = value
                result[key] = fields_values
            self.data_others_sheets[sheet_data[0][0]] = result

    def prepare_date_for_json_file(self, pn_codes: tuple[list, itertools.product]):
        """
        Reads data from data sheets and converts to Python data type
        :param pn_codes: pn_codes, which need insert in data
        :return: data for json file
        """
        data_from_general_sheet = self.read_general_sheet()

        self.generate_data_other_sheets()

        result = []

        for values in pn_codes[1]:
            element = {
                'PN': ''.join(str(value) for value in values)
            }
            element.update(data_from_general_sheet)

            for column, value in zip(pn_codes[0][1:], values[1:]):
                element.update(self.data_others_sheets[column][value])

            result.append(element)

        return result

    def write_data_to_json_file(self, data: list[dict]) -> None:
        """
        Writes data to json file
        :param data: data needed to write to json file
        """

        if not self.json_path:
            file = f"{int(time.time())}_{os.path.basename(self.excel_path).split('.')[0]}.json"
            path = os.path.dirname(self.excel_path)
            self.json_path = os.path.join(path, file)

        json_object = json.dumps(data)
        with open(self.json_path, 'w') as json_file:
            json_file.write(json_object)

    def run(self) -> str:
        """
        Class running file
        :return: path to json file
        """
        with open(self.excel_path, 'rb') as file:
            self.calamine_data = CalamineWorkbook.from_filelike(file)

        sheet_names = self.calamine_data.sheet_names
        if not ("PN" in sheet_names and 'GENERAL' in sheet_names):
            raise FileExistsError(f"{self.excel_path} doesn't have PN or GENERAL sheets")

        pn_codes = self.generate_pn_codes_from_pn_sheet()

        result = self.prepare_date_for_json_file(pn_codes)
        self.write_data_to_json_file(result)

        return self.json_path
