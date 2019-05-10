import openpyxl
import pyodbc
import tkinter as tk
import os
from tkinter import filedialog
from datetime import datetime

class ImmImport():
    def __init__(self, *args, **kwargs):
        root = tk.Tk()
        self._connection = pyodbc.connect('DSN=IMM Prod; Trusted_Connection=yes;')
        self.cursor = self._connection.cursor()
        return super().__init__(*args, **kwargs)

    def _get_spreadsheet(self, filename):
        try:
            spreadsheet = openpyxl.load_workbook(filename)
            result = 'Success'
        except FileNotFoundError as e:
            spreadsheet = []
            result = "Something went wrong!!"
        return spreadsheet, result

    def _get_keys(self, sheet):
        key_row = sheet[1]
        keys = {i: key_row[i].value for i in range(len(key_row))}
        return keys

    def _write_to_test(self):
        self._connection.close()
        self._connection = pyodbc.connect('DSN=ImportTest; Trusted_Connection=yes;')
        self.cursor = self._connection.cursor()
        return 0, 'Database connection changed to Test'

    def _write_to_prod(self):
        self._connection.close()
        self._connection = pyodbc.connect('DSN=IMM Prod; Trusted_Connection=yes;')
        self.cursor = self._connection.cursor()
        return 0, 'Database connection changed to Prod'

    def _handle_spreadsheet(self, spreadsheet):
        for tab in spreadsheet.sheetnames:
            working_sheet = spreadsheet[tab]
            keys = self.get_keys(working_sheet)
            for row in working_sheet.rows:
                data_to_import = self._handle_row(row, keys)
                query = self._write_query(data_to_import, tab)
                result = self._import(query, test_rand=1)
                

        return 0

    def _handle_row(self, row, keys):
        data = {keys[i]: row[i].value for i in range(len(row)) if row[i].value is not None and row[i].value is not ''}
        return data

    def _write_query(self, data, table):
        query_pt_1 = "INSERT into {0}({1})".format(table, ', '.join(list(data.keys())))
        values = ["'{0}'".format(value) if isinstance(value, str) else str(value) for value in data.values()]
        query_pt_2 = "VALUES ({0})".format(', '.join(values))
        return query_pt_1 + '\n' + query_pt_2

    def _import(self, query, test_rand=0):
        if test_rand == 1:
            print('item imported')
            return 0
        elif test_rand > 1:
            test_rand -= 2
            print('it broke')
            return (1, query)
        else:
            try:
                self.cursor.execute(query)
            except pyodbc.IntegrityError as e:
                print('it broke')
                return (1, query, e)
        return (0, query, 'Sucess')

    def main(self):
        directory = filedialog.askdirectory()
        for file in os.listdir(directory):
            spreadsheet = self._get_spreadsheet
            result = self._handle_spreadsheet(spreadsheet)



if __name__ == '__main__':
    impt = ImmImport()
    
    