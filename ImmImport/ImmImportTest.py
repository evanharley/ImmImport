import unittest
import random
from ImmImport import ImmImport
import tkinter as tk
from tkinter import filedialog
import os

class Test_ImmImport(unittest.TestCase):
    def setUp(self):
        self.impt = ImmImport()
        self.sheet, self.result = self.impt._get_spreadsheet('C:/Users/evharley/source/repos/ImmImport/ImmImport/Ornithology/1_Test.xlsx')
        self.keys = self.impt._get_keys(self.sheet['OrnithologyItem'])
        self.length = len(self.impt._get_keys(self.sheet['OrnithologyItem']).keys())
        self.impt._write_to_test()
        return super().setUp()

    def test_get_keys(self):
        value = {0: 'item_id',
                 1: 'sex_cd',
                 2: 'age_cd',
                 3: 'specimen_nature',
                 4: 'quantity',
                 5: 'specimen_note',
                 6: 'partial_specimen',
                 7: 'original_type_name',
                 8: 'behaviour',
                 9: 'breeding',
                 10: 'SPN',
                 11: 'plumage',
                 12: 'colour',
                 13: 'iris',
                 14: 'diet',
                 15: 'FAT',
                 16: 'eggs in coll',
                 17: 'eggs in clutch'}
        test_values = self.impt._get_keys(self.sheet['OrnithologyItem'])
        self.assertEqual(value, test_values)
        return 0

    def test_handle_spreadsheet(self):
        self.fail('not implemented')
        return 0

    def test_handle_row(self):
        working_sheet = self.sheet['OrnithologyItem']
        test_row = working_sheet[3]
        values = self.impt._handle_row(test_row, self.keys)
        test_values = {'item_id': 800002,
                       'sex_cd': 'male',
                       'age_cd':'adult',
                       'specimen_nature': 'skin, standard',
                       'SPN': 'RNGR'}
        self.assertDictEqual(test_values, values)
        return 0

    def test_write_query(self):
        test_query = "INSERT into OrnithologyItem(item_id, sex_cd, age_cd, specimen_nature, SPN)\n" +\
            "VALUES (800002, 'male', 'adult', 'skin, standard', 'RNGR')"
        data = self.impt._handle_row(self.sheet['OrnithologyItem'][3], self.keys)
        query = self.impt._write_query(data, 'OrnithologyItem')
        self.assertEqual(test_query, query)
        return 0

    def test_import(self):
        for run in range(random.randint(1, 11)):
            data = self.impt._handle_row(self.sheet['OrnithologyItem'][random.randint(2,51)], self.keys)
            query = self.impt._write_query(data, 'OrnithologyItem')
            fail = random.random()
            if fail >= 0.5:
                rand = random.randint(2, self.length + 2)
                result = self.impt._import(query, test_rand = rand)
                print(result[1])
                self.assertEqual(1, result[0])
            else:
                result = self.impt._import(query, test_rand=1)
                self.assertEqual(0, result)
        return 0

if __name__ == '__main__':
    unittest.main()
