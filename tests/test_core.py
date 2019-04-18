import unittest
import pathlib
from diff_excel.core import *


class CoreTestSuite(unittest.TestCase):

    def test_diff_excel_data(self):
        from_file_path = pathlib.Path("./tests/test_data/from_book.xlsx").resolve()
        to_file_path = pathlib.Path("./tests/test_data/to_book.xlsx").resolve()
        # print(from_file_path)
        actual = diff_excel_data(from_file_path, to_file_path, key="id")
        expected = {'diff_results': {'1': {'a': {'code': '==', 'status': 'equal', 'from': '1', 'to': '1', 'value': '1'},
                                           'b': {'code': '=>', 'status': 'replace', 'from': '2', 'to': '3',
                                                 'value': '3'},
                                           'c': {'code': '==', 'status': 'equal', 'from': '3', 'to': '3',
                                                 'value': '3'}},
                                     '2': {'a': {'code': '==', 'status': 'equal', 'from': '4', 'to': '4', 'value': '4'},
                                           'b': {'code': '=>', 'status': 'replace', 'from': '5', 'to': '3',
                                                 'value': '3'},
                                           'c': {'code': '==', 'status': 'equal', 'from': '6', 'to': '6',
                                                 'value': '6'}}, '3': {
                'a': {'code': '=>', 'status': 'replace', 'from': '7', 'to': '12', 'value': '12'},
                'b': {'code': '==', 'status': 'equal', 'from': '8', 'to': '8', 'value': '8'},
                'c': {'code': '=>', 'status': 'replace', 'from': '9', 'to': '2', 'value': '2'}}}, 'delete_rows': [],
                    'delete_columns': [], 'insert_rows': [], 'insert_columns': [],
                    'diff_columns': ['a', 'b', 'c'],
                    'diff_keys': ['1', '2', '3']}

        self.assertEqual(expected, actual)

    def test_show_diff_excel_data_as_text(self):
        diff_excel_data_results = {
            'diff_results': {'1': {'a': {'code': '', 'status': 'equal', 'from': '1', 'to': '1', 'value': '1'},
                                   'b': {'code': ' ', 'status': 'replace', 'from': '2', 'to': '3', 'value': '3'},
                                   'c': {'code': '', 'status': 'equal', 'from': '3', 'to': '3', 'value': '3'}},
                             '2': {'a': {'code': '', 'status': 'equal', 'from': '4', 'to': '4', 'value': '4'},
                                   'b': {'code': ' ', 'status': 'replace', 'from': '5', 'to': '3', 'value': '3'},
                                   'c': {'code': '', 'status': 'equal', 'from': '6', 'to': '6', 'value': '6'}},
                             '3': {'a': {'code': ' ', 'status': 'replace', 'from': '7', 'to': '12', 'value': '12'},
                                   'b': {'code': '', 'status': 'equal', 'from': '8', 'to': '8', 'value': '8'},
                                   'c': {'code': ' ', 'status': 'replace', 'from': '9', 'to': '2', 'value': '2'}}},
            'delete_rows': [], 'delete_columns': [], 'insert_rows': [], 'insert_columns': [],
            'diff_columns': ['a', 'b', 'c'],
            'diff_keys': ['1', '2', '3']}

        actual = show_diff_excel_data_as_text(diff_excel_data_results)
        expected = """行と列の削除や追加はありませんでした．
======================================
1行: 
  b列: 2 -> 3
2行: 
  b列: 5 -> 3
3行: 
  a列: 7 -> 12
  c列: 9 -> 2
======================================"""
        self.assertEqual(expected, actual)


if __name__ == '__main__':
    unittest.main()
