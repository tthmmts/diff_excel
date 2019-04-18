import unittest
from pathlib import Path
from diff_excel.helpers import *


class HelpersTestSuite(unittest.TestCase):
    """Basic test cases."""

    def test_get_dict_data_from_excel_file(self):
        file_path = Path("./test_data/from_book.xlsx")
        actual = get_dict_data_from_excel_file(file_path)
        expected = ({'1': {'a': '1', 'b': '2', 'c': '3', 'id': '1'},
                     '2': {'a': '4', 'b': '5', 'c': '6', 'id': '2'},
                     '3': {'a': '7', 'b': '8', 'c': '9', 'id': '3'}},
                    ["1", "2", "3"],
                    ["id", "a", "b", "c"])
        self.assertEqual(expected, actual)

    def test_get_dict_data_from_excel_file_with_key(self):
        file_path = Path("./test_data/from_book.xlsx")
        actual = get_dict_data_from_excel_file(file_path, "id")
        expected = ({'1': {'a': '1', 'b': '2', 'c': '3'},
                     '2': {'a': '4', 'b': '5', 'c': '6'},
                     '3': {'a': '7', 'b': '8', 'c': '9'}},
                    ["1", "2", "3"],
                    ["a", "b", "c"])
        self.assertEqual(expected, actual)

    def test_get_dict_data_from_excel_file_with_miss_key(self):
        file_path = Path("./test_data/from_book.xlsx")

        with self.assertRaises(Exception):
            get_dict_data_from_excel_file(file_path, "KEY")

    def test_list_difference(self):
        a = [1, 2, 3, 4]
        b = [2, 5]

        actual = list_difference(a, b)
        expected = [1, 3, 4]
        self.assertEqual(expected, actual)

    def test_list_difference_2nd(self):
        a = [1, 2, 3, 4]
        b = [2, 5]

        actual = list_difference(b, a)
        expected = [5]
        self.assertEqual(expected, actual)

    def test_list_intersection(self):
        a = [1, 2, 3, 4]
        b = [2, 5]

        actual = list_intersection(a, b)
        expected = [2]
        self.assertEqual(expected, actual)

        actual = list_intersection(b, a)
        expected = [2]
        self.assertEqual(expected, actual)

    def test_list_ndiff(self):
        a = ["1", "2", "3", "4"]
        b = ["2", "5"]

        actual = list_ndiff(a, b)
        expected = ['- 1', '  2', '+ 5', '- 3', '- 4']
        self.assertEqual(expected, actual)


if __name__ == '__main__':
    unittest.main()
