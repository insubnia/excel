#!/usr/local/bin/python3
import os
import sys
import string
import numpy as np

try:
    from win32com.client import Dispatch
except ModuleNotFoundError:
    import xlrd


class Excel(object):
    def __init__(self, filename):
        self.workbook = None
        self.sheet_names = None

        if 'win32com' in sys.modules:
            excel = Dispatch('Excel.Application')
            filename = os.path.abspath(filename)
            self.workbook = excel.Workbooks.Open(filename)
            self.sheet_names = [sheet.name for sheet in self.workbook.Sheets]
        else:
            self.workbook = xlrd.open_workbook(filename)
            self.sheet_names = self.workbook.sheet_names()

    def get_sheet_data(self, sheetname):
        if sheetname not in self.sheet_names:
            print(f"<{sheetname}> doesn't exist in workbook")
            sys.exit()

        if 'win32com' in sys.modules:
            rows = self.workbook.Worksheets(sheetname).UsedRange.Value
            arr = np.ndarray([len(rows), len(rows[0])], dtype=object)
            for r, row in enumerate(rows):
                for c, val in enumerate(row):
                    arr[r][c] = val
        else:
            ws = self.workbook.sheet_by_name(sheetname)
            arr = np.ndarray([ws.nrows, ws.ncols], dtype=object)
            for r in range(ws.nrows):
                for c, val in enumerate(ws.row_values(r)):
                    arr[r][c] = val
        return arr

    @staticmethod
    def num2col(num):
        num += 1
        string = ""
        while num > 0:
            num, remainder = divmod(num - 1, 26)
            string = chr(65 + remainder) + string
        return string

    @staticmethod
    def col2num(col):
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num - 1


if __name__ == "__main__":
    pass
