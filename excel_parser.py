import sys

from openpyxl import load_workbook
from collections import namedtuple
from io import BytesIO


class ExcelParser():
    def __init__(self, filename):
        self.wb = load_workbook(filename, data_only=True, guess_types=True)
        self.Order = namedtuple('Order', ['item', 'availableStock'])

    def getHeaders(self, row):
        headers = {}
        for cell in row:
            headers[cell.column] = cell.value
        return headers

    def readXformSheet(self, sheetname):
        rows = []
        try:
            sheet = self.wb.get_sheet_by_name(sheetname)
            sheet_rows = sheet.rows
            header_dict = self.getHeaders(next(sheet_rows))
            for row in sheet_rows:
                row_dict = {}
                for cell in row:
                    row_dict[header_dict[cell.column]] = cell.value
                rows.append(row_dict)
        except KeyError:
            pass
        return rows

    def readSheet(self):
        sheetNames = self.wb.get_sheet_names()
        data = self.readXformSheet(sheetNames[0]) if sheetNames else []
        return data

    def processData(self, data):
        if 'item' in data.keys() and 'availableStock' in data.keys():
            return self.Order(data['item'], data['availableStock'])
        return False

    def output(self, content):
        data = [self.processData(c) for c in content]
        return data

    def process(self):
        content = self.readSheet()
        data = self.output(content)
        return data


# You can also pass `sys.argv[1]` instead of filename if you have disk-file
if __name__ == "__main__":
    if len(sys.argv) == 2:
        file_data = open(sys.argv[1], 'rb')  # Coverting the file to bytes
        filename = BytesIO(file_data.read())
        obj = ExcelParser(filename)
        res = obj.process()
        print(res)
    else:
        print("Excel file missing!!")
