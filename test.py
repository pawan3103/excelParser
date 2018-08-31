from excel_parser import ExcelParser


class TestParser():
    def __init__(self):
        self.filename = 'sample.xlsx'

    def testoutput(self):
        obj = ExcelParser(self.filename)
        res = obj.process()
        data = [{i.item: i.availableStock} for i in res]
        assert len(data) == 2
        assert data[0]['Phone'] == 15
        assert data[1]['Water bottle'] == 50
        return True


if __name__ == "__main__":
    test = TestParser()
    print(test.testoutput())
