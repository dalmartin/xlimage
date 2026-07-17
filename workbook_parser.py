from zipfile import ZipFile

class WorkbookParser:

    def __init__(self, path: str):
        self.zip: ZipFile = ZipFile(path)

    def read_xml(self, path: str) -> bytes:
        return self.zip.read(path)


if __name__ == '__main__':
    wp = WorkbookParser("./testing_excel_files/test2.xlsx")
    res = wp.read_xml("xl/worksheets/sheet1.xml")
    print(res)

