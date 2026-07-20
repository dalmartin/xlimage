from zipfile import ZipFile
import xml.etree.ElementTree as ET

##############Files to Parse ##########
METADATAFILES = ["xl/metadata.xml",
              "xl/richData/richValueRel.xml",
              "xl/richData/_rels/richValueRel.xml.rels",
             ]

######################################

class WorkbookParser:

    def __init__(self, path: str):
        self.zip: ZipFile = ZipFile(path)
        self.parsed: dict[str, ET.Element] = {}

        self.sheets: list[str] = [n for n in self.zip.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")]
        self.parsefiles: list[str] = METADATAFILES + self.sheets
        
        for parsefile in self.parsefiles:
            self.readXml(parsefile)

    def getSheets(self):
        return self.sheets

    def readXml(self, path: str):
        bytes = self.zip.read(path)
        root = ET.fromstring(bytes)
        self.parsed[path] = root

    def getData(self, path: str):
        if path not in self.parsed:
            raise KeyError("Cannot retrieve xml data; specified path is not parsed")

        return self.parsed[path]
    


if __name__ == '__main__':
    wbp = WorkbookParser("testing_excel_files/test2.xlsx")
    _ = wbp.getData("xl/worksheets/sheet1.xml")
