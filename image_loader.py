from xml.etree.ElementTree import Element

from workbook_parser import WorkbookParser


class ImageLoader:

    def __init__(self, workbook_path: str):
        self.parser: WorkbookParser = WorkbookParser(workbook_path)
        self.sheetcellIMG: dict[tuple[str, str], str]

        #Intermediate dictionaries:
        self.cellToVM: dict[str, str] = {}

    def getVMs(self):
        # populate self.sheetcellIMG with mappings of (sheet, cell) -> img path
        
        # sheet.xml -> vm
        # metadata.xml -> valueMetadata[vm] -> v
        # richValueRel.xml -> richValueRels[v].id -> rId
        # rels/richValueRel.xml.rels -> Relationsihps[rId].Target -> Img path

        NS = {
            "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        }

        for sheet in self.parser.getSheets():
            # get sheet name
            sheetName: str = ""
            i_start: int = sheet.find("xl/worksheets/") + len("xl/worksheets/")
            i_end: int = sheet.find(".xml")
            if i_start != -1 and i_end != -1:
                sheetName = sheet[i_start:i_end]
            else:
                raise NameError(f"{sheet} is not named properly; Failed to store sheet name")

            # Get sheet data
            sheetData = self.parser.getData(sheet)
            for cell in sheetData.findall(".//main:c", NS):
                vm = cell.get("vm")
                if vm is not None:
                    self.cellToVM[(sheetName, cell.get('r'))] = vm

            return

    #TODO implement getV
    def getV(self):
        pass

    #TODO implement getRID
    def getRID(self):
        pass

    #TODO implement getImgPath
    def getImgPath(self):
        pass

if __name__ == '__main__':
    il = ImageLoader('testing_excel_files/test2.xlsx')
    il.getImgPaths()
    print(il.cellToVM)
