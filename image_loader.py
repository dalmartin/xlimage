from xml.etree.ElementTree import Element

from workbook_parser import WorkbookParser

################### Namespace Constant for XML parsing #################
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rvr": "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel",
    "relations": "http://schemas.openxmlformats.org/package/2006/relationships",
}
########################################################################

class ImageLoader:

    def __init__(self, workbook_path: str):
        self.parser: WorkbookParser = WorkbookParser(workbook_path)
        self.sheetcellIMG: dict[tuple[str, str], str]

        #Intermediate dictionaries:
        self.cellToVM: dict[tuple[str, str], str] = {}
        self.cellToV : dict[tuple[str, str], str] = {}
        self.cellToRID : dict[tuple[str, str], str] = {}

        # Final dict that contains cell, sheet -> image path
        self.cellToPath: dict[tuple[str, str], str] = {}

    def getVMs(self):
        # populate self.sheetcellIMG with mappings of (sheet, cell) -> img path
        
        # sheet.xml -> vm
        # metadata.xml -> valueMetadata[vm] -> v
        # richValueRel.xml -> richValueRels[v].id -> rId
        # rels/richValueRel.xml.rels -> Relationsihps[rId].Target -> Img path


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

    def getV(self):
        # get the metadata parsed
        data = self.parser.getData("xl/metadata.xml")
        idxV: dict[int, str] = {}

        # populate a metadata value index to a metadata value
        for i, rc in enumerate(data.findall(".//main:rc", NS)):
            v = rc.get("v")
            if v:
                idxV[i+1] = v

        # populate cell to metadata value
        for cell, vm in self.cellToVM.items():
            self.cellToV[cell] = idxV[int(vm)]

    def getRID(self):
        # get rich value relation data
        data = self.parser.getData("xl/richData/richValueRel.xml")
        idxRID: dict[int, str] = {}

        for i, rel in enumerate(data.findall(".//rvr:rel", NS)):
            rid = rel.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if rid:
                idxRID[i] = rid

        for cell, v in self.cellToV.items():
            self.cellToRID[cell] = idxRID[int(v)]

    def getImgPath(self):
        # get rich value relationships (actual path to images)
        data = self.parser.getData("xl/richData/_rels/richValueRel.xml.rels")
        ridImgPath: dict[str, str] = {}

        for relation in data.findall(".//relations:Relationship", NS):
            rid = relation.get("Id")
            path = relation.get("Target")
            if rid and path and relation.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                ridImgPath[rid] = path

        for cell, rId in self.cellToRID.items():
            self.cellToPath[cell] = ridImgPath[rId]


        print(self.cellToPath)

if __name__ == '__main__':
    il = ImageLoader('testing_excel_files/test2.xlsx')
    il.getVMs()
    il.getV()
    il.getRID()
    il.getImgPath()
