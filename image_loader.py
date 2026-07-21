from workbook_parser import WorkbookParser

class ImageLoader:

    def __init__(self, path_to_workbook: str):
        self.wbp: WorkbookParser = WorkbookParser(path_to_workbook)

    def get_image(self, sheet: str, cell: str):
        return self.wbp.getImage((sheet, cell))
