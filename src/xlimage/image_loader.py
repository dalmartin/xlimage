from workbook_parser import WorkbookParser

######################################################################
# ImageLoader is currently a neat wrapper for WorkbookParser. In the
# future, there will be more functionality from this class, but there
# is only trivial calling of the WorkbookParser getImage function as
# of now.
######################################################################

class ImageLoader:

    def __init__(self, path_to_workbook: str):
        self.wbp: WorkbookParser = WorkbookParser(path_to_workbook)

    def get_image(self, sheet: str, cell: str):
        return self.wbp.getImage((sheet, cell))
