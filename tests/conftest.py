import shutil
from pathlib import Path

import pytest

from xlimage.workbook_parser import WorkbookParser

DATA = Path(__file__).parent / "data"

# Make a copy of the test workbook, and return a path to the copy
@pytest.fixture
def workbook_path(tmp_path) -> str:
    dest = tmp_path / "test_workbook.xlsx"
    shutil.copy(DATA / "test_workbook.xlsx", dest)
    return dest

# Make a WorkbookParser object for every test using a seperate test workbook copy
@pytest.fixture
def wbp(workbook_path):
    return WorkbookParser(workbook_path)
