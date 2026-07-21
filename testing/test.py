import pytest

from ..workbook_parser import WorkbookParser


class TestWorkbookParser:

    @pytest.fixture
    def parser(self):
        return WorkbookParser("tests/data/test_workbook.xlsx")

    def test_parses_required_files(self, parser):
        for file in parser.parsefiles:
            assert file in parser.parsed

    def test_get_data_returns_xml_root(self, parser):
        data = parser.getData("xl/metadata.xml")

        assert data is not None

    def test_get_data_raises_for_unknown_file(self, parser):
        with pytest.raises(KeyError):
            parser.getData("does/not/exist.xml")

    def test_builds_cell_to_vm_mapping(self, parser):
        assert len(parser.cellToVM) > 0

    def test_builds_cell_to_v_mapping(self, parser):
        assert len(parser.cellToV) > 0

    def test_builds_cell_to_rid_mapping(self, parser):
        assert len(parser.cellToRID) > 0

    def test_builds_cell_to_path_mapping(self, parser):
        assert len(parser.cellToPath) > 0

    def test_get_image_returns_bytes(self, parser):
        # Adjust to a real image cell in your test workbook
        image = parser.getImage(("sheet1", "A1"))

        assert isinstance(image, bytes)
        assert len(image) > 0

    def test_get_image_caches_result(self, parser):
        # Adjust to a real image cell in your test workbook
        first = parser.getImage(("sheet1", "A1"))
        second = parser.getImage(("sheet1", "A1"))

        assert first == second
        assert len(parser.images) == 1

    def test_get_image_invalid_cell_raises(self, parser):
        with pytest.raises(KeyError):
            parser.getImage(("sheet1", "Z999"))
