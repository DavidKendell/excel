from openpyxl import Workbook
import pytest
import CRUDcontroller



class test_CRUDcontroller:
    def __init__(self) -> None:
        mockbook = Workbook()
        mockbook.create_sheet("testing")
        mocksheet = mockbook["testing"]
        for i in range(1, 6):
            mocksheet.append([f"{chr(65+j)}{i}" for j in range(3)])
        self.controller = CRUDcontroller.CRUDcontroller(mocksheet)
    def test_find(self):
        row = self.controller.find("A3")
        assert row == 3
        row = self.controller.find("A4")
        assert row == 4
        row = self.controller.find("A7")
        assert row == 0
    def test_add(self):
        self.controller.add("A6", [1, 2, 3])
        assert [x.value for x in self.controller[6]] == ["A6", 1, 2, 3]