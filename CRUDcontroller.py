from openpyxl.worksheet.worksheet import Worksheet
class CRUDcontroller:
    def __init__(self, data: Worksheet) -> None:
        self.data = data
    def find(self, findId) -> int:
        for x in self.data["A"]:
            if x.value == findId:
                return x.row
        return 0
    def add(self, newId, newValues) -> None:
        self.data.append([newId, *newValues])
    def delete(self, delId) -> None:
        row = self.find(delId)
        if not row:
            print("Id not found")
            return
        self.data.delete_rows(row)
    def update(self, row, newValues: list):
        for entry in self.data[row][1:]:
            value = newValues[entry.column-2]
            entry.value = value if value != "" else entry.value
