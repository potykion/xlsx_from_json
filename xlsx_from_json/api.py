from typing import Dict

from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.conftest import Worksheet
from openpyxl.styles import Font, Side, Border


def xlsx_from_json(json_data: Dict) -> Workbook:
    wb = Workbook()
    sheet = wb.active
    _fill_sheet(sheet, json_data)
    return wb


def _fill_sheet(sheet: Worksheet, json_data: Dict) -> None:
    offset = json_data.get("offset", 0)

    for row_index, row in enumerate(json_data["rows"]):
        for cell_index, cell_data in enumerate(row["cells"]):
            cell = sheet.cell(row_index + 1, cell_index + 1 + offset)
            cell.value = cell_data["value"]
            _apply_styles(cell, cell_data)


def _apply_styles(cell: Cell, cell_data: Dict) -> None:
    for attr, data in cell_data["style"].items():
        if attr == "font":
            cell.font = Font(**data)
        elif attr == "border":
            sides = {side: Side(**side_data) for side, side_data in data.items()}
            cell.border = Border(**sides)
