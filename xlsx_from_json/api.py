from typing import Dict, List

import attr
from openpyxl import Workbook
from openpyxl.conftest import Worksheet

from .models import Style, CellWithSize
from .utils import str_cell_range, style_and_merge_cell_range, style_single_cell


def xlsx_from_json(json_data: Dict, default_style: Style = None) -> Workbook:
    default_style = default_style or Style()

    wb = Workbook()

    start_row = json_data.get("start_row", 1)
    start_column = json_data.get("start_column", 1)

    filler = SheetFiller(wb.active, start_column, start_row, default_style)
    filler.fill(json_data)

    return wb


@attr.s(auto_attribs=True)
class SheetFiller:
    sheet: Worksheet
    start_column: int
    start_row: int
    default_style: Style

    def fill(self, json_data: Dict) -> None:
        current_row = self.start_row

        for row_data in json_data["rows"]:
            row_height = self._fill_row(current_row, row_data["cells"])
            current_row += row_height

    def _fill_row(self, row: int, cells_data: List[Dict]) -> int:
        max_cell_height = 1

        for cell_index, cell_data in enumerate(cells_data):
            cell = self._create_cell(row, cell_index + self.start_column, cell_data)
            max_cell_height = max(cell.height, max_cell_height)

        return max_cell_height

    def _create_cell(self, row: int, column: int, cell_data: Dict) -> CellWithSize:
        cell = self.sheet.cell(row, column)

        cell.value = cell_data["value"]

        width = cell_data.get("width", 1)
        height = cell_data.get("height", 1)

        style = Style.from_json(cell_data.get("style", {}), self.default_style)

        if width == 1 and height == 1:
            style_single_cell(cell, style)
        else:
            cell_range = str_cell_range(column, row, column + width, row + height)
            style_and_merge_cell_range(self.sheet, cell_range, style)

        return CellWithSize(cell, width, height)
