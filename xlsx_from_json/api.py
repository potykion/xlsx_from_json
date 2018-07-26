from typing import Dict

from openpyxl import Workbook
from openpyxl.conftest import Worksheet


def xlsx_from_json(json_data: Dict) -> Workbook:
    wb = Workbook()
    sheet = wb.active
    _fill_sheet(sheet, json_data)
    return wb


def _fill_sheet(sheet: Worksheet, json_data: Dict) -> None:
    offset = json_data.get("offset", 0)

    for row_index, row in enumerate(json_data["rows"]):
        for cell_index, cell in enumerate(row["cells"]):
            sheet.cell(row_index + 1, cell_index + 1 + offset).value = cell["value"]
