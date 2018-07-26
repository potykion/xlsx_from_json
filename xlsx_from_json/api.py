from typing import Dict

import attr
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.conftest import Worksheet
from openpyxl.styles import Font, Side, Border, Fill, Alignment, PatternFill
from openpyxl.utils import get_column_letter


@attr.s(auto_attribs=True)
class Style:
    font: Font
    border: Border
    fill: Fill
    alignment: Alignment


def xlsx_from_json(json_data: Dict) -> Workbook:
    wb = Workbook()
    sheet = wb.active
    _fill_sheet(sheet, json_data)
    return wb


def _fill_sheet(sheet: Worksheet, json_data: Dict) -> None:
    offset = json_data.get("offset", 0)

    for row_index, row_data in enumerate(json_data["rows"]):
        for cell_index, cell_data in enumerate(row_data["cells"]):
            row = row_index + 1
            column = cell_index + 1 + offset

            cell = sheet.cell(row, column)
            cell.value = cell_data["value"]

            width = max(cell_data.get("width", 1), 1)
            height = max(cell_data.get("height", 1), 1)

            style = style_from_json(cell_data["style"])

            if width == 1 and height == 1:
                _apply_styles_to_single_cell(cell, style)
            else:
                from_column = get_column_letter(column)
                to_column = get_column_letter(column + width)
                cell_range = f"{from_column}{row}:{to_column}{row + height}"
                sheet.merge_cells(cell_range)
                style_range(sheet, cell_range, style)


def style_from_json(style_json: Dict) -> Style:
    font_data = style_json.get("font", {})
    font = Font(**font_data)

    sides_data = style_json.get("border", {})
    border_data = {side: Side(**side_data) for side, side_data in sides_data.items()}
    border = Border(**border_data)

    fill_data = style_json.get("fill", {})
    fill = PatternFill(**fill_data)

    alignment_data = style_json.get("alignment", {})
    alignment = Alignment(**alignment_data)

    return Style(font, border, fill, alignment)


def _apply_styles_to_single_cell(cell: Cell, style: Style) -> None:
    for attrib, value in attr.asdict(style).items():
        if value:
            setattr(cell, attrib, value)


def style_range(ws: Worksheet, cell_range: str, style: Style) -> None:
    """
    Source:
    https://openpyxl.readthedocs.io/en/2.5/styles.html#styling-merged-cells
    """
    top = Border(top=style.border.top)
    left = Border(left=style.border.left)
    right = Border(right=style.border.right)
    bottom = Border(bottom=style.border.bottom)

    first_cell = ws[cell_range.split(":")[0]]
    if style.alignment:
        ws.merge_cells(cell_range)
        first_cell.alignment = style.alignment

    if style.font:
        first_cell.font = style.font

    rows = ws[cell_range]
    for cell in rows[0]:
        cell.border = cell.border + top
    for cell in rows[-1]:
        cell.border = cell.border + bottom

    for row in rows:
        l = row[0]
        r = row[-1]
        l.border = l.border + left
        r.border = r.border + right
        if style.fill:
            for c in row:
                c.fill = style.fill
