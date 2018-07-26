from typing import Dict

import attr
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.conftest import Worksheet
from openpyxl.styles import Font, Side, Border, Fill, Alignment, PatternFill
from openpyxl.utils import get_column_letter


@attr.s(auto_attribs=True)
class Style:
    font: Font = Font()
    border: Border = Border()
    fill: Fill = Fill()
    alignment: Alignment = Alignment()


def xlsx_from_json(json_data: Dict, default_style: Style = None) -> Workbook:
    default_style = default_style or Style()

    wb = Workbook()
    sheet = wb.active
    _fill_sheet(sheet, json_data, default_style)
    return wb


def _fill_sheet(sheet: Worksheet, json_data: Dict, default_style: Style) -> None:
    start_column = json_data.get("start_column", 1)
    current_row = json_data.get("start_row", 1)

    for row_data in json_data["rows"]:
        max_height = 1

        for cell_index, cell_data in enumerate(row_data["cells"]):
            row = current_row
            column = cell_index + start_column

            cell = sheet.cell(row, column)
            cell.value = cell_data["value"]

            width = cell_data.get("width", 1)
            height = cell_data.get("height", 1)
            max_height = max(max_height, height)

            style = style_from_json(cell_data.get("style", {}), default_style)

            if width == 1 and height == 1:
                _apply_styles_to_single_cell(cell, style)
            else:
                cell_range = str_cell_range(column, row, column + width, row + height)
                sheet.merge_cells(cell_range)
                style_range(sheet, cell_range, style)

        current_row += max_height


def str_cell_range(start_column: int, start_row: int, end_column: int, end_row: int) -> str:
    from_column = get_column_letter(start_column)
    to_column = get_column_letter(end_column)
    return f"{from_column}{start_row}:{to_column}{end_row}"


def style_from_json(style_json: Dict, default_style: Style) -> Style:
    font_data = style_json.get("font", {})
    font_data = {**vars(default_style.font), **font_data}
    font = Font(**font_data)

    sides_data = style_json.get("border", {})
    border_data = {side: Side(**side_data) for side, side_data in sides_data.items()}
    border_data = {**vars(default_style.border), **border_data}
    border = Border(**border_data)

    fill_data = style_json.get("fill", {})
    fill_data = {**vars(default_style.fill), **fill_data}
    fill = PatternFill(**fill_data)

    alignment_data = style_json.get("alignment", {})
    alignment_data = {**vars(default_style.alignment), **alignment_data}
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
