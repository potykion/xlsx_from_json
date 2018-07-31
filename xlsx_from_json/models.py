from typing import Dict

import attr
from openpyxl.cell import Cell
from openpyxl.styles import Font, Border, Fill, Alignment, PatternFill, Side


@attr.s(auto_attribs=True)
class Style:
    font: Font = Font()
    border: Border = Border()
    fill: Fill = Fill()
    alignment: Alignment = Alignment()

    @classmethod
    def from_json(cls, style_json: Dict, default_style: 'Style' = None) -> 'Style':
        default_style = default_style or Style()

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

        return cls(font, border, fill, alignment)


@attr.s(auto_attribs=True)
class CellWithSize:
    cell: Cell
    width: int
    height: int
    ignore_height: bool = False
