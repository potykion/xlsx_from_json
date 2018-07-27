from itertools import chain

import pytest
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.conftest import Worksheet
from openpyxl.styles import Font

from xlsx_from_json import xlsx_from_json, Style


@pytest.fixture()
def default_style():
    return Style(font=Font(bold=True))


@pytest.fixture()
def json_data_with_single_cell():
    return {
        "rows": [
            {
                "cells": [
                    {
                        "value": "Sample text",
                        "style": {
                            "font": {
                                "name": "Times New Roman",
                                "size": 12
                            },
                            "border": {
                                "bottom": {
                                    "border_style": "medium",
                                    "color": "00000000"
                                }
                            },
                            "alignment": {
                                "horizontal": "center"
                            },
                            "fill": {
                                "patternType": "gray125"
                            }
                        }
                    }
                ]
            }
        ],
        "start_column": 3
    }


@pytest.fixture()
def json_data_with_sized_cell():
    return {
        "rows": [
            {
                "cells": [
                    {
                        "value": "Sample text",
                        "width": 5,
                        "height": 2,
                        "style": {
                            "font": {
                                "name": "Times New Roman",
                                "size": 12
                            },
                            "border": {
                                "bottom": {
                                    "border_style": "medium",
                                    "color": "FFFFFFFF"
                                }
                            }
                        }
                    }
                ]
            }
        ]
    }


@pytest.fixture()
def json_data_with_start_row_and_multiple_cells():
    return {
        "rows": [
            {
                "cells": [
                    {
                        "value": "1x2",
                        "height": 2
                    }
                ]
            },
            {
                "cells": [
                    {
                        "value": "2x1",
                        "width": 2
                    },
                    {
                        "value": "1x1"
                    }
                ]
            }
        ],
        "start_row": 2
    }


@pytest.fixture()
def workbook(json_data_with_single_cell, default_style) -> Workbook:
    return xlsx_from_json(json_data_with_single_cell, default_style)


@pytest.fixture()
def sheet(workbook) -> Worksheet:
    return workbook.active


@pytest.fixture()
def cell(sheet) -> Cell:
    return sheet.cell(row=1, column=3)


def test_sheet_has_values(sheet):
    assert list(filter(None, chain.from_iterable(sheet.values))) == ["Sample text"]


def test_cell_is_shifted_by_offset(cell):
    assert cell.column == 'C'
    assert cell.value == "Sample text"


def test_cell_has_font(cell):
    assert cell.font.name == "Times New Roman"
    assert cell.font.size == 12


def test_cell_has_font_bold_from_default_style(cell):
    assert cell.font.bold


def test_cell_has_border(workbook, cell):
    assert cell.border.bottom.border_style == "medium"
    assert cell.border.bottom.color.rgb == "00000000"


def test_cell_has_alignment(workbook, cell):
    assert cell.alignment.horizontal == "center"


def test_cell_has_fill(workbook, cell):
    assert cell.fill.patternType == "gray125"


def test_sized_cell_is_rendered_as_merged_cells_and_style_set(json_data_with_sized_cell):
    workbook: Workbook = xlsx_from_json(json_data_with_sized_cell)
    sheet = workbook.active
    assert sheet.cell(2, 5).border.bottom.border_style == "medium"


def test_sheet_fill_starts_with_start_row(json_data_with_start_row_and_multiple_cells):
    workbook: Workbook = xlsx_from_json(json_data_with_start_row_and_multiple_cells)
    sheet = workbook.active
    assert sheet.cell(row=2, column=1).value == "1x2"


def test_sheet_has_two_rows(json_data_with_start_row_and_multiple_cells):
    workbook: Workbook = xlsx_from_json(json_data_with_start_row_and_multiple_cells)
    sheet = workbook.active
    assert sheet.cell(row=2, column=1).value == "1x2"
    assert sheet.cell(row=4, column=1).value == "2x1"


def test_sheet_has_two_columns(json_data_with_start_row_and_multiple_cells):
    workbook: Workbook = xlsx_from_json(json_data_with_start_row_and_multiple_cells)
    sheet = workbook.active
    assert sheet.cell(row=4, column=1).value == "2x1"
    assert sheet.cell(row=4, column=3).value == "1x1"


def test_empty_row_render():
    workbook: Workbook = xlsx_from_json({
        "rows": [
            {"cells": []},
            {"cells": [{"value": "op"}]},
            {},
            {"cells": [{"value": "op"}]},
        ]
    })
    sheet = workbook.active
    assert sheet.cell(2, 1).value == "op"
    assert sheet.cell(4, 1).value == "op"


def test_row_and_column_skip():
    workbook: Workbook = xlsx_from_json({
        "rows": [
            {"cells": [{"value": "op"}], "skip_rows": 2, "skip_columns": 3}
        ]
    })
    sheet = workbook.active
    assert sheet.cell(3, 4).value == "op"
