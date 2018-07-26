from itertools import chain

import pytest
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.conftest import Worksheet

from xlsx_from_json import xlsx_from_json


@pytest.fixture()
def json_data():
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
                            }
                        }
                    }
                ]
            }
        ],
        "offset": 2
    }


@pytest.fixture()
def sheet(json_data) -> Worksheet:
    workbook: Workbook = xlsx_from_json(json_data)
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
