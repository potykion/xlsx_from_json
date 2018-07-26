from itertools import chain

import pytest
from openpyxl import Workbook

from xlsx_from_json import xlsx_from_json


@pytest.fixture()
def json_data():
    return {
        "rows": [
            {
                "cells": [
                    {
                        "value": "Sample text",
                    }
                ]
            }
        ],
        "offset": 2
    }


def test_created_workbook_has_values(json_data):
    workbook: Workbook = xlsx_from_json(json_data)
    sheet = workbook.active
    assert list(filter(None, chain.from_iterable(sheet.values))) == ["Sample text"]



def test_wb_row_has_cell_with_offset(json_data):
    workbook: Workbook = xlsx_from_json(json_data)
    sheet = workbook.active
    assert sheet.cell(row=1, column=3).value == "Sample text"
