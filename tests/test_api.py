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
        ]
    }


def test_xlsx_from_json(json_data):
    workbook: Workbook = xlsx_from_json(json_data)
    sheet = workbook.active
    assert list(chain.from_iterable(sheet.values)) == ["Sample text"]
