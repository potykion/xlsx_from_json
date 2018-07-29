# xlsx_from_json

Creates xlsx from json via [openpyxl](https://openpyxl.readthedocs.io/en/latest/index.html).

## Usage

Create .json file with following structure:

```java
{
    "rows": [
        {
            "cells": [
                {
                    "value": "Sample text",
                    // merge 5x2 cell range
                    "width": 5,
                    "height": 2,
                    // openpyxl style: https://openpyxl.readthedocs.io/en/2.5/styles.html
                    "style": {
                        "font": {
                            "name": "Times New Roman",
                            "size": 12
                        },
                        "border": {
                          "bottom": {
                            // openpyxl border styles: // https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/borders.html
                            "border_style": "medium",
                            "color": "FFFFFFFF"
                          }
                        }
                    }
                }
            ],
            // move row by 2x1
            "skip_columns": 2,
            "skip_rows": 1,
            // change row height
            "row_height": 10
        }
    ],
    "start_column": 1,
    "start_row": 1,
    // to change column widths
    "column_widths": [
        {
            // column_number or column_letter
            "column_number": 1,
            "width": 10
        }
    ]
}
```

Create openpyxl workbook via ``xlsx_from_json`` function:

```python
import json
from xlsx_from_json import xlsx_from_json

with open("data.json", encoding="utf-8") as f:
    json_data = json.load(f)
    
wb = xlsx_from_json(json_data)
```

Created workbook will have values and styles defined above:

```python
sheet = wb.active
assert sheet.cell(row=2, cell=3).value == "Sample text"
```

You can pass default style to xlsx builder:

```python
from xlsx_from_json import Style
from openpyxl.styles import Font

default_style = Style(font=Font("Times New Roman", 12, True))

wb = xlsx_from_json(json_data, default_style)
sheet = wb.active

assert sheet.cell(row=2, cell=3).font.name == "Times New Roman"
assert sheet.cell(row=2, cell=3).font.size == 12
assert sheet.cell(row=2, cell=3).font.bold
```


Now you can use workbook according to openpyxl [guide](https://openpyxl.readthedocs.io/en/latest/usage.html).
