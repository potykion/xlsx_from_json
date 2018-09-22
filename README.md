# xlsx_from_json

Creates xlsx from json via [openpyxl](https://openpyxl.readthedocs.io/en/latest/index.html).

## Usage

Let's create following table:

![Alt-text](/static/example1.jpg?raw=true)


Firstly, define .json representation of table:

```java
{
    // rows to render
    "rows": [
        {
            "cells": [
                {
                    "value": "Firstname",
                    "style": {"font": {"bold": true}}
                },
                {
                    "value": "Lastname",
                    "style": {"font": {"bold": true}}
                }
            ]
        },
        {
            "cells": [
                {"value": "Jill"},
                {"value": "Smith"}
            ]
        },
        {
            "cells": [
                {"value": "Eve"},
                {"value": "Jackson"}
            ]
        }
    ],
    // style applied to all cells
    "default_style": {
        "font": {"name": "Times New Roman", "size": 14}
    }
}
```


Then create openpyxl workbook via ``xlsx_from_json`` function:

```python
import json
from xlsx_from_json import xlsx_from_json

with open("example.json", encoding="utf-8") as f:
    json_data = json.load(f)
    
wb = xlsx_from_json(json_data)
```

Created workbook will have cells with defined values and styles:

```python
sheet = wb.active
assert sheet.cell(row=1, cell=1).value == "Firstname"
assert sheet.cell(row=1, cell=1).font.bold == True
assert sheet.cell(row=1, cell=1).font.name == "Times New Roman"
```


Now you can use workbook according to openpyxl [guide](https://openpyxl.readthedocs.io/en/latest/usage.html).
