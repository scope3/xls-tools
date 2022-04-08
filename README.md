# xls-tools
Excel and Excel-like tools

# xlrd_like

A minimal interface for accessing excel-type files using the `xlrd` API:

XlrdLike: workbook equivalent
 - `.sheet_names()` - return list of sheet names
 - `.sheet_by_name()` - return an XlSheetLike given by name
 - `.sheet_by_index()` - return an XlSheetLike given by index
 - `.sheets()` - return a list of XlSheetLikes-- requires initializing every sheet

XlSheetLike: worksheet equivalent
 - `.name` - the name of the sheet according to the workbook
 - `.nrows` - int number of rows, 0-indexed
 - `.ncols` - int number of columns, 0-indexed
 - `.row(n)` - return a list of XlCellLike corresponding to the nth (0-indexed) row, or IndexError
 - `.col(k)` - return a list of XlCellLike corresponding to the kth (0-indexed) column, or IndexError
 - `.cell(n,k)` - return the nth row, kth cell, or IndexError
 - `.get_rows()` - row iterator
 - `.row_dict(n)` - return a dict of row n, using row 0 (headers) as keys and XlCellLike as values

XlCellLike: cell equivalent
 - `.ctype` - int, as indicated in `xlrd`
 - `.value` - native value

`xlrd` ctypes are as follows:

```
from xlrd.biffh import (
    XL_CELL_EMPTY,  # 0
    XL_CELL_TEXT,   # 1
    XL_CELL_NUMBER, # 2
    XL_CELL_DATE,   # 3
    XL_CELL_BOOLEAN,# 4
    # XL_CELL_ERROR, # 5
    # XL_CELL_BLANK, # 6 - for use in debugging, gathering stats, etc
)

```

## To use:

```
>>> from xls_tools import open_xl
>>> xl = open_xl(filename)
>>>
```

## Google sheets

Also provides an xlrd-like interface for accessing google sheets.  Can also write to google sheets.
For this you need credentials for Google's service API.  See: 
[Obtaining a service account](https://docs.gspread.org/en/latest/oauth2.html#service-account)

```shell
$ python setup.py install xls_tools[gsheet]
```

# xl_reader and xl_sheet

Moderately clever sheets for auto-detecting tabular data in spreadsheets, and manipulating it. 
"Clever" enough to get in trouble perhaps.  
