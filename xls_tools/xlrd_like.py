"""
This module presents an xlrd-like interface to an openpyxl (e.g. excel 2007+) spreadsheet.  The supported xlrd
API is as demanded by existing MFA code, and is as follows:

XlrdLike: workbook equivalent
 .sheet_names() - return list of sheet names
 .sheet_by_name() - return an XlSheetLike given by name
 .sheet_by_index() - return an XlSheetLike given by index
 .sheets() - return a list of XlSheetLikes-- requires initializing every sheet

XlSheetLike: worksheet equivalent
 .name - the name of the sheet according to the workbook
 .nrows - int number of rows, 0-indexed
 .ncols - int number of columns, 0-indexed
 .row(n) - return a list of XlCellLike corresponding to the nth (0-indexed) row, or IndexError
 .col(k) - return a list of XlCellLike corresponding to the kth (0-indexed) column, or IndexError
 .cell(n,k) - return the nth row, kth cell, or IndexError
 .get_rows() - row iterator

XlCellLike: cell equivalent
 .ctype - int, as indicated below
 .value - native value
"""
from datetime import datetime

import openpyxl
from xlrd.biffh import (
    XL_CELL_EMPTY,  # 0
    XL_CELL_TEXT,   # 1
    XL_CELL_NUMBER, # 2
    XL_CELL_DATE,   # 3
    XL_CELL_BOOLEAN,# 4
    # XL_CELL_ERROR, # 5
    # XL_CELL_BLANK, # 6 - for use in debugging, gathering stats, etc
)


class XlrdCellLike(object):
    """
    Subclass this to change how cells are interpreted as value + type
    """
    def __init__(self, cell):
        self._cell = cell

    @property
    def value(self):
        if self.ctype == XL_CELL_TEXT:
            return str(self._cell)
        return self._cell

    @property
    def ctype(self):
        if self._cell is None:
            return XL_CELL_EMPTY
        elif isinstance(self._cell, openpyxl.compat.NUMERIC_TYPES):
            # TODO: figure out how to detect excel-style dates
            return XL_CELL_NUMBER
        elif isinstance(self._cell, bool):
            return XL_CELL_BOOLEAN
        elif isinstance(self._cell, datetime):
            return XL_CELL_DATE
        else:
            return XL_CELL_TEXT


class XlrdSheetLike(object):
    @property
    def name(self):
        raise NotImplementedError

    @property
    def nrows(self):
        raise NotImplementedError

    @property
    def ncols(self):
        raise NotImplementedError

    def row(self, row):
        raise NotImplementedError

    def col(self, col):
        raise NotImplementedError

    def cell(self, row, col):
        raise NotImplementedError

    def get_rows(self):
        raise NotImplementedError

    def row_dict(self, row):
        """
        Creates a dictionary of the nth row using the 0th row as keynames
        :param row:
        :return:
        """
        headers = [k.value for k in self.row(0)]
        return {headers[i]: k.value for i, k in enumerate(self.row(row))}



class XlrdWorkbookLike(object):
    """
     .sheet_names() - return list of sheet names
     .sheet_by_name() - return an XlSheetLike given by name
     .sheet_by_index() - return an XlSheetLike given by index
     .sheets() - return a list of XlSheetLikes-- requires initializing every sheet
    """
    def sheet_names(self):
        raise NotImplementedError

    def sheet_by_name(self, name):
        raise NotImplementedError

    def sheet_by_index(self, index):
        raise NotImplementedError

    def sheets(self):
        raise NotImplementedError

    def __getitem__(self, item):
        if isinstance(item, int):
            return self.sheet_by_index(item)
        else:
            return self.sheet_by_name(item)

    @property
    def filename(self):
        raise NotImplementedError
