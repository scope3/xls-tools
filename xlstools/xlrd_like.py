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
import abc
from datetime import datetime

import openpyxl
from xlrd.biffh import (
    XL_CELL_EMPTY,    # 0
    XL_CELL_TEXT,     # 1
    XL_CELL_NUMBER,   # 2
    XL_CELL_DATE,     # 3
    XL_CELL_BOOLEAN,  # 4
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
        TODO: need to assign this capability to native xlrd sheets, which don't have it
        :param row:
        :return:
        """
        headers = [k.value for k in self.row(0)]
        return {headers[i]: k.value for i, k in enumerate(self.row(row)[:len(headers)])}


class XlrdWorkbookLike(abc.ABC):
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

    def __contains__(self, item):
        raise NotImplementedError

    def __getitem__(self, item):
        if isinstance(item, int):
            return self.sheet_by_index(item)
        else:
            return self.sheet_by_name(item)

    @property
    def filename(self):
        raise NotImplementedError


class XlrdWriteWorkbook(XlrdWorkbookLike, abc.ABC):
    """
     .create_sheet()
     .write_cell()
     .write_row()
     .write_column()
     .write_rectangle_by_rows()
     .clear_region()

     These together enable the super-useful write_dataframe_to_sheet()
    """
    def create_sheet(self, sheetname, **kwargs):
        raise NotImplementedError

    def write_cell(self, sheet, row, col, value, **kwargs):
        raise NotImplementedError

    def write_row(self, sheet, row, values, start_col=0, **kwargs):
        raise NotImplementedError

    def write_col(self, sheet, col, values, start_row=0, **kwargs):
        raise NotImplementedError

    def write_rectangle_by_rows(self, sheet, row_gen, start_row=0, start_col=0, **kwargs):
        raise NotImplementedError

    def clear_region(self, sheet, start_row=0, start_col=0, end_row=None, end_col=None, **kwargs):
        raise NotImplementedError

    def write_dataframe(self, sheetname, df, clear_sheet=True, write_header=True, header_levels=None,
                        fillna='NA', write_index=True):
        """

        :param self: a GoogleSheetReader
        :param sheetname: sheet to write to or create
        :param df: a pandas dataframe
        :param clear_sheet: [True]
        :param write_header: [True] whether to write header (False: leave it standing)
        :param header_levels: number of header levels to write. Must be <= nlevels
        :param fillna:
        :param write_index:
        :return:
        """

        ncol = len(df.columns)
        if not write_index:
            ncol -= 1
        if header_levels is None or header_levels > df.columns.nlevels:
            header_levels = df.columns.nlevels

        if sheetname in self.sheet_names():
            # start by clearing the sheet- with or without headers
            if clear_sheet:
                if write_header:
                    self.clear_region(sheetname)
                else:
                    self.clear_region(sheetname, start_row=header_levels)
            else:
                if write_header:
                    self.clear_region(sheetname, end_col=ncol, end_row=header_levels - 1)
        else:
            self.create_sheet(sheetname)

        # then populate
        def _row_gen(_df):
            for _i, row in _df.fillna(fillna).iterrows():
                if write_index:
                    yield [_i] + list(row.values)
                else:
                    yield list(row.values)

        if write_header:
            for i in range(header_levels):
                if write_index:
                    h = [''] + list(df.columns.get_level_values(i))
                else:
                    h = list(df.columns.get_level_values(i))
                self.write_row(sheetname, i, h)
            if df.index.name is not None:
                print('index names not handled')

        self.write_rectangle_by_rows(sheetname, _row_gen(df), start_row=header_levels)
