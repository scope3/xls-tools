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

XlCellLike: cell equivalent
 .ctype - int, as indicated below
 .value - native value
"""

import xlrd
import openpyxl


(
    XL_CELL_EMPTY,
    XL_CELL_TEXT,
    XL_CELL_NUMBER,
    XL_CELL_DATE,
    XL_CELL_BOOLEAN,
    XL_CELL_ERROR,
    XL_CELL_BLANK, # for use in debugging, gathering stats, etc
) = range(7)


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



class OpenpyxlSheetLike(XlrdSheetLike):

    @property
    def sheet(self):
        """
        Allow native access to sheet
        :return:
        """
        return self._xlsx

    def __init__(self, xlsx_sheet):
        self._xlsx = xlsx_sheet
        self._nrows = self._xlsx.max_row
        self._ncols = self._xlsx.max_column

    @property
    def name(self):
        return self._xlsx.title

    @property
    def ncols(self):
        return self._ncols

    @property
    def nrows(self):
        return self._nrows

    def row(self, row):
        """
        zero-indexed!
        :param row:
        :return:
        """
        row += 1
        if row > self._nrows:
            raise IndexError

        rows = list(self._xlsx.iter_rows(min_row=row, max_row=row))  # 2nd order list)
        return [XlrdCellLike(k.value) for k in rows[0]]

    def col(self, col):
        """
        Zero-indexed!
        :param col:
        :return:
        """
        '''
        This is somewhat DRY
        '''
        col += 1
        if col > self._ncols:
            raise IndexError

        cols = list(self._xlsx.iter_cols(min_col=col, max_col=col))
        return [XlrdCellLike(k.value) for k in cols[0]]

    def cell(self, row, col):
        row += 1
        col += 1
        if row > self._nrows or col > self._ncols:
            raise IndexError

        cell = self._xlsx.cell(row, col)
        return XlrdCellLike(cell.value)


class OpenpyXlrdWorkbook(XlrdWorkbookLike):

    @classmethod
    def from_file(cls, file, **kwargs):
        return cls(openpyxl.load_workbook(file, **kwargs))

    @property
    def book(self):
        """
        Allow native access to book
        :return:
        """
        return self._book

    def __init__(self, xl_book):
        """
        :param xl_book: an initialized Openpyxl Workbook
        """
        self._book = xl_book
        self._names = {k: i for i, k in enumerate(self._book.sheetnames)}
        self._sheets = [OpenpyxlSheetLike(self._book[k]) for k in self._book.sheetnames]

    def sheet_names(self):
        return self._book.sheetnames

    def sheet_by_name(self, name):
        return self._sheets[self._names[name]]

    def sheet_by_index(self, index):
        return self._sheets[index]

    def sheets(self):
        return self._sheets


def open_xl(path, formatting_info=False, **kwargs):
    if path.lower().endswith('xls'):
        return xlrd.open_workbook(path, formatting_info=formatting_info)
    else:
        '''
        try:
        except:
        '''
        return OpenpyXlrdWorkbook.from_file(path, **kwargs)
