"""
The purpose of these tools is to normalize access to tabular data in excel format.

The XlSheet class takes in a single sheet along with specifications for how to interpret the data-- key parameters
are the location of the data region (first row and first column of data).  Other parameters deal with specialty
aspects of interpreting header columns, and possibly data but for now that is ad hoc.

The XlSheet class provides some useful functions:
 - return headers (column names)
 - return record (row) as a list
 - return column as a list
 - return total of numeric columns
 -

For the moment, the
"""


import os

from .open_xl import open_xl
from .xlrd_like import XlrdWorkbookLike
from .xl_sheet import XlSheet


class XlReader(XlrdWorkbookLike):
    """
    How many times and in how many variants has this class been created?
    """
    def _get_sheet_index(self, sheet):
        if isinstance(sheet, int):
            return sheet
        try:
            inx = self._xl.sheet_names().index(sheet)
        except IndexError:
            try:
                inx = next(i for i, k in enumerate(self._xl.sheet_names()) if k.startswith(sheet))
            except StopIteration:
                raise KeyError('Sheet not found %s' % sheet)
        return inx

    def select_sheet(self, sheet):
        return self.__getitem__(sheet)

    def _check_xl_sheet(self, inx):
        if not isinstance(self._sheets[inx], XlSheet):
            self._sheets[inx] = XlSheet(self._xl.sheet_by_index(inx), **self._args)
        return self._sheets[inx]

    def __getitem__(self, item):
        inx = self._get_sheet_index(item)
        if inx is None:
            raise KeyError
        return self._check_xl_sheet(inx)

    def __init__(self, xlfile, formatting_info=False, **kwargs):
        """
        Open an Xl file for tabular data access
        :param xlfile: an XlrdWorkbookLike or a filename
        :param formatting_info: whether to open the spreadsheet with formatting (not implemented upstream for XLSX)
        :param kwargs: defaults to get passed to every XlSheet
        """
        self._args = kwargs
        if isinstance(xlfile, XlrdWorkbookLike):
            self._xl = xlfile
            self._fname = xlfile.filename
        else:
            self._xl = open_xl(xlfile, formatting_info=formatting_info)
            self._fname = os.path.abspath(xlfile)

        self._sheets = [None] * len(self._xl.sheet_names())

    @property
    def filepath(self):
        return self._fname

    @property
    def filename(self):
        return os.path.basename(self._fname)

    def __len__(self):
        return len(self._xl.sheets())

    def gen_rows(self, sheet=None):
        sh = self.__getitem__(sheet)
        return sh.gen_rows()

    @property
    def sheet_names(self):
        return self._xl.sheet_names()

    def sheet_by_name(self, name):
        return self.__getitem__(name)

    def sheet_by_index(self, index):
        return self.__getitem__(index)

    def sheets(self):
        return [self.__getitem__(k) for k in self.sheet_names]
