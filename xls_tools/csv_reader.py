import os

try:
    # don't know how to test this in CI both with and without pandas (other than just test; pip install pandas; test again)
    import pandas as pd
except ImportError:
    from .pd_emulator import PandasEmulator as pd


from .xlrd_like import XlrdCellLike, XlrdSheetLike, XlrdWorkbookLike


def _make_cell(val):
    if pd.isna(val):
        return XlrdCellLike(None)
    return XlrdCellLike(val)


class CsvSheet(XlrdSheetLike):
    """
    Import a CSV and interact with it like an xlrd workbook containing only one sheet.

    Uses pandas to import the csv if pandas is available; otherwise, uses a minimal pandas DataFrame emulator.
    """
    def __init__(self, csvfile, **kwargs):
        name, ext = os.path.splitext(os.path.basename(csvfile))
        if ext.lower() != '.csv':
            print('Does not appear to be a csv: %s' % ext)
        self._name = name
        self._df = pd.read_csv(csvfile, **kwargs)

        self._headers = list(self._df.columns)

    def _find_column(self, column):
        """
        if number,
        :param column:
        :return:
        """
        if isinstance(column, int):
            return self._headers[column]
        elif column in self._headers:
            return column
        raise KeyError(column)

    @property
    def name(self):
        return self._name

    @property
    def nrows(self):
        return len(self._df) + 1

    @property
    def ncols(self):
        return len(self._headers)

    def row(self, row):
        if row == 0:
            return [_make_cell(k) for k in self._headers]
        else:
            return [_make_cell(c) for c in self._df.loc[row - 1]]

    def col(self, col):
        f = self._find_column(col)
        return [_make_cell(f)] + [_make_cell(c) for c in self._df[f]]

    def cell(self, row, col):
        return self.row(row)[col]

    def get_rows(self):
        for i in range(self.nrows):
            yield self.row(i)


class CsvWorkbook(XlrdWorkbookLike):
    def __init__(self, csvfile, **kwargs):
        self._csvfile = csvfile
        self._xl = CsvSheet(csvfile, **kwargs)

    @property
    def filename(self):
        return os.path.basename(self._csvfile)

    def sheet_names(self):
        return [self._xl.name]

    def sheet_by_index(self, index):
        if index == 0:
            return self._xl
        raise IndexError

    def sheet_by_name(self, name):
        if name == self._xl.name:
            return self._xl
        raise KeyError

    def sheets(self):
        return [self._xl]

    def __getitem__(self, item):
        if isinstance(item, int):
            return self.sheet_by_index(item)
        return self.sheet_by_name(item)
