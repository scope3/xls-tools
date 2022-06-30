from googleapiclient import discovery
from googleapiclient.http import HttpError
from oauth2client.service_account import ServiceAccountCredentials
from .xlrd_like import XlrdCellLike, XlrdSheetLike, XlrdWorkbookLike
from .util import colnum_to_col

import time


class GoogleSheetError(Exception):
    pass


class GSheetCell(XlrdCellLike):
    def __init__(self, str_value):
        if len(str_value) == 0:
            value = None
        else:
            # detect type: either blank, number, or string
            try:
                value = float(str_value)
            except (TypeError, ValueError):
                value = str_value
        super(GSheetCell, self).__init__(value)


class GSheetEmulator(XlrdSheetLike):
    def __init__(self, value_data):
        """

        :param value_data: must follow the google sheets v4 API .spreadsheets().values().get(... range=sheetname)
        """
        if len(value_data['values']) == 0:
            _nr = _nc = 0
            data = [[]]
        else:
            if value_data['majorDimension'] == 'ROWS':
                data = value_data['values']
            elif value_data['majorDimension'] == 'COLUMNS':
                raise NotImplementedError
            else:
                raise GoogleSheetError('Cannot interpret value_data majorDimension')
            _nr = len(data)
            _nc = max([len(k) for k in data])

        self._range = value_data['range']

        self._nr = int(_nr)
        self._nc = int(_nc)
        self._data = data

    @property
    def name(self):
        return self._range.split('!')[0].strip('\'\"')

    @property
    def range(self):
        return self._range.split('!')[1]

    @property
    def nrows(self):
        return self._nr

    @property
    def ncols(self):
        return self._nc

    def row(self, row):
        return list(GSheetCell(k) for k in self._data[row])

    def get_rows(self):
        for i in range(self.nrows):
            yield self.row(i)

    def col(self, col):
        cd = []
        for row in range(self.nrows):
            try:
                cd.append(GSheetCell(self._data[row][col]))
            except IndexError:
                cd.append(GSheetCell(''))
        return cd

    def cell(self, row, col):
        return GSheetCell(self._data[row][col])


class GoogleSheetReader(XlrdWorkbookLike):
    """
    Creates an xlrd-like google sheet reader with the following properties:

    sheet_by_name(sheet): returns a 'sheet-like' object

    'sheet-like' object has the following minimal API:
    nrows: (property) returns int number of rows
    ncols: (property) returns int number of columns
    row(n): returns a list of 'cell-like' objects in row n (0-indexed)

    'cell-like' object has the following minimal API:
    ctype: (property) integer 0=empty, 1=text, 2=number, 3=date, 4=bool, 5=error, 6=debug
    value: (property) value of the cell

    """
    def __init__(self, credential_file, sheet_id):
        """
        Creates an Xlrd-like object that also has create-sheet and write-to-sheet capabilities.

        For instructions on obtaining a JWT credential file, please visit:
         https://docs.gspread.org/en/latest/oauth2.html#service-account

        the sheet_id is the long alphanumeric string that appears in the URL, e.g.:
        https://docs.google.com/spreadsheets/d/{sheet_id}/edit#...

        You must grant your service account authority to access the sheet.

        :param credential_file:
        :param sheet_id:
        """
        cred = ServiceAccountCredentials.from_json_keyfile_name(credential_file,
                                                                scopes=['https://spreadsheets.google.com/feeds'])
        self._res = discovery.build('sheets', 'v4', credentials=cred)

        self._sheet_id = sheet_id

        self._sheetnames = self.sheet_names()

    @property
    def filename(self):
        return self._sheet_id

    def sheet_names(self):
        req = self._res.spreadsheets().get(spreadsheetId=self._sheet_id)
        d = req.execute()
        return [k['properties']['title'] for k in d['sheets']]

    def sheet_by_name(self, sheetname):
        """
        Runs a new request every time- no caching
        :param sheetname:
        :return:
        """
        quoted_sheetname = "'%s'" % sheetname  # without quotes it may be interpreted as a named range
        req = self._res.spreadsheets().values().get(spreadsheetId=self._sheet_id, range=quoted_sheetname)
        try:
            d = req.execute()
        except HttpError:
            raise KeyError('Unable to open sheet %s' % sheetname)
        return GSheetEmulator(d)

    def sheet_by_index(self, index):
        return self.sheet_by_name(self._sheetnames[index])

    def sheets(self):
        """
        No sheet caching!
        :return:
        """
        return [self.sheet_by_name(name) for name in self._sheetnames]

    def create_sheet(self, name, **kwargs):
        kwargs['title'] = name

        body = {'requests': [
            {'addSheet':
                 {'properties': kwargs}
             }

        ]}
        req = self._res.spreadsheets().batchUpdate(spreadsheetId=self._sheet_id,
                                                   body=body)
        ret = req.execute()
        self._sheetnames = self.sheet_names()
        return ret

    def write_to_sheet(self, sheet, range, data, **kwargs):
        """
        The data must be a 2d array that matches the size of the range argument
        :param sheet:
        :param range:
        :param data:
        :param kwargs: added to request body
        :return:
        """
        r = '%s!%s' % (sheet, range)
        kwargs['values'] = data
        req = self._res.spreadsheets().values().update(spreadsheetId=self._sheet_id, range=r,
                                                       body=kwargs, valueInputOption='RAW')
        result = req.execute()
        time.sleep(1)  # standard quota is only 60 requests per minute per user (300 per minute per project)
        # use write_rectangle_by_rows and [nonimpl] write_rectangle_by_columns
        return result

    def write_cell(self, sheet, row, col, value, **kwargs):
        """

        :param sheet:
        :param row: 0-indexed
        :param col: 0-indexed number or alphabetical column name e.g. 'A'
        :param value:
        :param kwargs: added to request body
        :return:
        """
        col = colnum_to_col(col)
        data = [[value]]
        rn = '%s%d:%s%d' % (col, row + 1, col, row + 1)
        return self.write_to_sheet(sheet, rn, data, **kwargs)

    def write_column(self, sheet, col, values, start_row=0, **kwargs):
        """
        Write sequential data into a column of the google sheet.
        :param sheet:
        :param col: either a column string (e.g. 'AA') or a 0-indexed column number (e.g. 0 = 'A', 1 = 'B', ...)
        :param values:
        :param start_row: 0-indexed row to begin (note: google-sheets are 1-indexed so start_row=0 corresponds to row 1)
        :param kwargs:
        :return:
        """
        col = colnum_to_col(col)
        data = [[k] for k in values]
        n = len(data)
        rn = '%s%d:%s%d' % (col, start_row+1, col, start_row + n)
        self.write_to_sheet(sheet, rn, data, **kwargs)

    def write_row(self, sheet, row, values, start_col=0, **kwargs):
        """
        Write sequential data into a row of the google sheet.  (Note: google-sheets are 1-indexed so row = 0
        corresponds to the spreadsheet's native row 1)
        :param sheet:
        :param row:
        :param values:
        :param start_col: 0-indexed Column to begin (default 0 / 'A')
        :param kwargs:
        :return:
        """
        row += 1
        data = [[k for k in values]]
        n = len(data[0])
        rn = '%s%d:%s%d' % (colnum_to_col(start_col), row, colnum_to_col(start_col + n - 1), row)
        self.write_to_sheet(sheet, rn, data, **kwargs)

    def write_rectangle_by_rows(self, sheet, row_gen, start_row=0, start_col=0, **kwargs):
        """
        Write data to a rectangular area, starting at start_row and start_col.  The size of the rectangle
        is determined by the longest row {{short rows are padded with Nones, which are ignored by gsheet API}}

        This is vital to avoiding unbearably slow execution, due to google's rate limit of 60 queries/minute/user

        :param sheet:
        :param row_gen: a generator that produces iterables of values for each row, beginning with start_col
        :param start_row: 0-indexed start row
        :param start_col: 0-indexed start column
        :param kwargs:
        :return:
        """
        end_row = start_row
        start_row += 1
        data = []
        n = 0
        for row in row_gen:
            nextdata = [value for value in row]
            n = max([n, len(nextdata)])
            data.append(nextdata)
            end_row += 1

        for row in data:
            while len(row) < n:
                row.append(None)

        rn = '%s%d:%s%d' % (colnum_to_col(start_col), start_row, colnum_to_col(start_col + n - 1), end_row)
        self.write_to_sheet(sheet, rn, data, **kwargs)

    def clear_region(self, sheet, start_row=0, start_col=0, end_row=None, end_col=None, **kwargs):
        """
        Clear the region using the gsheet API.  Note: input args are 0-indexed, noting that API is 1-indexed.
        Default is to clear the entire sheet.
        :param sheet: must exist
        :param start_row: 0-indexed. defaults to first row
        :param start_col: 0-indexed. defaults to first column
        :param end_row: 0-indexed. defaults to last row
        :param end_col: 0-indexed. defaults to last column
        :param kwargs: passed as request body
        :return:
        """
        s = self.sheet_by_name(sheet)
        if end_row is None or end_row > (s.nrows - 1):
            end_row = s.nrows
        else:
            end_row += 1
        if end_col is None or end_col > (s.ncols - 1):
            end_col = s.ncols
        else:
            end_col += 1
        start_row = max([start_row + 1, 1])
        start_col = max([start_col + 1, 1])

        rn = '%s!R%dC%d:R%dC%d' % (sheet, start_row, start_col, end_row, end_col)
        req = self._res.spreadsheets().values().clear(spreadsheetId=self._sheet_id, range=rn, body=kwargs)
        req.execute()

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

        #then populate
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
        self.write_rectangle_by_rows(sheetname, _row_gen(df), start_row=header_levels)
