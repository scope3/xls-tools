from googleapiclient import discovery
from googleapiclient.http import HttpError
from oauth2client.service_account import ServiceAccountCredentials
from .xlrd_like import XlrdCellLike, XlrdSheetLike, XlrdWorkbookLike


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

    def row_dict(self, row):
        """
        Creates a dictionary of the nth row using the 0th row as keynames
        :param row:
        :return:
        """
        headers = [k.value for k in self.row(0)]
        return {headers[i]: k.value for i, k in enumerate(self.row(row))}


def _col_to_colnum(col):
    """
    Convert an excel column name to a 0-indexed number 'A' = 0, 'B' = 1, ... 'AA' = 26, ... 'AAZ' = 727, ...
    :param col:
    :return:
    """
    if not isinstance(col, str):
        return int(col)
    cols = list(col.upper())
    num = 0
    while len(cols) > 0:
        num *= 26
        c = cols.pop(0)
        num += (ord(c) - 64)
    return num - 1


def _colnum_to_col(num):
    """
    Convert a 0-indexed numeric index into an alphabetical column label. 0 = 'A', 1 = 'B'... 26 = 'AA', 27 = 'AB', ...
    :param num:
    :return:
    """
    if isinstance(num, str):
        return num
    col = ''
    num = int(num)
    while 1:
        rad = num % 26
        col = chr(ord('A') + rad) + col
        num //= 26
        num -= 1
        if num < 0:
            return col


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
        req = self._res.spreadsheets().values().get(spreadsheetId=self._sheet_id, range=sheetname)
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
        return req.execute()

    def write_cell(self, sheet, row, col, value, **kwargs):
        """

        :param sheet:
        :param row: 0-indexed
        :param col: 0-indexed number or alphabetical column name e.g. 'A'
        :param value:
        :param kwargs: added to request body
        :return:
        """
        col = _colnum_to_col(col)
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
        col = _colnum_to_col(col)
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
        rn = '%s%d:%s%d' % (_colnum_to_col(start_col), row, _colnum_to_col(start_col + n - 1), row)
        self.write_to_sheet(sheet, rn, data, **kwargs)
