from googleapiclient import discovery
from googleapiclient.http import HttpError
from oauth2client.service_account import ServiceAccountCredentials


class GoogleSheetError(Exception):
    pass


class GSheetCell(object):
    def __init__(self, str_value, ctype=None):
        if len(str_value) == 0:
            ctype = 0
            value = ''
        elif ctype is None:
            # detect type: either blank, number, or string
            try:
                nval = float(str_value)
            except (TypeError, ValueError):
                nval = None
            if nval is None:
                ctype = 1
                value = str_value
            else:
                ctype = 2
                value = nval
        else:
            value = str_value
            ctype = int(ctype)

        self._c = ctype
        self._v = value

    @property
    def ctype(self):
        return self._c

    @property
    def value(self):
        return self._v


class GSheetEmulator(object):
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
    Convert a numeric index into an alphabetical column label. 0 = 'A', 1 = 'B'... 26 = 'AA', 27 = 'AB', ...
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


class GoogleSheetReader(object):
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
        cred = ServiceAccountCredentials.from_json_keyfile_name(credential_file,
                                                                scopes=['https://spreadsheets.google.com/feeds'])
        self._res = discovery.build('sheets', 'v4', credentials=cred)

        self._sheet_id = sheet_id

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

    def create_sheet(self, name, **kwargs):
        kwargs['title'] = name

        body = {'requests': [
            {'addSheet':
                 {'properties': kwargs}
             }

        ]}
        req = self._res.spreadsheets().batchUpdate(spreadsheetId=self._sheet_id,
                                                   body=body)
        return req.execute()

    def write_to_sheet(self, sheet, range, data, **kwargs):
        r = '%s!%s' % (sheet, range)
        kwargs['values'] = data
        req = self._res.spreadsheets().values().update(spreadsheetId=self._sheet_id, range=r,
                                                       body=kwargs, valueInputOption='RAW')
        return req.execute()

    def write_cell(self, sheet, row, col, value, **kwargs):
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
        :param start_row: Row to begin (note: google-sheets are 1-indexed so start_row = 0 corresponds to row 1)
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
        :param start_col: Column to begin (default 0 / 'A')
        :param kwargs:
        :return:
        """
        row += 1
        data = [[k for k in values]]
        n = len(data[0])
        rn = '%s%d:%s%d' % (_colnum_to_col(start_col), row, _colnum_to_col(start_col + n - 1), row)
        self.write_to_sheet(sheet, rn, data, **kwargs)

