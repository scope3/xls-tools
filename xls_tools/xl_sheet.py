"""
Provide powerful, flexible access to tablular data stored in an Excel sheet.

The canonical format for an Excel table is row 1 having headers; tabular content being contiguous and complete
in rows 2-x; and no non-tabular data present in the sheet.

A lot of the functionality of this file is dedicated to auto-detecting the edges and column headers for tables
with non-canonical formatting.  Lots of heuristics. Not [yet] tested.

Uses the cheap, lightweight xlrd-like class as an access layer.
"""


from .xlrd_like import XL_CELL_EMPTY, XL_CELL_TEXT, XL_CELL_NUMBER, XlrdSheetLike
from xlrd.biffh import XL_CELL_ERROR

N_OPTS = 4
(MULTI, ROW_GAPS, COL_GAPS, MATRIX) = range(N_OPTS)

def _mk_xl_opts():
    return [None] * N_OPTS


def chunks(xcol):
    """
    Finds contiguous chunks of data, identified as non-empty cells, in an iterable of xlrd cells
    Generates a list of 2-tuples (starting row, chunk length)
    :param xcol: iterable of cells
    :return:
    """
    st = None
    for i, k in enumerate(xcol):
        if k.ctype == XL_CELL_EMPTY:
            if st is not None:  # falling edge
                yield st, i - st
                st = None
        else:
            if st is None:  # rising edge
                st = i
    if st is not None:
        yield st, len(xcol) - st


def share(xcol):
    """
    Returns the fraction of the input data that is non-empty, and the first non-empty cell
    :param xcol:
    :return:
    """
    first = None
    c = d = 0
    for k in xcol:
        c += 1
        if k.ctype == XL_CELL_EMPTY:
            continue
        d += 1
        if first is None:
            first = c - 1
    return d/c, first


def _longest(xcol):
    """
    returns row, len for longest len in the iterable
    :param xcol:
    :return:
    """
    try:
        return sorted(chunks(xcol), key=lambda _x: _x[1], reverse=True)[0]
    except IndexError:
        return 0, 0


def clean_value(cell):
    if cell.ctype == XL_CELL_TEXT:
        val = cell.value.strip()
    else:
        val = cell.value
    return val


class _EmptyRow(Exception):
    pass


class XlSheet(XlrdSheetLike):
    """
    This class handles access to a single SHEET_NAME---
    Fully defined, the SHEET_NAME has a data row (default 1), data column (default 0), and a set of column headers
    """
    def _next_row_thresh(self, start=0, thresh=0.7):
        while start < self._s.nrows:
            sh, first = share(self._s.row(start))
            if sh > thresh:
                return start, first
            start += 1
        return 0, None

    def _next_col_thresh(self, start=0, thresh=0.6):
        """
        Here we monkey a little bit to ignore long header blocks over short data
        :param start:
        :param thresh:
        :return:
        """
        apparent_firstrow, apparent_start = self._next_row_thresh()
        use_thresh = (self._s.nrows - apparent_firstrow) * thresh

        if apparent_start and (apparent_start > start):
            start = apparent_start

        while start < self._s.ncols:
            row, ck = _longest(self._s.col(start))
            if ck > use_thresh:
                return start, row
            start += 1
        return 0, 0

    def _discover(self, datarow, datacol):
        """
        This is the magic part.  We use the chunks method above to identify contiguous data chunks.
        We call 'datacol' the first auto-detected column that has a single chunk accounting for over half the
        number of rows in the spreadsheet.

        We will write this as a set of naive cases and then consolidate it as/if opportunities to do so appear
        :return:
        """
        # defaults
        dr = 1
        dc = 0
        hr = 1

        if self.multi:
            # for multi, if no row gaps, we assume long data blocks unless they are specified
            # headerrow is ignored in multi case
            if self._getopt(ROW_GAPS):
                print('Not handling row gaps!')
            else:
                dc_det, dr_det = self._next_col_thresh(thresh=0.5)
                if datacol is None:
                    dc = dc_det
                else:
                    dc = int(datacol)

                if datarow is None:
                    dr = dr_det
                else:
                    dr = int(datarow)
                    hr = dr - 1

            if dr == 0:
                # second attempt: look for full row as header; next full row as data row
                #hr, dc_det = self._next_row_thresh()
                ## if dc_det > dc:
                ##    dc = dc_det

                dr, check = self._next_row_thresh(start=hr + 1)
                if check != dc:
                    print('warning: blank leading entries in first detected data row')

        else:
            # default case-- should handle strict
            # define header_row as first row whose longest stretch exceeds 80% of the spreadsheet width
            hr, dc_det = self._next_row_thresh()
            if datacol is None:
                dc = dc_det
            else:
                dc = int(datacol)

            if datarow is not None:
                dr = int(datarow)
            else:
                if self._getopt(ROW_GAPS):
                    dr, check = self._next_row_thresh(start=hr + 1)
                    if check != dc:
                        print('warning: blank leading entries in first detected data row')
                else:
                    dr = hr + 1
        self.datarow = dr
        self.datacol = dc
        self.headerrow = hr

    def _setopt(self, opt, val):
        self._opts[opt] = bool(val)
        # reset internal lastrow
        self._lr_int = None
        # recompute headers
        self._compute_headers()

    def _getopt(self, opt):
        return self._opts[opt]

    def set_option(self, option, value):
        if not isinstance(option, int):
            option = {'mu': MULTI,
                      'ro': ROW_GAPS,
                      'co': COL_GAPS,
                      'ma': MATRIX}[str(option).lower()[:2]]
        self._setopt(option, value)

    def __init__(self, sheet, strict=False, datarow=None, datacol=None, #  headerrow=None,
                 multiheader=False,
                 row_gaps=False,
                 col_gaps=False):
        """

        :param sheet:
        :param strict: [False] if true, strict-tabular defaults are assumed: datarow=1, datacol=0, multi=False,
         unless overridden at the command line.
         if false, discovery is attempted for any non-specified params
        :param datarow:
        :param datacol:
        :param multiheader:
        """
        self._s = sheet
        self._r = None
        self._lr = None
        self._lr_int = None
        self._hr = None
        self._c = None

        # don't know what I'm doing with this
        self._opts = _mk_xl_opts()
        self._setopt(MULTI, multiheader)
        self._setopt(ROW_GAPS, row_gaps)
        self._setopt(COL_GAPS, col_gaps)
        self._cached_headers = []

        if strict:
            self.datarow = datarow or 1
            self.datacol = datacol or 0
            self.headerrow = self.datarow - 1
        else:
            self._discover(datarow, datacol)

    @property
    def is_null(self):
        return self._s.nrows == 0

    @property
    def name(self):
        return self._s.name

    @property
    def nrows(self):
        return self._s.nrows

    @property
    def ncols(self):
        return self._s.ncols

    @property
    def datarow(self):
        return self._r

    @datarow.setter
    def datarow(self, row):
        self._r = int(row)

    @property
    def headerrow(self):
        if self._hr is None:
            return self.datarow - 1
        return self._hr

    @headerrow.setter
    def headerrow(self, row):
        self._hr = int(row)
        self._compute_headers()

    @property
    def headers(self):
        return self._cached_headers

    @property
    def datacol(self):
        return self._c

    @datacol.setter
    def datacol(self, col):
        self._c = int(col)
        self._lr_int = None
        self._compute_headers()

    @property
    def lastrow(self):
        if self._lr is not None:
            return self._lr
        if self._lr_int is None:
            if self._getopt(ROW_GAPS):
                # if ROW_GAPS is true, lastrow is the last row with a nonempty entry in the data column
                self._lr_int = max(i for i, k in enumerate(self._s.col(self.datacol))
                                   if k.ctype != XL_CELL_EMPTY) + 1
            else:
                # if ROW_GAPS is false: lastrow is the last row before the first empty row after the first data row
                try:
                    self._lr_int = next(i for i, k in enumerate(self._s.col(self.datacol))
                                        if i > self.datarow and k.ctype == XL_CELL_EMPTY)
                except StopIteration:
                    self._lr_int = self._s.nrows
        return self._lr_int

    @lastrow.setter
    def lastrow(self, value):
        if value is None:
            self._lr = value
            self._lr_int = None
        else:
            self._lr = min([int(value), self._s.nrows])

    @property
    def multi(self):
        return self._getopt(MULTI)

    def _header(self, i, multi, start=None):
        """
        construct header i based on multi-row header specification and current headerrow
        :param i:
        :param multi:
        :param start:
        :return:
        """
        st_m = start or 0
        st_s = start or self.headerrow

        if multi:
            val = ' '.join(str(self._s.cell(k, i).value).strip() for k in range(st_m, self.datarow)).strip()
        else:
            val = clean_value(self._s.cell(st_s, i))
        return val

    def _compute_headers(self, multi=None, start=None):
        """
        Generate a list of headers from the current configuration (slow)
        access the protected variable _headers to use the cached list (fast)

        :param multi:
        :param start: if multi is false, start is the header row. if multi is true, start is the start of the header
        :return:
        """
        if self.datacol is None:
            return
        multi = multi or self.multi
        headers = []
        for i in range(self.datacol, self._s.ncols):
            headers.append(self._header(i, multi, start))
        self._cached_headers = headers

    def _read_row(self, rownum, _make_dict=None):
        _o = []
        _empty = True
        for i, k in enumerate(self._s.row(rownum)):
            if i < self.datacol:
                continue
            if k.ctype == XL_CELL_TEXT:
                _empty = False
                _o.append(k.value.strip())
            elif k.ctype == XL_CELL_ERROR:
                _o.append('Error:%d' % k.value)
            elif k.ctype == XL_CELL_EMPTY:
                _o.append(None)
            else:
                _empty = False
                _o.append(k.value)

        if _empty and self._getopt(ROW_GAPS):
            raise _EmptyRow

        if _make_dict is not None:
            return {k: v for k, v in zip(_make_dict, _o)}
        return _o

    def get_rows(self):
        for i, row in self.gen_rows():
            yield row

    def gen_rows(self, mask=None, rowdict=False):
        """
        Blank cells have value None
        :param mask:
        :param rowdict:
        :return: generates i, row tuples, but only for [non-blank] data rows
        """
        if rowdict:
            h = self.headers
        else:
            h = None
        for i in range(self.datarow, self.lastrow):
            in_mask = i - self.datarow
            if mask is not None:
                if not mask[in_mask]:
                    continue
            try:
                yield i, self._read_row(i, _make_dict=h)
            except _EmptyRow:
                continue

    def __getitem__(self, item):
        if isinstance(item, int):
            return self._s.row(item + self.datarow)[self.datacol:self._s.ncols]
        elif isinstance(item, tuple):
            dat = self._s.row(item[0] + self.datarow)[self.datacol:self._s.ncols]
            return dat[self.find_column(item[1])].value

    def __call__(self, item):
        return self._header(int(item), self.multi)

    def find_column(self, column):
        try:
            return int(column)
        except ValueError:
            try:
                return self.headers.index(column)
            except ValueError:
                try:
                    return next(i for i, k in enumerate(self.headers) if k is not None and k.startswith(column))
                except StopIteration:
                    raise KeyError('Column %s not found' % column)

    def _find_column(self, column):
        return self.find_column(column) + self.datacol

    def row(self, row, rowdict=False):
        if rowdict:
            h = self.headers
        else:
            h = None
        return self._read_row(row + self.datarow, _make_dict=h)

    def col(self, column, mask=None):
        column = self._find_column(column)
        if mask is None:
            return self._s.col(column)[self.datarow:self.lastrow]
        else:
            dat = self._s.col(column)[self.datarow:self.lastrow]
            return [k for i, k in enumerate(dat) if mask[i]]

    def cell(self, row, col):
        return self._s.cell(row, col)

    def col_data(self, column, mask=None):
        return [clean_value(k) for k in self.col(column, mask=mask)]

    def total(self, column, mask=None):
        return sum(k.value for k in self.col(column, mask=mask) if k.ctype == XL_CELL_NUMBER)

    def unique(self, *columns, mask=None):
        if len(columns) == 1:
            try:
                return sorted(set(self.col_data(columns[0], mask=mask)))
            except TypeError:
                return set(self.col_data(columns[0], mask=mask))
        else:
            try:
                return sorted(set(zip(*(self.col_data(column, mask=mask) for column in columns))))  # that's some "pythonic" notation
            except TypeError:
                return set(zip(*(self.col_data(column, mask=mask) for column in columns)))

    def to_dataframe(self, mask=None, **kwargs):
        import pandas as pd
        return pd.DataFrame({k: self.col_data(i, mask=mask) for i, k in enumerate(self.headers)}, **kwargs)


