"""
No need to introduce a dependency on pandas just to import csv files.  This is a minimal implementation of the pandas
dataframe that is produced from pd.read_csv().  It includes a list of NA values that pandas automatically detects.
"""

import csv

_NA_VALUES = [  # taken from pandas._libs.parsers.py
    b'',
    b'-NaN',
    b'#N/A N/A',
    b'-1.#IND',
    b'1.#QNAN',
    b'N/A',
    b'<NA>',
    b'null',
    b'NA',
    b'NaN',
    b'nan',
    b'NULL',
    b'-nan',
    b'-1.#QNAN',
    b'#NA',
    b'1.#IND',
    b'n/a',
    b'#N/A',
]


class PandasEmulator(object):
    na_values = [k.decode('utf8') for k in _NA_VALUES]

    @classmethod
    def read_csv(cls, csvfile, quoting=csv.QUOTE_MINIMAL, **kwargs):
        with open(csvfile, 'r') as fp:
            dr = csv.reader(fp, quoting=quoting, **kwargs)
            return cls(dr)

    def __init__(self, iter_row):
        self._header = list(next(iter_row))
        self._data = list(iter_row)

    @property
    def columns(self):
        return self._header

    def __len__(self):
        return len(self._data)

    @property
    def loc(self):
        return self._data

    def __getitem__(self, item):
        inx = self._header.index(item)
        for row in self._data:
            yield row[inx]

    @classmethod
    def isna(cls, value):
        return value is None or value in cls.na_values
