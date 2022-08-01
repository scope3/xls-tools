"""
TODO: import CSV
"""

import os
import xlrd
import re

from datetime import datetime

try:
    from .google_sheet_reader import GoogleSheetReader
except ImportError:
    print('Unable to import GoogleSheetReader - try python setup.py install [gsheet]')
    GoogleSheetReader = False
from .xl_reader import XlReader
from .xl_sheet import XlSheet
from .openpyxlrd import OpenpyXlrdWorkbook
from .open_xl import open_xl
from .util import colnum_to_col, col_to_colnum
# from .exchanges_from_spreadsheet import exchanges_from_spreadsheet


def xl_date(cell_or_value, mode=0, short=True):
    """
    This uses an xlrd utility function to convert excel integer dates to date tuples.
    :param cell_or_value:
    :param mode:
    :param short:
    :return:
    """
    if isinstance(cell_or_value, xlrd.sheet.Cell):
        val = cell_or_value.value
    else:
        val = cell_or_value
    if isinstance(val, datetime):
        tup = val.timetuple()
    else:
        tup = xlrd.xldate_as_tuple(val, mode)
    if short:
        return tup[:3]
    else:
        return tup


def xls_files(_path):
    """
    Generates xls / xlsx files in a directory
    :param _path:
    :return:
    """
    for k in os.listdir(_path):
        if bool(re.match('^[^\.].+\.xlsx?$', k, flags=re.I)):
            yield os.path.join(_path, k)
