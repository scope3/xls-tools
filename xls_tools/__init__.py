import os
import xlrd
import re

from .google_sheet_reader import GoogleSheetReader
from .xl_reader import XlReader
from .xl_sheet import XlSheet
# from .exchanges_from_spreadsheet import exchanges_from_spreadsheet


def xl_date(cell_or_value, mode=0, short=True):
    if isinstance(cell_or_value, xlrd.sheet.Cell):
        val = cell_or_value.value
    else:
        val = cell_or_value
    if short:
        return xlrd.xldate_as_tuple(val, mode)[:3]
    else:
        return xlrd.xldate_as_tuple(val, mode)


def xls_files(_path):
    """
    Generates xls / xlsx files in a directory
    :param _path:
    :return:
    """
    for k in os.listdir(_path):
        if bool(re.match('^[^\.].+\.xlsx?$', k, flags=re.I)):
            yield os.path.join(_path, k)



