import xlrd
from .openpyxlrd import OpenpyXlrdWorkbook
from .csv_reader import CsvWorkbook


def open_xl(path, formatting_info=False, data_only=True, **kwargs):
    """
    Reads XLS, XLSX, or CSV files into an object with a consistent, minimal read-only interface based on xlrd
    :param path:
    :param formatting_info:
    :param data_only:
    :param kwargs:
    :return:
    """
    if path.lower().endswith('xls'):
        return xlrd.open_workbook(path, formatting_info=formatting_info)
    elif path.lower().endswith('csv'):
        return CsvWorkbook(path, **kwargs)
    else:
        '''
        try:
        except:
        '''
        return OpenpyXlrdWorkbook.from_file(path, data_only=data_only, **kwargs)
