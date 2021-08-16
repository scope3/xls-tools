import xlrd
from .openpyxlrd import OpenpyXlrdWorkbook


def open_xl(path, formatting_info=False, data_only=True, **kwargs):
    if path.lower().endswith('xls'):
        return xlrd.open_workbook(path, formatting_info=formatting_info)
    else:
        '''
        try:
        except:
        '''
        return OpenpyXlrdWorkbook.from_file(path, data_only=data_only, **kwargs)
