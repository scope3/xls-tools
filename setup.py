from setuptools import setup, find_packages

requires = [
    "xlrd>=2.0.0",
    "openpyxl>=3.0.7"
]

"""
Version History:

0.1.0  - 8 April 2022 - initial public release

0.1.1 - 1 August 2022 - add CSV reader (with pandas emulation!)
        change package name to xlstools
        
0.1.2 - 19 February 2023 - finish csv; add write class

0.1.3 - 21 February 2023 - dumb abc error. we should think about writing some tests :P

0.1.4 - 15 May 2024 - bugfixes: Openpyxl workbook name; gsheet empty-sheet

0.1.5 - 22 September 2024 - Get rid of google sheet nag
"""


VERSION = '0.1.5'

setup(
    name="xlstools",
    version=VERSION,
    author="Brandon Kuczenski",
    author_email="brandon@scope3consulting.com",
    install_requires=requires,
    extras_require={
        'gsheet': ["google-api-python-client>=2.2.0", "oauth2client>=4.1.3"]
    },
    url="https://github.com/scope3/xls-tools",
    summary="Tricky tricks with XLS",
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    packages=find_packages()
)
