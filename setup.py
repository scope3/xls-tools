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

"""


VERSION = '0.1.1'

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
