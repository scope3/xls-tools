from setuptools import setup, find_packages

requires = [
    "xlrd",
    "google-api-python-client",
    "oauth2client"
]

# optional: pandas

VERSION = '0.1.0'

setup(
    name="xls_tools",
    version=VERSION,
    author="Brandon Kuczenski",
    author_email="brandon@scope3consulting.com",
    install_requires=requires,
    url="https://github.com/scope3/xls-tools",
    summary="Tricky tricks with XLS",
    long_description=open('README.md').read(),
    packages=find_packages()
)
