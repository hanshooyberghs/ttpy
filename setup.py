from setuptools import setup, find_packages

setup(
    name='ttpy',
    description='package to process VTTL table tennis data',
    version='1.0',
    packages=find_packages(),
    install_requires=['pandas>=2.0','zeep','openpyxl'],
)





