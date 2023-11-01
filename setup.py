from setuptools import setup, find_packages

setup(
    name='ttpy',
    description='package to process VTTL table tennis data',
    version='1.0',
    package_dir={'': 'ttpy'},  
    packages=find_packages(where='ttpy'),
    install_requires=['pandas>=2.0','zeep','openpyxl'],
)




