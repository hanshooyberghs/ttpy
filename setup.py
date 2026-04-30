from setuptools import setup, find_packages

setup(
    name='ttpy',
    description='package to process VTTL table tennis data',
    version='1.1',
    packages=find_packages(),
    install_requires=['pandas>=2.0','zeep','openpyxl','requests','python-docx','xlsxwriter'],
    entry_points={
        'console_scripts': [
            'ttpy_trigger_download=ttpy.vttl_download:run',
            'ttpy_mailing_naam=ttpy.mailing_naam:run',
            'ttpy_mailing_clubs=ttpy.mailing_clubs:run',
            'ttpy_inschrijving_tornooi=ttpy.inschrijving_tornooi:run',
            'ttpy_mails_tornooi=ttpy.mails_tornooi:run',
            'ttpy_check=ttpy.check:run',
        ],
    },
)

