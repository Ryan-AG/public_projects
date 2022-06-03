
import os
import sys
import shutil
import logging
import pandas as pd
import pyautogui as gui
from glob import glob
from pathlib import Path
from datetime import date
from win32com import client
from openpyxl import load_workbook, Workbook
from dateutil.relativedelta import relativedelta

todays_date = date.today()

LOG_FORMAT = "%(levelname)s:%(asctime)s:%(message)s"
logging.basicConfig(
    filename="mk32_test.log",
    level=logging.DEBUG,
    format=LOG_FORMAT,
)
logger = logging.getLogger()


# PATHS

# Path to template file that will be written to
mk32_template_folder_path = Path(
    r'C:\Users\grane\Documents\Projects\Python\Excel\MK32_Report_RG3.xlsx'
)

# This folder houses both MK30 and CSM001 reports.
customer_center_folder_path = Path(
    r'C:\Users\grane\Documents\Projects\Python\Excel\Customer Center//'
)

maxim_report_folder_path = Path(
    r'C:\Users\grane\Documents\Projects\Python\Excel\MaximRecurringEFT report//'
)

# The temp_csm_file file will be used instead of the csv file and deleted at the end if this script.
temp_csm_file_name = Path(
    'temp_csm_file.xlsx'
)

csm_temp_folder_path = Path(
    r'C:\Users\grane\Documents\Projects\Python\Excel\temp_c3_folder//'
)

temp_csm_file = Path(
    f'{csm_temp_folder_path}\{temp_csm_file_name}'
)


def get_monday(todays_date):
    """Finds the Monday of the current week. Required to select the correct files"""
    day_index = todays_date.weekday()
    if todays_date.weekday() != 0:
        report_date = todays_date + relativedelta(days=- day_index)
        return report_date
    else:
        return todays_date


def convert_csm_to_temp(path_to_file, destination_path):
    """Converts csv file to xlsx for simplicity later"""
    logger.debug(f"Converting csm to temp file.")
    read_file = pd.read_csv(path_to_file)
    read_file.to_excel(destination_path, index=False, header=True)
    logger.debug(f"Conversion successful.")


class FilePath:
    """Generate and hold the file paths for various reports"""

    def __init__(self, folder_path, todays_date, report_type):
        self.folder_path = Path(folder_path)
        self.todays_date = todays_date
        self.monday_date = get_monday(todays_date)
        self.year = self.monday_date.year
        self.month = self.monday_date.strftime('%m')
        self.day = self.monday_date.strftime('%d')
        self.report_type = report_type.upper()

    def __repr__(self):
        return f"FilePath(folder_path=r'{self.folder_path}', todays_date={self.todays_date}, report_type={self.report_type})"

    def __str__(self):
        return f'''{self.report_type} report: 
        Date of Report: {self.monday_date}
        Workbook Path: r"{self.file_selection()}"'''

    def file_format(self):
        if self.report_type == 'CSM':
            return '-'
        elif self.report_type == 'MAXIM':
            return '_'
        else:
            return ''

    def date_string(self):
        string = f'{self.year}{self.file_format()}{self.month}{self.file_format()}{self.day}'
        return string

    def file_selection(self):
        files = glob(
            f'{self.folder_path}\{self.report_type}*{self.date_string()}*')
        return files[0]


# Set path objects for load_workbook() to pull in file.
csm = FilePath(customer_center_folder_path, todays_date, 'csm')
maxim = FilePath(maxim_report_folder_path, todays_date, 'maxim')
mk30 = FilePath(customer_center_folder_path, todays_date, 'mk30')


def main():

    # Create the temporary file that CSM data will be read from.
    convert_csm_to_temp(
        csm.file_selection(),
        temp_csm_file
    )

    # Load template workbook and worksheets that will be written to.
    mk32_template = load_workbook(mk32_template_folder_path)
    csm_mk32_worksheet = mk32_template['CSM']
    mk30_mk32_worksheet = mk32_template['PIF']
    maxim_mk32_worksheet = mk32_template['MAX_EFT']

    # Use 'read_only=True' to avoid long processing times when loading a workbook.
    csm_workbook = load_workbook(temp_csm_file, read_only=True)
    mk30_workbook = load_workbook(mk30.file_selection(), read_only=True)
    maxim_workbook = load_workbook(maxim.file_selection(), read_only=True)

    # Load all worksheets
    csm_wb_worksheet = csm_workbook.worksheets[0]

    mk32_template.close()
    csm_workbook.close()
    mk30_workbook.close()
    maxim_workbook.close()


if __name__ == '__main__':
    main()
