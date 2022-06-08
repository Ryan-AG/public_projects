from pathlib import Path

# Path to template file that will be written to.
mk32_template_folder_path = Path(
    r'C:\Users\grane\Desktop\mk32_EDD_EFT Report\MK32_Report_Template.xlsx')

# This folder houses both MK30 and CSM001 reports.
customer_center_folder_path = Path(
    r'C:\Users\grane\Desktop\mk32_EDD_EFT Report\Customer Center//')

maxim_report_folder_path = Path(
    r'C:\Users\grane\Desktop\mk32_EDD_EFT Report\MaximRecurringEFT report//')

# The temp_csm_file file will be used instead of the csv file and deleted at the end if this script.
temp_csm_file_name = Path('temp_csm_file.xlsx')

csm_temp_folder_path = Path(
    r'C:\Users\grane\Desktop\mk32_EDD_EFT Report\temp_c3_folder//')

temp_csm_file = Path(f'{csm_temp_folder_path}\{temp_csm_file_name}')

report_output_path = Path(
    r'C:\Users\grane\Desktop\mk32_EDD_EFT Report\Output//')

fds_path = r"C:\Users\grane\Desktop\mk32_EDD_EFT Report\FDS//"

report_excel_archive = r'C:\Users\grane\Desktop\mk32_EDD_EFT Report\Output\Excel Archive//'
