import os
import py_pears.utils as utils
import py_pears.reports.sites_report as sites_report
import py_pears.reports.staff_report as staff_report


# Calculate the path to the root directory of this package
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

# Set directories using package defaults
EXPORT_DIR = ROOT_DIR + '/pears_exports/'
OUT_DIR = ROOT_DIR + "/reports/outputs/"
TEST_INPUTS_DIR = os.path.realpath(os.path.join(ROOT_DIR, '..')) + '/tests/test_inputs/'

# Set paths to external data inputs
staff_list = TEST_INPUTS_DIR + 'FY22_INEP_Staff_List.xlsx'

creds = utils.load_credentials()

# Run Sites Report with default inputs
sites_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR)

# Run Sites Report with default inputs
staff_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR, staff_list=staff_list)
