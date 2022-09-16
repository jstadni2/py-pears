import os
import py_pears.utils as utils
import py_pears.reports.sites_report as sites_report
import py_pears.reports.staff_report as staff_report
import py_pears.reports.monthly_data_cleaning as monthly_data_cleaning
import py_pears.reports.quarterly_program_evaluation as quarterly_program_evaluation


# Calculate the path to the root directory of this package
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

# Set directories using package defaults
EXPORT_DIR = ROOT_DIR + '/pears_exports/'
OUT_DIR = ROOT_DIR + "/reports/outputs/"
TEST_INPUTS_DIR = os.path.realpath(os.path.join(ROOT_DIR, '..')) + '/tests/test_inputs/'

# Set paths to external data inputs
staff_list = TEST_INPUTS_DIR + 'FY22_INEP_Staff_List.xlsx'
names_list = TEST_INPUTS_DIR + 'BABY_NAMES_IL.TXT'
unit_counties = TEST_INPUTS_DIR + 'Illinois Extension Unit Counties.xlsx'
update_notifications = TEST_INPUTS_DIR + 'Update Notifications.xlsx'

creds = utils.load_credentials()

# Run Sites Report with default inputs
# sites_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR)
#
# # Run Staff Report with default inputs
# staff_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR, staff_list=staff_list)
#
# # Run Monthly Data Cleaning with default inputs
# monthly_data_cleaning.main(creds=creds,
#                            export_dir=EXPORT_DIR,
#                            output_dir=OUT_DIR,
#                            staff_list=staff_list,
#                            names_list=names_list,
#                            unit_counties=unit_counties,
#                            update_notifications=update_notifications)

# Run Quarterly Program Evaluation with default inputs
quarterly_program_evaluation.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR)
