import os
import py_pears.utils as utils
import py_pears.reports.sites_report as sites_report
import py_pears.reports.staff_report as staff_report
import py_pears.reports.monthly_data_cleaning as monthly_data_cleaning
import py_pears.reports.quarterly_program_evaluation as quarterly_program_evaluation
import py_pears.reports.partnerships_entry as partnerships_entry
import py_pears.reports.coalition_survey_cleaning as coalition_survey_cleaning
import py_pears.reports.partnerships_intervention_type as partnerships_intervention_type


# Calculate the path to the root directory of this package
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

# Set directories using package defaults
EXPORT_DIR = ROOT_DIR + '/pears_exports/'
OUT_DIR = ROOT_DIR + "/reports/outputs/"
TEST_INPUTS_DIR = os.path.realpath(os.path.join(ROOT_DIR, '..')) + '/tests/test_inputs/'
TEST_INPUTS_PEARS_DIR = TEST_INPUTS_DIR + '/pears/'
TEST_INPUTS_PEARS_PREV_YEAR_DIR = TEST_INPUTS_PEARS_DIR + '/prev_year/'
TEST_INPUTS_PEARS_COA_SURVEYS_DIR = TEST_INPUTS_PEARS_DIR + '/coalition_survey_exports/'

# Set paths to external data inputs
staff_list = TEST_INPUTS_DIR + 'FY22_INEP_Staff_List.xlsx'
names_list = TEST_INPUTS_DIR + 'BABY_NAMES_IL.TXT'
unit_counties = TEST_INPUTS_DIR + 'Illinois Extension Unit Counties.xlsx'
update_notifications = TEST_INPUTS_DIR + 'Update Notifications.xlsx'

creds = utils.load_credentials()

# Run Sites Report with default inputs
sites_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR)

# Run Staff Report with default inputs
staff_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR, staff_list=staff_list)

# Run Monthly Data Cleaning with default inputs
monthly_data_cleaning.main(creds=creds,
                           export_dir=EXPORT_DIR,
                           output_dir=OUT_DIR,
                           staff_list=staff_list,
                           names_list=names_list,
                           unit_counties=unit_counties,
                           update_notifications=update_notifications)

# Run Quarterly Program Evaluation with default inputs
quarterly_program_evaluation.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR)

# Run Partnerships Entry with default inputs
partnerships_entry.main(creds=creds,
                        export_dir=EXPORT_DIR,
                        output_dir=OUT_DIR,
                        staff_list=staff_list,
                        unit_counties=unit_counties,
                        prev_year_part_export=TEST_INPUTS_PEARS_PREV_YEAR_DIR + 'Partnership_Export.xlsx')

# Run Coalition Survey Cleaning with default inputs
coalition_survey_cleaning.main(creds=creds,
                               export_dir=EXPORT_DIR,
                               output_dir=OUT_DIR,
                               coalition_surveys_dir=TEST_INPUTS_PEARS_COA_SURVEYS_DIR,
                               staff_list=staff_list,
                               unit_counties=unit_counties,
                               update_notifications=update_notifications)

# Annual Reports

# Run Partnerships Intervention Type Cleaning with default inputs
partnerships_intervention_type.main(creds=creds,
                                    export_dir=EXPORT_DIR,
                                    output_dir=OUT_DIR,
                                    staff_list=staff_list,
                                    )
