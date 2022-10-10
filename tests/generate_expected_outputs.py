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
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '..'))

PY_PEARS_DIR = ROOT_DIR + '/py_pears'  # not necessary?
EXPORT_DIR = PY_PEARS_DIR + '/pears_exports/'

TEST_INPUTS_DIR = ROOT_DIR + '/tests/test_inputs/'
TEST_INPUTS_PEARS_DIR = TEST_INPUTS_DIR + 'pears/'
TEST_INPUTS_PEARS_PREV_YEAR_DIR = TEST_INPUTS_PEARS_DIR + 'prev_year/'
TEST_COALITION_SURVEY_EXPORTS_DIR = TEST_INPUTS_PEARS_DIR + 'coalition_survey_exports/'

EXPECTED_OUTPUTS_DIR = ROOT_DIR + '/tests/expected_outputs/'

# Set paths to external data inputs
staff_list = TEST_INPUTS_DIR + 'FY23_INEP_Staff_List.xlsx'
names_list = TEST_INPUTS_DIR + 'BABY_NAMES_IL.TXT'
unit_counties = TEST_INPUTS_DIR + 'Illinois Extension Unit Counties.xlsx'
update_notifications = TEST_INPUTS_DIR + 'Update Notifications.xlsx'


creds = utils.load_credentials()

sites_report.main(creds=creds,
                  sites_export=TEST_INPUTS_PEARS_DIR + "Site_Export.xlsx",
                  users_export=TEST_INPUTS_PEARS_DIR + "User_Export.xlsx",
                  output_dir=EXPECTED_OUTPUTS_DIR)

staff_report.main(creds=creds,
                  users_export=TEST_INPUTS_PEARS_DIR + "User_Export.xlsx",
                  program_activities_export=TEST_INPUTS_PEARS_DIR + "Program_Activities_Export.xlsx",
                  indirect_activities_export=TEST_INPUTS_PEARS_DIR + "Indirect_Activity_Export.xlsx",
                  coalitions_export=TEST_INPUTS_PEARS_DIR + "Coalition_Export.xlsx",
                  partnerships_export=TEST_INPUTS_PEARS_DIR + "Partnership_Export.xlsx",
                  pse_site_activities_export=TEST_INPUTS_PEARS_DIR + "PSE_Site_Activity_Export.xlsx",
                  success_stories_export=TEST_INPUTS_PEARS_DIR + "Success_Story_Export.xlsx",
                  staff_list=staff_list,
                  output_dir=EXPECTED_OUTPUTS_DIR)

monthly_data_cleaning.main(creds=creds,
                           coalitions_export=TEST_INPUTS_PEARS_DIR + "Coalition_Export.xlsx",
                           indirect_activities_export=TEST_INPUTS_PEARS_DIR + "Indirect_Activity_Export.xlsx",
                           partnerships_export=TEST_INPUTS_PEARS_DIR + "Partnership_Export.xlsx",
                           program_activities_export=TEST_INPUTS_PEARS_DIR + "Program_Activities_Export.xlsx",
                           pse_site_activities_export=TEST_INPUTS_PEARS_DIR + "PSE_Site_Activity_Export.xlsx",
                           staff_list=staff_list,
                           names_list=names_list,
                           unit_counties=unit_counties,
                           update_notifications=update_notifications,
                           output_dir=EXPECTED_OUTPUTS_DIR)

partnerships_entry.main(creds=creds,
                        users_export=TEST_INPUTS_PEARS_DIR + "User_Export.xlsx",
                        sites_export=TEST_INPUTS_PEARS_DIR + "Site_Export.xlsx",
                        program_activities_export=TEST_INPUTS_PEARS_DIR + "Program_Activities_Export.xlsx",
                        indirect_activities_export=TEST_INPUTS_PEARS_DIR + "Indirect_Activity_Export.xlsx",
                        partnerships_export=TEST_INPUTS_PEARS_DIR + "Partnership_Export.xlsx",
                        staff_list=staff_list,
                        unit_counties=unit_counties,
                        prev_year_part_export=TEST_INPUTS_PEARS_PREV_YEAR_DIR + 'Partnership_Export.xlsx',
                        output_dir=EXPECTED_OUTPUTS_DIR)
