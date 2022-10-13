import os

import py_pears.utils as utils
import py_pears.reports.sites_report as sites_report
import py_pears.reports.staff_report as staff_report
import py_pears.reports.monthly_data_cleaning as monthly_data_cleaning
import py_pears.reports.quarterly_program_evaluation as quarterly_program_evaluation
import py_pears.reports.partnerships_entry as partnerships_entry
import py_pears.reports.coalition_survey_cleaning as coalition_survey_cleaning


# Calculate the path to the root directory of this package
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '..'))

PY_PEARS_DIR = ROOT_DIR + '/py_pears'  # not necessary?
EXPORT_DIR = PY_PEARS_DIR + '/pears_exports/'

TEST_INPUTS_DIR = ROOT_DIR + '/tests/test_inputs/'
TEST_INPUTS_PEARS_DIR = TEST_INPUTS_DIR + 'pears/'
TEST_INPUTS_PEARS_PREV_YEAR_DIR = TEST_INPUTS_PEARS_DIR + 'prev_year/'
TEST_COALITION_SURVEY_EXPORTS_DIR = TEST_INPUTS_PEARS_DIR + 'coalition_survey_exports/'

EXPECTED_OUTPUTS_DIR = ROOT_DIR + '/tests/expected_outputs/'

STAFF_LIST = TEST_INPUTS_DIR + 'FY23_INEP_Staff_List.xlsx'
# Set following paths to external data inputs instead of test inputs
NAMES_LIST = TEST_INPUTS_DIR + 'BABY_NAMES_IL.TXT'
UNIT_COUNTIES = TEST_INPUTS_DIR + 'Illinois Extension Unit Counties.xlsx'
UPDATE_NOTIFICATIONS = TEST_INPUTS_DIR + 'Update Notifications.xlsx'


def main(test_pears_dir=TEST_INPUTS_PEARS_DIR,
         expected_outputs_dir=EXPECTED_OUTPUTS_DIR,
         staff_list=STAFF_LIST,
         names_list=NAMES_LIST,
         unit_counties=UNIT_COUNTIES,
         update_notifications=UPDATE_NOTIFICATIONS,
         test_pears_prev_year_dir=TEST_INPUTS_PEARS_PREV_YEAR_DIR,
         test_coalition_surveys_dir=TEST_COALITION_SURVEY_EXPORTS_DIR):
    creds = utils.load_org_settings()

    sites_report.main(creds=creds,
                      sites_export=test_pears_dir + "Site_Export.xlsx",
                      users_export=test_pears_dir + "User_Export.xlsx",
                      output_dir=expected_outputs_dir)

    staff_report.main(creds=creds,
                      users_export=test_pears_dir + "User_Export.xlsx",
                      program_activities_export=test_pears_dir + "Program_Activities_Export.xlsx",
                      indirect_activities_export=test_pears_dir + "Indirect_Activity_Export.xlsx",
                      coalitions_export=test_pears_dir + "Coalition_Export.xlsx",
                      partnerships_export=test_pears_dir + "Partnership_Export.xlsx",
                      pse_site_activities_export=test_pears_dir + "PSE_Site_Activity_Export.xlsx",
                      success_stories_export=test_pears_dir + "Success_Story_Export.xlsx",
                      staff_list=staff_list,
                      output_dir=expected_outputs_dir)

    monthly_data_cleaning.main(creds=creds,
                               coalitions_export=test_pears_dir + "Coalition_Export.xlsx",
                               indirect_activities_export=test_pears_dir + "Indirect_Activity_Export.xlsx",
                               partnerships_export=test_pears_dir + "Partnership_Export.xlsx",
                               program_activities_export=test_pears_dir + "Program_Activities_Export.xlsx",
                               pse_site_activities_export=test_pears_dir + "PSE_Site_Activity_Export.xlsx",
                               staff_list=staff_list,
                               names_list=names_list,
                               unit_counties=unit_counties,
                               update_notifications=update_notifications,
                               output_dir=expected_outputs_dir)

    partnerships_entry.main(creds=creds,
                            users_export=test_pears_dir + "User_Export.xlsx",
                            sites_export=test_pears_dir + "Site_Export.xlsx",
                            program_activities_export=test_pears_dir + "Program_Activities_Export.xlsx",
                            indirect_activities_export=test_pears_dir + "Indirect_Activity_Export.xlsx",
                            partnerships_export=test_pears_dir + "Partnership_Export.xlsx",
                            staff_list=staff_list,
                            unit_counties=unit_counties,
                            prev_year_part_export=test_pears_prev_year_dir + 'Partnership_Export.xlsx',
                            output_dir=expected_outputs_dir)

    coalition_survey_cleaning.main(creds=creds,
                                   coalitions_export=test_pears_dir + "Coalition_Export.xlsx",
                                   coalition_surveys_dir=test_coalition_surveys_dir,
                                   staff_list=staff_list,
                                   unit_counties=unit_counties,
                                   update_notifications=update_notifications,
                                   output_dir=expected_outputs_dir)

    quarterly_program_evaluation.main(coalitions_export=test_pears_dir + "Coalition_Export.xlsx",
                                      indirect_activities_export=test_pears_dir + "Indirect_Activity_Export.xlsx",
                                      partnerships_export=test_pears_dir + "Partnership_Export.xlsx",
                                      program_activities_export=test_pears_dir + "Program_Activities_Export.xlsx",
                                      pse_site_activities_export=test_pears_dir + "PSE_Site_Activity_Export.xlsx",
                                      output_dir=expected_outputs_dir)


if __name__ == '__main__':
    main()
