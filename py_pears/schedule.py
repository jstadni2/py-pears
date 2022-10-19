import os
from datetime import date
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
PEARS_EXPORT_DIR = ROOT_DIR + '/pears_exports/'
OUT_DIR = ROOT_DIR + "/reports/outputs/"

# Set following paths to external data inputs instead of test inputs
TEST_INPUTS_DIR = os.path.realpath(os.path.join(ROOT_DIR, '..')) + '/tests/test_inputs/'
PEARS_DIR = TEST_INPUTS_DIR + '/pears/'
PEARS_PREV_YEAR_DIR = PEARS_DIR + '/prev_year/'
PEARS_COALITION_SURVEYS_DIR = PEARS_DIR + '/coalition_survey_exports/'

STAFF_LIST = TEST_INPUTS_DIR + 'FY23_INEP_Staff_List.xlsx'
NAMES_LIST = TEST_INPUTS_DIR + 'BABY_NAMES_IL.TXT'
UNIT_COUNTIES = TEST_INPUTS_DIR + 'Illinois Extension Unit Counties.xlsx'
UPDATE_NOTIFICATIONS = TEST_INPUTS_DIR + 'Update Notifications.xlsx'

# Refactor schedule using OOP?


# Compare date to the given year, month, or day
# left_date: datetime date that subsequent arguments are compared to (default: date.today())
# year: int for the year to compare left_date to (default: date.today().year)
# month: int for the month to compare left_date to (default: date.today().month)
# day: int for the day to compare left_date to (default: date.today().day)
def compare_date(left_date=date.today(),
                 year=date.today().year,
                 month=date.today().month,
                 day=date.today().day):
    return left_date == date(year, month, day)


# Compare today's date at the start of each quarter
# days: int list for the days to today's date to
def compare_date_quarterly(days):
    quarter_start_months = [1, 4, 7, 10]
    for month in quarter_start_months:
        for day in days:
            if compare_date(month=month, day=day):
                return True
    return False


def main(pears_export_dir=PEARS_EXPORT_DIR,
         prev_year_dir=PEARS_PREV_YEAR_DIR,
         coalition_surveys_dir=PEARS_COALITION_SURVEYS_DIR,
         staff_list=STAFF_LIST,
         names_list=NAMES_LIST,
         unit_counties=UNIT_COUNTIES,
         update_notifications=UPDATE_NOTIFICATIONS,
         outputs_dir=OUT_DIR):

    creds = utils.load_org_settings()

    # Monthly Reports

    # Run Sites Report with default inputs
    if compare_date(day=2):
        # Download required PEARS exports from S3
        utils.download_s3_exports(profile=creds['aws_profile'],
                                  org=creds['s3_organization'],
                                  dst=pears_export_dir,
                                  modules=['Site', 'User'])
        sites_report.main(creds=creds,
                          sites_export=pears_export_dir + "Site_Export.xlsx",
                          users_export=pears_export_dir + "User_Export.xlsx",
                          output_dir=outputs_dir)

    # Run Staff Report with default inputs
    if compare_date(day=11):
        utils.download_s3_exports(profile=creds['aws_profile'],
                                  org=creds['s3_organization'],
                                  modules=['User',
                                           'Program_Activities',
                                           'Indirect_Activity',
                                           'Coalition',
                                           'Partnership',
                                           'PSE_Site_Activity',
                                           'Success_Story'])
        staff_report.main(creds=creds,
                          users_export=pears_export_dir + "User_Export.xlsx",
                          program_activities_export=pears_export_dir + "Program_Activities_Export.xlsx",
                          indirect_activities_export=pears_export_dir + "Indirect_Activity_Export.xlsx",
                          coalitions_export=pears_export_dir + "Coalition_Export.xlsx",
                          partnerships_export=pears_export_dir + "Partnership_Export.xlsx",
                          pse_site_activities_export=pears_export_dir + "PSE_Site_Activity_Export.xlsx",
                          success_stories_export=pears_export_dir + "Success_Story_Export.xlsx",
                          staff_list=staff_list,
                          output_dir=outputs_dir)

    # Run Monthly Data Cleaning with default inputs
    if compare_date(day=12):
        utils.download_s3_exports(profile=creds['aws_profile'],
                                  org=creds['s3_organization'],
                                  modules=['Coalition',
                                           'Indirect_Activity',
                                           'Partnership',
                                           'Program_Activities',
                                           'PSE_Site_Activity'])
        monthly_data_cleaning.main(creds=creds,
                                   coalitions_export=pears_export_dir + "Coalition_Export.xlsx",
                                   indirect_activities_export=pears_export_dir + "Indirect_Activity_Export.xlsx",
                                   partnerships_export=pears_export_dir + "Partnership_Export.xlsx",
                                   program_activities_export=pears_export_dir + "Program_Activities_Export.xlsx",
                                   pse_site_activities_export=pears_export_dir + "PSE_Site_Activity_Export.xlsx",
                                   staff_list=staff_list,
                                   names_list=names_list,
                                   unit_counties=unit_counties,
                                   update_notifications=update_notifications,
                                   output_dir=outputs_dir)

    # Run Monthly Partnerships Entry with default inputs
    if compare_date(day=20):
        utils.download_s3_exports(profile=creds['aws_profile'],
                                  org=creds['s3_organization'],
                                  modules=['User',
                                           'Site',
                                           'Program_Activities',
                                           'Indirect_Activity',
                                           'Partnership'])
        partnerships_entry.main(creds=creds,
                                users_export=pears_export_dir + "User_Export.xlsx",
                                sites_export=pears_export_dir + "Site_Export.xlsx",
                                program_activities_export=pears_export_dir + "Program_Activities_Export.xlsx",
                                indirect_activities_export=pears_export_dir + "Indirect_Activity_Export.xlsx",
                                partnerships_export=pears_export_dir + "Partnership_Export.xlsx",
                                staff_list=staff_list,
                                unit_counties=unit_counties,
                                prev_year_part_export=prev_year_dir + 'Partnership_Export.xlsx',
                                output_dir=outputs_dir)

    # Quarterly Reports

    # Run Coalition Survey Cleaning with default inputs
    if compare_date_quarterly(days=[12, 23]):
        utils.download_s3_exports(profile=creds['aws_profile'],
                                  org=creds['s3_organization'],
                                  modules=['Coalition'])
        coalition_survey_cleaning.main(creds=creds,
                                       coalitions_export=pears_export_dir + "Coalition_Export.xlsx",
                                       coalition_surveys_dir=coalition_surveys_dir,
                                       staff_list=staff_list,
                                       unit_counties=unit_counties,
                                       update_notifications=update_notifications,
                                       output_dir=outputs_dir)

    # Run Quarterly Program Evaluation with default inputs
    if compare_date_quarterly(days=[13]):
        utils.download_s3_exports(profile=creds['aws_profile'],
                                  org=creds['s3_organization'],
                                  modules=['Program_Activities',
                                           'Indirect_Activity',
                                           'Coalition',
                                           'Partnership',
                                           'PSE_Site_Activity'])
        quarterly_program_evaluation.main(coalitions_export=pears_export_dir + "Coalition_Export.xlsx",
                                          indirect_activities_export=pears_export_dir + "Indirect_Activity_Export.xlsx",
                                          partnerships_export=pears_export_dir + "Partnership_Export.xlsx",
                                          program_activities_export=pears_export_dir + "Program_Activities_Export.xlsx",
                                          pse_site_activities_export=pears_export_dir + "PSE_Site_Activity_Export.xlsx",
                                          output_dir=outputs_dir)

    # Annual Reports

    # Run Partnerships Intervention Type Cleaning with default inputs
    if compare_date(month=10, day=4):
        partnerships_intervention_type.main(creds=creds,
                                            export_dir=pears_export_dir,
                                            output_dir=outputs_dir,
                                            staff_list=staff_list)

    # Annual Program Evaluation Report
    if compare_date(month=10, day=18):
        quarterly_program_evaluation.main(coalitions_export=creds['pears_prev_year'] + "Coalition_Export.xlsx",
                                          indirect_activities_export=creds['pears_prev_year'] + "Indirect_Activity_Export.xlsx",
                                          partnerships_export=creds['pears_prev_year'] + "Partnership_Export.xlsx",
                                          program_activities_export=creds['pears_prev_year'] + "Program_Activities_Export.xlsx",
                                          pse_site_activities_export=creds['pears_prev_year'] + "PSE_Site_Activity_Export.xlsx",
                                          output_dir=outputs_dir)


if __name__ == '__main__':
    main()
