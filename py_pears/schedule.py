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


# Monthly Reports

# Run Sites Report with default inputs
if compare_date(day=2):
    sites_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR)

# Run Staff Report with default inputs
if compare_date(day=11):
    staff_report.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR, staff_list=staff_list)

# Run Monthly Data Cleaning with default inputs
if compare_date(day=12):
    monthly_data_cleaning.main(creds=creds,
                               export_dir=EXPORT_DIR,
                               output_dir=OUT_DIR,
                               staff_list=staff_list,
                               names_list=names_list,
                               unit_counties=unit_counties,
                               update_notifications=update_notifications)

# Run Monthly Partnerships Entry with default inputs
if compare_date(day=20):
    partnerships_entry.main(creds=creds,
                            export_dir=EXPORT_DIR,
                            output_dir=OUT_DIR,
                            staff_list=staff_list,
                            unit_counties=unit_counties,
                            prev_year_part_export=TEST_INPUTS_PEARS_PREV_YEAR_DIR + 'Partnership_Export.xlsx')

# Quarterly Reports

# Run Coalition Survey Cleaning with default inputs
if compare_date_quarterly(days=[12, 23]):
    coalition_survey_cleaning.main(creds=creds,
                                   export_dir=EXPORT_DIR,
                                   output_dir=OUT_DIR,
                                   coalition_surveys_dir=TEST_INPUTS_PEARS_COA_SURVEYS_DIR,
                                   staff_list=staff_list,
                                   unit_counties=unit_counties,
                                   update_notifications=update_notifications)

# Run Quarterly Program Evaluation with default inputs
if compare_date_quarterly(days=[13]):
    quarterly_program_evaluation.main(creds=creds, export_dir=EXPORT_DIR, output_dir=OUT_DIR)


# Annual Reports

# Run Partnerships Intervention Type Cleaning with default inputs
if compare_date(month=10, day=4):
    partnerships_intervention_type.main(creds=creds,
                                        export_dir=EXPORT_DIR,
                                        output_dir=OUT_DIR,
                                        staff_list=staff_list)
