import pytest

import os
import openpyxl
import pandas as pd
import numpy as np

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
TEST_INPUTS_DIR = ROOT_DIR + '/test_inputs/'
TEST_INPUTS_PEARS_DIR = TEST_INPUTS_DIR + '/pears/'
TEST_INPUTS_PEARS_PREV_YEAR_DIR = TEST_INPUTS_PEARS_DIR + '/prev_year/'
TEST_INPUTS_PEARS_COA_SURVEYS_DIR = TEST_INPUTS_PEARS_DIR + '/coalition_survey_exports/'

# Set paths to external data inputs
staff_list = TEST_INPUTS_DIR + 'FY22_INEP_Staff_List.xlsx'
names_list = TEST_INPUTS_DIR + 'BABY_NAMES_IL.TXT'
unit_counties = TEST_INPUTS_DIR + 'Illinois Extension Unit Counties.xlsx'
update_notifications = TEST_INPUTS_DIR + 'Update Notifications.xlsx'

EXPECTED_OUTPUTS_DIR = ROOT_DIR + '/expected_outputs/'
ACTUAL_OUTPUTS_DIR = ROOT_DIR + '/actual_outputs/'

creds = utils.load_credentials()


# Compare Excel Workbook objects
def compare_workbooks(xlsx1, xlsx2, diff_filename):
    wb1 = openpyxl.load_workbook(xlsx1)  # use openpyxl?
    wb2 = openpyxl.load_workbook(xlsx2)
    # Return False if sheet names aren't equal
    sheets1 = wb1.sheetnames
    sheets2 = wb2.sheetnames
    if not sheets1 == sheets2:
        # print('')
        return False

    diff_dfs = {}
    # Check column labels and data for mismatches between sheets
    for sheet in sheets1:
        df1 = pd.DataFrame(wb1[sheet].values)
        df2 = pd.DataFrame(wb2[sheet].values)

        if not df1.equals(df2):
            comparison_values = df1.values == df2.values

            rows, cols = np.where(comparison_values == False)
            for item in zip(rows, cols):
                df1.iloc[item[0], item[1]] = '{} --> {}'.format(df1.iloc[item[0],
                                                                         item[1]],
                                                                df2.iloc[item[0],
                                                                         item[1]])

            diff_dfs.update({sheet: utils.first_row_to_cols(df1)})

    if diff_dfs:
        utils.write_report(file=diff_filename, report_dict=diff_dfs)
        return False
    else:
        return True


# Create a helper function for compare_workbooks_true() tests, parameterize for reports
# Delete diff files if they already exist
# Display warnings?

# Compare copies of the same workbook
def test_compare_workbooks_true():
    diff = ACTUAL_OUTPUTS_DIR + 'wb_1_wb_1_diff.xlsx'
    result = compare_workbooks(xlsx1=TEST_INPUTS_DIR + 'test_wb_1.xlsx',
                               xlsx2=EXPECTED_OUTPUTS_DIR + 'test_wb_1.xlsx',
                               diff_filename=diff)
    assert result is True
    # compare_workbooks() for test workbooks 1 and 1 should NOT export diff
    assert os.path.isfile(diff) is False


# Compare different workbooks
def test_compare_workbooks_false():
    # Compare workbooks with different data
    diff = ACTUAL_OUTPUTS_DIR + 'wb_1_wb_2_diff.xlsx'
    result_1_2 = compare_workbooks(xlsx1=TEST_INPUTS_DIR + 'test_wb_1.xlsx',
                                   xlsx2=EXPECTED_OUTPUTS_DIR + 'test_wb_2.xlsx',
                                   diff_filename=diff)
    assert result_1_2 is False
    # compare_workbooks() for test workbooks 1 and 2 should export diff
    assert os.path.isfile(diff) is True
    # Diff should contain all unequal sheets
    test_wb_1 = openpyxl.load_workbook(TEST_INPUTS_DIR + 'test_wb_1.xlsx')
    assert openpyxl.load_workbook(diff).sheetnames == test_wb_1.sheetnames

    # Compare workbooks with different sheet names
    diff = ACTUAL_OUTPUTS_DIR + 'wb_1_wb_3_diff.xlsx'
    result_1_3 = compare_workbooks(xlsx1=TEST_INPUTS_DIR + 'test_wb_1.xlsx',
                                   xlsx2=EXPECTED_OUTPUTS_DIR + 'test_wb_3.xlsx',
                                   diff_filename=diff)
    assert result_1_3 is False
    # compare_workbooks() for test workbooks 1 and 3 should NOT export diff
    assert os.path.isfile(diff) is False


def test_sites_report():
    diff = ACTUAL_OUTPUTS_DIR + 'sites_report_diff.xlsx'
    sites_report.main(creds=creds,
                      sites_export=TEST_INPUTS_PEARS_DIR + "Site_Export.xlsx",  # Generate with Faker
                      users_export=TEST_INPUTS_PEARS_DIR + "User_Export.xlsx",  # Generate with Faker
                      output_dir=ACTUAL_OUTPUTS_DIR)
    result = compare_workbooks(xlsx1=ACTUAL_OUTPUTS_DIR + 'PEARS Sites Report 2022-09.xlsx',
                               xlsx2=EXPECTED_OUTPUTS_DIR + 'PEARS Sites Report 2022-09.xlsx',  # Generate with module
                               diff_filename=ACTUAL_OUTPUTS_DIR + 'sites_report_diff.xlsx')
    assert result is True
    # Report output changed if diff exists
    assert os.path.isfile(diff) is False
