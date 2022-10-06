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
TEST_OUTPUTS_DIR = ROOT_DIR + '/test_outputs/'

creds = utils.load_credentials()


# Compare Excel Workbook objects
def compare_workbooks(xlsx1, xlsx2, diff_output_dir):
    wb1 = openpyxl.load_workbook(xlsx1)  # use openpyxl?
    wb2 = openpyxl.load_workbook(xlsx2)
    # Return False if sheet names aren't equal
    sheets1 = wb1.get_sheet_names()
    sheets2 = wb2.get_sheet_names()
    if not sheets1 == sheets2:
        # print('')
        return False
    with pd.ExcelWriter(diff_output_dir + 'Excel_diff.xlsx') as writer:
        for sheet in sheets1:
            # check if sheet is in the other Excel file
            df1 = wb1[sheet]
            df2 = wb2[sheet]
            if df1 == df2:
                continue
            else:
                comparison_values = df1.values == df2.values

                # print(comparison_values)

                rows, cols = np.where(comparison_values == False)
                for item in zip(rows, cols):
                    df1.iloc[item[0], item[1]] = '{} --> {}'.format(df1.iloc[item[0],
                                                                             item[1]],
                                                                    df2.iloc[item[0],
                                                                             item[1]])

                df1.to_excel(writer, sheet_name=sheet, index=False, header=True)
                # Print all unequal sheets instead of just the first one
                return False
        return True


def test_sites_report():
    result = compare_workbooks(xlsx1=EXPECTED_OUTPUTS_DIR + 'PEARS Sites Report 2022-09.xlsx',
                               xlsx2=TEST_OUTPUTS_DIR + 'PEARS Sites Report 2022-09.xlsx',
                               diff_output_dir=TEST_OUTPUTS_DIR)
    assert result is True
