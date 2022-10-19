# py-pears: unofficial PEARS development kit

[![Build Status](https://github.com/jstadni2/py-pears/workflows/test/badge.svg?branch=master&event=push)](https://github.com/jstadni2/py-pears/actions?query=workflow%3Atest)
[![codecov](https://codecov.io/gh/jstadni2/py-pears/branch/master/graph/badge.svg)](https://codecov.io/gh/jstadni2/py-pears)
[![Python Version](https://img.shields.io/pypi/pyversions/py-pears.svg)](https://pypi.org/project/py-pears/)
[![wemake-python-styleguide](https://img.shields.io/badge/style-wemake-000000.svg)](https://github.com/wemake-services/wemake-python-styleguide)

`py-pears` was developed to consolidate [Illinois Extension's](https://extension.illinois.edu/) reporting and data 
cleaning infrastructure for [PEARS](https://www.k-state.edu/oeie/pears/) into a single Python package. 

## Features

- [schedule.py](https://github.com/jstadni2/py-pears/blob/master/py_pears/schedule.py) serves as the entry point for a 
job scheduler. Operation and recommendations are detailed in [Schedule](#schedule).
- The [utils.py](https://github.com/jstadni2/py-pears/blob/master/py_pears/utils.py) module compiles methods shared 
across multiple scripts to streamline report development and maintenance.
- A brief summary of each report is provided in the [Reports](#reports) section. 
- Several modules are provided to facilitate automated testing of PEARS reports. See [Testing](#testing) below for
more information.

## Installation

The recommended way to install `py-pears` is through git, which can be downloaded 
[here](https://git-scm.com/downloads). Once downloaded, run the following command:

```bash
git clone https://github.com/jstadni2/py-pears
```

This package uses [Poetry](https://python-poetry.org/docs/) for dependency management, so follow the installation 
instructions given in the previous link.

### Setup

A JSON file of organizational settings is required to utilize `py-pears`. Create a file named `org_settings.json`
in the [/py_pears](https://github.com/jstadni2/py-pears/tree/master/py_pears) directory. 

Example `org_settings.json`:

```json
{
  "aws_profile": "your_profile_name",
  "s3_organization": "org_prefix",
  "admin_username": "your_username@domain.com",
  "admin_password": "your_password",
  "admin_send_from": "your_username@domain.com",
  "staff_list": "/path/to/Staff_List.xlsx",
  "pears_prev_year": "/path/to/annual_pears_exports/2022/",
  "coalition_survey_exports": "/path/to/coalition_survey_exports/2022/"
}
```

This package's [.gitignore](https://github.com/jstadni2/py-pears/blob/master/.gitignore#L220) file will exclude
`org_settings.json` from git commits. Follow the instructions below for obtaining necessary credentials.

#### Amazon Web Services 

An AWS named profile will need to be created for accessing automated PEARS exports from the organization's 
[AWS S3](https://aws.amazon.com/s3/) bucket. 
- Contact [PEARS support](mailto:support@pears.io) to set up an AWS S3 bucket to store 
automated PEARS exports.
- Obtain the key, secret, and organization's S3 prefix from PEARS support.
- Install [AWS CLI](https://docs.aws.amazon.com/cli/latest/userguide/getting-started-install.html).
- Use AWS CLI to [create a named profile](https://docs.aws.amazon.com/cli/latest/userguide/cli-configure-profiles.html) 
for the PEARS S3 credentials using the following command:

```bash
aws configure --profile your_profile_name
```

- Set the value of `"aws_profile"` to the name of the profile in `org_settings.json`.
- Set the value of `"s3_organization"` to the S3 prefix obtained from PEARS support.

#### Email Credentials 
Administrative credentials are required for email delivery of reports and PEARS user notifications.
- Set the `"admin_username"` and `"admin_password"` variables in `org_settings.json` to valid Office 365 credentials.
- The `"admin_send_from"` variable can be optionally set to a different address linked to `"admin_username"`. Otherwise, 
assign the same value to both variables.
- The `send_mail()` function in [utils.py](https://github.com/jstadni2/py-pears/blob/master/py_pears/utils.py#L332) is
defined using Office 365 as the host. Change the host to the appropriate email service provider if necessary.

#### External Data
The following file/directory paths are required to run some reports in `py-pears`. 
- `"staff_list"`: The path to a workbook that compiles organizational staff. 
  - See [FY23_INEP_Staff_List.xlsx](https://github.com/jstadni2/py-pears/blob/master/tests/test_inputs/FY23_INEP_Staff_List.xlsx)
  as an example.
  - Reports dependent on `"staff_list"` may require additional alterations depending on the staff list format.
  - If your organization actively maintains its staff list internally in PEARS, the 
  [User_Export.xlsx](https://github.com/jstadni2/py-pears/blob/master/tests/test_inputs/pears/User_Export.xlsx) workbook 
  could be used in lieu of external staff lists.
- `"pears_prev_year"`: The path to a directory of the previous report year's PEARS exports for each module.
  - This may not be necessary if your organization does not intent to use the 
  [Partnerships Entry Report](https://github.com/jstadni2/py-pears/blob/master/py_pears/reports/partnerships_entry.py)
- `"coalition_survey_exports"`: The path to a directory of PEARS Coalition Survey exports.
  - This may not be necessary if your organization does not intent to use the 
    [Coalition Survey Cleaning Report](https://github.com/jstadni2/py-pears/blob/master/py_pears/reports/coalition_survey_cleaning.py)

## Schedule

The run dates, input and output directories, and email recipients for each report are set in
[schedule.py](https://github.com/jstadni2/py-pears/blob/master/py_pears/schedule.py). Scheduled dates are compared to a 
timestamp before importing PEARS data from the AWS S3 and running the report. To run the schedule, execute the following
system command within the package directory:

```bash
poetry run schedule
```

Trigger dates for your organization's job scheduler should mirror the run dates set in `schedule.py`.

## Reports

### Monthly Data Cleaning

The [Monthly Data Cleaning](https://github.com/jstadni2/py-pears/blob/master/py_pears/reports/monthly_data_cleaning.py) 
script flags records based on guidance provided to PEARS users by the Illinois 
[SNAP-Ed](https://www.fns.usda.gov/snap/snap-ed) implementing agency. Users are notified via email how to update their 
flagged records.

### Staff Report

The [Staff Report](https://github.com/jstadni2/py-pears/blob/master/py_pears/reports/staff_report.py) summarizes the 
PEARS activity of SNAP-Ed staff on a monthly basis. Separate reports are generated for each Illinois SNAP-Ed
implementing agency, [Illinois Extension](https://inep.extension.illinois.edu/) and 
[Chicago Partnership for Health Promotion \(CPHP\)](https://cphp.uic.edu/).

### Quarterly Program Evaluation Report

The [Quarterly Program Evaluation Report](https://github.com/jstadni2/py-pears/blob/master/py_pears/reports/quarterly_program_evaluation.py) 
generates metrics for Illinois Extension's quarterly SNAP-Ed evaluation report. Data from PEARS is used to calculate 
evaluation metrics specified by the [SNAP-Ed Evaluation Framework](https://snapedtoolkit.org/framework/index/) and 
[Illinois Department of Human Services \(IDHS\)](https://www.dhs.state.il.us/page.aspx).

### Sites Report

The [Sites Report](https://github.com/jstadni2/py-pears/blob/master/py_pears/reports/sites_report.py) compiles the site 
records created in PEARS by Illinois Extension staff on a monthly basis. In order to prevent site duplication, select 
staff are authorized to manage requests for new site records. Other users are notified when they enter sites into PEARS 
without permission.

### Partnerships Entry Report

The [Partnerships Entry Report](https://github.com/jstadni2/py-pears/blob/master/py_pears/reports/partnerships_entry.py)
generates Partnerships to enter for the current report year. Program Activity and Indirect Activity records are 
cross-referenced with existing Partnerships to create new Partnership or copy-forward records from the previous report 
year. Separate reports are generated for each Illinois SNAP-Ed implementing agency, Illinois Extension and CPHP.

### Coalition Survey Cleaning

The [Coalition Survey Cleaning](https://github.com/jstadni2/py-pears/blob/master/py_pears/reports/coalition_survey_cleaning.py) 
script flags Coalition records if a corresponding Coalition Survey is not submitted for the previous quarter. Users are 
notified via email how to submit a survey for the applicable Coalitions.

### Partnerships Intervention Type Report 

## Testing

Since the schema of PEARS export workbooks changes periodically, `py-pears` includes several modules to enable
automated testing of exports and report outputs.

### Test PEARS

The [Test PEARS](https://github.com/jstadni2/py-pears/blob/master/tests/test_pears.py) test suites determine whether 
expected PEARS exports are present on the AWS S3. The schema of the current export workbooks are also compared to those found in
[/tests/test_inputs](https://github.com/jstadni2/py-pears/tree/master/tests/test_inputs).

Execute the following command from the root directory of the package to run `test_reports.py`:

```bash
poetry run pytest tests/test_pears.py
```

Alternatively, you can run all test suites simply via:

```bash
poetry run pytest
```

### Generate Test Inputs

The [Generate Test Inputs](https://github.com/jstadni2/py-pears/blob/master/tests/generate_test_inputs.py) script
downloads PEARS exports from the current day's AWS S3 subdirectory to 
[/tests/test_inputs](https://github.com/jstadni2/py-pears/tree/master/tests/test_inputs). Identifying information for users, sites, and 
partnering organizations is replaced with data generated from the [Faker](https://faker.readthedocs.io/en/master/)
Python package. A copy of the `"staff_list"` Excel workbook specified in `org_settings.json` is populated with fake 
users. Fields used for Illinois Extension's program evaluation are also replaced with random numeric values. Once
schema changes in PEARS export works are discovered and report scripts are updated accordingly, rerun 
`generate_test_inputs.py` and the subsequent modules and test suites.

Execute the following command to run `generate_test_inputs.py`:

```bash
poetry run generate_test_inputs
```

### Generate Expected Outputs

The [Generate Expected Outputs](https://github.com/jstadni2/py-pears/blob/master/tests/generate_expected_outputs.py) 
script runs reports with data produced by `generate_test_inputs.py`. The resulting Excel workbooks are stored in 
[/tests/actual_outputs](https://github.com/jstadni2/py-pears/tree/master/tests/actual_outputs) for use in the
[Test Reports](#test-reports) test suites.

Execute the following command to run `generate_expected_outputs.py`:

```bash
poetry run generate_expected_outputs
```

### Test Reports

The [Test Reports](https://github.com/jstadni2/py-pears/blob/master/tests/test_reports.py) test suites compare report
outputs with the Excel workbooks generated from `generate_expected_outputs.py` using the 
[pytest](https://docs.pytest.org/en/7.1.x/) framework. Any report output alterations introduced during refactoring are
detailed in diff Excel workbooks exported to
[/tests/actual_outputs](https://github.com/jstadni2/py-pears/tree/master/tests/actual_outputs).

Execute the following command from the root directory of the package to run `test_reports.py`:

```bash
poetry run pytest tests/test_reports.py
```





## License

[MIT](https://github.com/jstadni2/py-pears/blob/master/LICENSE)


## Credits

This project was generated with [`wemake-python-package`](https://github.com/wemake-services/wemake-python-package). Current template version is: [ffbf87a961dab34c346b27d0d8468fc90c215646](https://github.com/wemake-services/wemake-python-package/tree/ffbf87a961dab34c346b27d0d8468fc90c215646). See what is [updated](https://github.com/wemake-services/wemake-python-package/compare/ffbf87a961dab34c346b27d0d8468fc90c215646...master) since then.
