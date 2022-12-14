import os
import shutil
import boto3
import pandas as pd
import numpy as np
import json
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders


# Calculate the path to the root directory of this script
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))


# Download PEARS Exports from AWS S3
# profile: string for AWS named profile
# org: string for the organization's bucket subdirectory (eg. 'uie')
# date: string in %Y/%m/%d format for the export date (default is today's date)
# dst: string for destination directory to download PEARS exports to
# modules: list of strings for the PEARS modules to download
def download_s3_exports(profile, org, date=pd.to_datetime("today").strftime("%Y/%m/%d"),
                        dst=ROOT_DIR + "/pears_exports", modules=None):
    # Use PEARS AWS S3 credentials
    session = boto3.Session(profile_name=profile)

    # Access S3 objects uploaded the day reformatting script is run
    conn = session.client('s3')
    my_bucket = 'exports.pears.oeie.org'

    response = conn.list_objects_v2(
        Bucket=my_bucket,
        Prefix=org + '/' + date + '/',
        MaxKeys=100)

    # Create a list of filenames to download from the S3
    # Might need additional string operations (capitalization, spaces to underscores)
    # Throw exception for invalid modules
    # What happens if modules is None? > 'NoneType' object is not iterable
    # Add _Export.xlsx for all exc
    module_filenames = []
    if modules is not None:
        module_filenames = [s + '_Export.xlsx' for s in modules]

    # Download the Excel files to the destination directory
    for f in response['Contents']:
        file = f['Key']
        filename = file[file.rfind('/') + 1:]
        if modules is not None and filename not in module_filenames:
            continue
        conn.download_file(my_bucket, file, dst + '/' + filename)


# Function to convert custom field's label to its dropdown value
# text: string value of the label suffix
# custom_field_label: string for the custom field label's prefix
def replace_all(text, custom_field_label):
    dic = {
        custom_field_label: '',
        # Map label suffixes to response options
        'family_life': 'Family Life',
        'nutrition_wellness': 'Nutrition & Wellness',
        'consumer_economics': 'Consumer Economics',
        'snap_ed': 'SNAP-Ed',
        'efnep': 'EFNEP',
        'improve_diet_quality': 'Improve diet quality',
        'increase_physical_activity_opportunities': 'Increase physical activity opportunities',
        'increase_food_access': 'Increase food access',
        'none': 'None',
        'abcs_of_school_nutrition': 'ABCs of School Nutrition',
        'growing_together_illinois': 'Growing Together Illinois',
        'heat': 'HEAT',
        'cphp_shape_up_chicago_youth_trainers': 'CPHP - Shape Up Chicago Youth Trainers',
        'cphp_chicago_grows_groceries': 'CPHP - Chicago Grows Groceries',
        '_': '',
    }
    for i, j in dic.items():
        text = text.replace(i, j)
    return text


# Convert custom field value binary columns into a single custom field column of list-like strings
# df: dataframe of records to reformat
# labels: list of custom labels to iterate through
def reformat(df, labels):
    reformatted_df = df.copy()
    # Remove custom data tag from column labels
    reformatted_df.columns = reformatted_df.columns.str.replace(r'_custom_data', '')
    for label in labels:
        binary_cols = reformatted_df.columns[reformatted_df.columns.str.contains(label)]
        if binary_cols.empty:
            continue
        for col in binary_cols:
            reformatted_df.loc[reformatted_df[col] == 1, col] = replace_all(col, label)
            reformatted_df.loc[reformatted_df[col] == 0, col] = ''
        # Create custom field column of list-like strings
        reformatted_df[label] = reformatted_df[binary_cols].apply(lambda row:
                                                                  ','.join(row.values.astype(str)),
                                                                  axis=1).str.strip(',').str.replace(r',+',
                                                                                                     ',', regex=True)
        reformatted_df.loc[reformatted_df[label] == '', label] = np.nan
        # Remove custom field value binary columns
        reformatted_df.drop(columns=binary_cols, inplace=True)
    return reformatted_df


# Return the previous month
# return_type: either 'datetime', 'period', '%m', '%Y-%m' (default: 'datetime')
def previous_month(return_type='datetime'):
    prev_month = pd.to_datetime("today") - pd.DateOffset(months=1)
    if return_type == 'datetime':
        return prev_month
    elif return_type == 'period':
        return prev_month.to_period('M')
    elif return_type == '%m':
        return prev_month.strftime('%m')
    elif return_type == '%Y-%m':
        return prev_month.strftime('%Y-%m')


# Return a DataFrame of datatypes associated with the previous fiscal quarter
# columns: any of 'fq', 'fq_int', 'month', 'survey_fq'
# UPDATE: Implement a try-catch?
def previous_fq(columns='fq'):
    fq_lookup = pd.DataFrame({'month': ['12', '03', '06', '09'],
                              'fq': ['Q1', 'Q2', 'Q3', 'Q4'],
                              'fq_int': [1, 2, 3, 4],
                              'survey_fq': ['Quarter 1 (October-December)', 'Quarter 2 (January-March)',
                                            'Quarter 3 (April-June)', 'Quarter 4 (July-September)']})

    prev_month = previous_month(return_type='%m')
    if prev_month not in fq_lookup['month'].values.tolist():
        return fq_lookup.loc[fq_lookup['month'] == 'Q4', columns]
    else:
        return fq_lookup.loc[fq_lookup['month'] == prev_month, columns]


# IMPLEMENT def current_fy()


# Select records from PEARS module export
# df: dataframe of PEARS module records
# record_name_field: field label for the record name
# test_records: boolean, whether to include records with 'test' in the record_name_field (default: False)
# exclude_sites: list of strings for sites to exclude from the 'site_name' field (default: ['abc placeholder'])
# columns: list of column labels to subset df by (default: return all columns)
# UPDATE: Add input arg for program_area?
def select_pears_data(df, record_name_field, test_records=False, exclude_sites=['abc placeholder'], columns=[]):
    out_df = df.copy()
    if not test_records:
        out_df = out_df.loc[~out_df[record_name_field].str.contains('(?i)TEST', regex=True, na=False)]
    if 'site_name' in out_df.columns:
        out_df = out_df.loc[~out_df['site_name'].isin(exclude_sites)]
    if not columns:  # Refactor?
        columns = out_df.columns
    return out_df[columns]


# function for reordering comma-separated name
# df: dataframe of staff list
# name_field: column label of name field
# reordered_name_field: column label of reordered name field
# drop_substr_fields: bool for dropping name substring fields
def reorder_name(df, name_field, reordered_name_field, drop_substr_fields=False):
    out_df = df.copy(deep=True)
    out_df[name_field] = out_df[name_field].str.split(pat=', ')
    out_df['first_name'] = out_df[name_field].str[1]
    out_df['last_name'] = out_df[name_field].str[0]
    out_df[reordered_name_field] = out_df['first_name'].map(str) + ' ' + out_df['last_name'].map(str)
    if drop_substr_fields:
        out_df = out_df.drop(columns=['first_name', 'last_name'])

    return out_df


# Convert county values in the 'unit' field to units
# data: dataframe of PEARS module data
# unit_field: string for the label of the unit field (default: 'unit')
# unit_counties: dataframe of counties mapped to units (default: empty dataframe)
def counties_to_units(data, unit_field='unit', unit_counties=pd.DataFrame()):
    out_data = data.copy()
    out_data[unit_field] = out_data[unit_field].str.replace('|'.join([' \(County\)', ' \(District\)', 'Unit ']),
                                                            '', regex=True)
    out_data = pd.merge(out_data, unit_counties, how='left', left_on=unit_field, right_on='County')
    out_data.loc[(~out_data[unit_field].isin(unit_counties['Unit #'])) &
                 (out_data[unit_field].isin(unit_counties['County'])), unit_field] = out_data['Unit #']
    return out_data


# Get the update notification
# update_notes: dataframe of update notification
# module: string for the PEARS module
# update: string for the label of the update column
# notification: string for the desired notification column from update_notes (default: 'Notification1')
def get_update_note(update_notes, module, update, notification='Notification1'):
    return update_notes.loc[(update_notes['Module'] == module)
                            & (update_notes['Update'] == update), notification].item()


# Merge records to Partnerships via site_id, compute module counts
# primary_records: dataframe of records that related records will be left-joined to
# primary_id: string for the unique ID column of primary_records
# related_records: dataframe of related records that will be left-joined to primary_records
# merge_on: string for the column to join the primary and related records
# related_id: string for the unique ID column of related_records
# count_label: string for the label of the count column
def count_related_records(primary_records, primary_id, related_records, merge_on, related_id, count_label,
                          binary=False):
    out_df = primary_records.copy()
    out_df = pd.merge(out_df, related_records, how='left', on=merge_on)
    out_df = out_df.groupby(primary_id)[related_id].count().reset_index(name=count_label)
    out_df = pd.merge(primary_records, out_df, how='left', on=primary_id)
    if binary:
        out_df.loc[out_df[count_label] > 0, count_label] = 1
    return out_df


# Function to calculate total records for each module and update.
# df: dataframe of module corrections
# module: string value of module name
# total: boolean for whether to include a count of total module corrections
def corrections_sum(df, module, total=True):
    df_sum = df.count().to_frame(name="# of Entries").reset_index().rename(columns={'index': 'Update'})
    df_sum = df_sum.loc[df_sum['Update'].str.contains('UPDATE')]
    if total:
        df_total = {'Update': 'Total', '# of Entries': len(df)}
        df_sum = df_sum.append(df_total, ignore_index=True)
    df_sum['Module'] = module
    return df_sum


# Export a list of dataframes as an Excel workbook
# file: string for the name or path of the file
# sheet_names: list of strings for the name of each sheet
# dfs: list of dataframes for the report
# report_dict: a dict to be used in place of sheet_names and dfs
def write_report(file, sheet_names=None, dfs=None, report_dict=None):
    if report_dict is None:
        report_dict = dict(zip(sheet_names, dfs))
    writer = pd.ExcelWriter(file, engine='xlsxwriter')
    # Loop through dict of dataframes
    for sheet_name, df in report_dict.items():
        # Send df to writer
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        # Pull worksheet object
        worksheet = writer.sheets[sheet_name]
        # Loop through all columns
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
                # Len of the largest item
                series.astype(str).map(len).max(),
                # Len of column name/header
                len(str(series.name))
            )) + 1  # adding a little extra space
            # Set column width
            worksheet.set_column(idx, idx, max_len)
            worksheet.autofilter(0, 0, 0, len(df.columns) - 1)
    writer.close()


# Convert sheet openpyxl Workbook Worksheet to DataFrame
# wb: openpyxl Workbook object
# sheet: string for the sheet label
# returns a DataFrame with the first row of sheet data used for column labels
def wb_sheet_to_df(wb, sheet):
    data = wb[sheet].values
    columns = next(data)[0:]
    return pd.DataFrame(data, columns=columns)


# Set the first row of a dataframe as columns
# df: dataframe
def first_row_to_cols(df):
    out_df = df.copy()
    out_df.columns = out_df.iloc[0]
    return out_df.drop(out_df.index[0])


# load org_settings.json file as a dict
# org_settings.json must be created in /py_pears
def load_org_settings():
    org_settings_f = open(ROOT_DIR + '/org_settings.json')
    org_settings_data = json.load(org_settings_f)
    org_settings_f.close()
    return org_settings_data


# Send an email with or without a xlsx attachment
# send_from: string for the sender's email address
# send_to: string for the recipient's email address
# Cc: string of comma-separated cc addresses
# subject: string for the email subject line
# html: string for the email body
# username: string for the username to authenticate with
# password: string for the password to authenticate with
# isTls: boolean, True to put the SMTP connection in Transport Layer Security mode (default: True)
# wb: boolean, whether an Excel file should be attached to this email (default: False)
# file_path: string for the xlsx attachment's filepath (default: '')
# filename: string for the xlsx attachments filename (default: '')
def send_mail(send_from,
              send_to,
              cc,
              subject,
              html,
              username,
              password,
              is_tls=True,
              wb=False,
              file_path='',
              filename=''):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Cc'] = cc
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(html, 'html'))

    if wb:
        fp = open(file_path, 'rb')
        part = MIMEBase('application', 'vnd.ms-excel')
        part.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)

    smtp = smtplib.SMTP('smtp.office365.com', 587)
    if is_tls:
        smtp.starttls()
    try:
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to.split(',') + msg['Cc'].split(','), msg.as_string())
    except smtplib.SMTPAuthenticationError:
        print("Authentication failed. Make sure to provide a valid username and password.")
    smtp.quit()


# Function to subset module corrections for a specific staff member
# df: dataframe of module corrections
# former: boolean, True if subsetting corrections for a former staff member
# staff_email: string for the staff member's email
def staff_corrections(df, former=True, staff_email='', former_staff=pd.DataFrame()):
    if former:
        return df.loc[df['reported_by_email'].isin(former_staff['reported_by_email'])].reset_index()
    else:
        return df.loc[df['reported_by_email'] == staff_email].drop(columns=['reported_by', 'reported_by_email', 'unit'])


# Function to insert a staff member's corrections into a html email template
# dfs: dicts of module names to staff members' corrections dataframes for that module
# strs: list of strings that will be appended to the html email template string
def insert_dfs(dfs, strs):
    for heading, df in dfs.items():
        if df.empty is False:
            strs.append('<h1> ' + heading + ' </h1>' + df.to_html(border=2, justify='center'))
        else:
            strs.append('')


# Notify admin of any failed attempts to send an email
# failed_recipients: string list of recipients who failed to receive an email
# send_from: string for the sender's email address
# send_to: string for the recipient's email address
# username: string for the username to authenticate with
# password: string for the password to authenticate with
# fail_subject: string for the failure email subject line (default: 'Failure Notice')
# success_msg: String to print to console if all emails are sent successfully (default: 'Report sent successfully')
def send_failure_notice(failed_recipients,
                        send_from,
                        send_to,
                        username,
                        password,
                        fail_subject='Failure Notice',
                        success_msg='Report sent successfully'):
    if failed_recipients:
        fail_html = """The following recipients failed to receive an email:<br>
        {}
        """
        new_string = '<br>'.join(map(str, failed_recipients))
        fail_html = fail_html.format(new_string)
        send_mail(send_from=send_from,
                  send_to=send_to,
                  cc='',
                  subject=fail_subject,
                  html=fail_html,
                  username=username,
                  password=password,
                  is_tls=True)
    else:
        print(success_msg)


# Delete all the files contained in a directory
# directory: string for the directory to empty
def empty_directory(directory):
    for filename in os.listdir(directory):
        if filename == '.gitignore':
            continue

        file_path = os.path.join(directory, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
