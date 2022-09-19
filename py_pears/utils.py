import os
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
    if columns.empty:  # Refactor?
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


# load credentials.json file as a dict
# credentials.json must be created in /py_pears
def load_credentials():
    creds_f = open(ROOT_DIR + '/credentials.json')
    creds_data = json.load(creds_f)
    creds_f.close()
    return creds_data


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
