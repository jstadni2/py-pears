import os
import boto3
import pandas as pd

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
def download_s3_exports(profile, org, date=pd.to_datetime("today").strftime("%Y/%m/%d"),
                        dst=ROOT_DIR + "/pears_exports"):
    # Use PEARS AWS S3 credentials
    session = boto3.Session(profile_name=profile)

    # Access S3 objects uploaded the day reformatting script is run
    conn = session.client('s3')
    my_bucket = 'exports.pears.oeie.org'

    response = conn.list_objects_v2(
        Bucket=my_bucket,
        Prefix=org + '/' + date + '/',
        MaxKeys=100)

    # Download the Excel files to the destination directory
    for f in response['Contents']:
        file = f['Key']
        filename = file[file.rfind('/') + 1:]
        conn.download_file(my_bucket, file, dst + '/' + filename)


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
# ALTERNATIVELY, INPUT DICT
def write_report(file, sheet_names, dfs):
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

