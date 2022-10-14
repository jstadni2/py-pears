import pytest

import pandas as pd
import boto3
from botocore.exceptions import ClientError
# import os
# import openpyxl
# import pandas as pd
# import numpy as np

import py_pears.utils as utils

creds = utils.load_org_settings()
profile = creds['aws_profile']
org = creds['s3_organization']
date = pd.to_datetime("today").strftime("%Y/%m/%d")


def test_aws_profile():
    try:
        boto3.Session(profile_name=profile)
    except Exception as e:
        assert False, f"'aws_profile' in org_settings.json raised an exception: \n{e}"


def test_pears_bucket():
    try:
        session = boto3.Session(profile_name=profile)
        conn = session.client('s3')
        my_bucket = 'exports.pears.oeie.org'

        conn.list_objects_v2(
            Bucket=my_bucket,
            Prefix=org + '/' + date + '/',
            MaxKeys=5)
    except conn.exceptions.NoSuchBucket as e:
        assert False, f"Bucket: {my_bucket} does not exist. Contact PEARS support via support@pears.io. \n{e}"
    except ClientError as e:
        assert False, f"Verify 's3_organization' in org_settings.json is set correctly. \n" \
                      f"Attempt to connect to the PEARS AWS S3 raised an exception: \n{e}"
    except Exception as e:
        assert False, f"Attempt to connect to the PEARS AWS S3 raised an exception: \n{e}"


def test_pears_s3_objects():
    session = boto3.Session(profile_name=profile)

    conn = session.client('s3')
    my_bucket = 'exports.pears.oeie.org'

    response = conn.list_objects_v2(
        Bucket=my_bucket,
        Prefix=org + '/' + date + '/',
        MaxKeys=20)

    assert 'Contents' in response, f"No directory or contents for {date}"

    object_filenames = []
    for f in response['Contents']:
        file = f['Key']
        object_filenames.append(file[file.rfind('/') + 1:])

    expected_exports = ['Action_Plan_Outcomes_Export.xlsx',
                        'Action_Plans_Export.xlsx',
                        'Coalition_Export.xlsx',
                        'Indirect_Activity_Export.xlsx',
                        'PSE_Site_Activity_Export.xlsx',
                        'Partnership_Export.xlsx',
                        'Program_Activities_Export.xlsx',
                        'Quarterly_Effort_Export.xlsx',
                        'Site_Export.xlsx',
                        'Social_Marketing_Campaigns_Export.xlsx',
                        'Success_Story_Export.xlsx',
                        'User_Export.xlsx']

    assert object_filenames == expected_exports
