import pandas as pd
import numpy as np
import smtplib
import py_pears.utils as utils


def main(creds,
         export_dir,
         output_dir,
         coalition_surveys_dir,
         staff_list,
         unit_counties,
         update_notifications,
         send_emails=False,
         notification_cc='',
         report_cc='',
         report_recipients=''):

    # Download required PEARS exports from S3
    utils.download_s3_exports(profile=creds['aws_profile'],
                              org=creds['s3_organization'],
                              dst=export_dir,
                              modules=['Coalition'])

    # Custom fields that require reformatting
    # Only needed for multi-select dropdowns
    custom_field_labels = ['fcs_program_team', 'snap_ed_grant_goals', 'fcs_grant_goals', 'fcs_special_projects',
                           'snap_ed_special_projects']

    coa_export = pd.ExcelFile(export_dir + "Coalition_Export.xlsx")
    coa_data = pd.read_excel(coa_export, 'Coalition Data')
    # create a utils function for removing _custom_data from column labels
    coa_data = utils.reformat(coa_data, custom_field_labels)  # this is only necessary for the on_hiatus field
    coa_data = coa_data.loc[coa_data['program_area'].isin(['SNAP-Ed', 'Family Consumer Science'])]
    coa_data['coalition_id'] = coa_data['coalition_id'].astype(str)

    # UPDATE: Implement a try-catch and/or util function
    prev_month = (pd.to_datetime("today") - pd.DateOffset(months=1)).strftime('%m')
    fq_lookup = pd.DataFrame({'fq': ['Q1', 'Q2', 'Q3', 'Q4'], 'month': ['12', '03', '06', '09'],
                              'survey_fq': ['Quarter 1 (October-December)', 'Quarter 2 (January-March)',
                                            'Quarter 3 (April-June)', 'Quarter 4 (July-September)']})
    fq = fq_lookup.loc[fq_lookup['month'] == prev_month, 'fq'].item()
    survey_fq = fq_lookup.loc[fq_lookup['month'] == prev_month, 'survey_fq'].item()
    # Using manual filename convention
    coa_surveys = pd.read_excel(coalition_surveys_dir + "Coalition_Survey_" + fq + "_Export.xlsx",
                                sheet_name='Response Data')
    # Default PEARS survey export convention
    # coa_surveys = pd.read_excel(coalition_surveys_dir + "Responses By Survey - Coalition Survey - " + fq + ".xlsx",
                                # sheet_name='Response Data')

    # filter Responses By Survey by Completed == ---- to export all responses
    coa_surveys = utils.select_pears_data(coa_surveys,
                                          record_name_field='coalition_name',
                                          columns=['Program Activity ID',
                                                   'Program Name',
                                                   'Unique PEARS ID of Response',
                                                   'staff_email',
                                                   'What is the Coalition ID from the PEARS Coalition module that '
                                                   'corresponds to this survey?',
                                                   'coalition_name',
                                                   'For which Quarter are you completing this survey?&nbsp;'])

    coa_surveys = coa_surveys.loc[coa_surveys['For which Quarter are you completing this survey?&nbsp;'] == survey_fq]
    coa_surveys = coa_surveys.rename(columns={'Program Activity ID': 'program_id',
                                              'Program Name': 'program_name',
                                              'staff_email': 'reported_by_email',
                                              'What is the Coalition ID from the PEARS Coalition module that '
                                              'corresponds to this survey?': 'coalition_id',
                                              'Unique PEARS ID of Response': 'response_id',
                                              'For which Quarter are you completing this survey?&nbsp;':
                                                  'survey_quarter'})
    # Remove all characters besides digits from coalition_id
    coa_surveys['coalition_id'] = coa_surveys['coalition_id'].astype(str)
    coa_surveys.loc[~coa_surveys['coalition_id'].str.isnumeric(),
                    'coalition_id'] = coa_surveys['coalition_id'].str.extract('(\d+)', expand=False)
    # Auto export?
    # Responses by Survey filters: Name == Coalition Survey & Survey Status == Active
    # Export filters: Reporting Period == Extension 2021 & Type of Export == Individual Responses

    # Import Update Notifications, used for the Corrections Report
    update_notes = pd.read_excel(update_notifications,
                                 sheet_name='Quarterly Data Cleaning').drop(columns='Tab')

    # Import and consolidate staff lists
    # Data cleaning is only conducted on records related to SNAP-Ed and Family Consumer Science programming

    fy22_inep_staff = pd.ExcelFile(staff_list)
    snap_ed_staff = pd.read_excel(fy22_inep_staff, sheet_name='SNAP-Ed Staff List', header=0)
    heat_staff = pd.read_excel(fy22_inep_staff, sheet_name='HEAT Project Staff', header=0)
    state_staff = pd.read_excel(fy22_inep_staff, sheet_name='FCS State Office', header=0)
    staff_cols = ['NAME', 'E-MAIL']
    staff_dfs = [snap_ed_staff[staff_cols], heat_staff[staff_cols], state_staff[staff_cols]]
    inep_staff = pd.concat(staff_dfs, ignore_index=True).rename(columns={'E-MAIL': 'email'})
    inep_staff = inep_staff.loc[~inep_staff.isnull().any(1)]
    inep_staff['NAME'] = inep_staff['NAME'].str.split(pat=', ')
    inep_staff['first_name'] = inep_staff['NAME'].str[1]
    inep_staff['last_name'] = inep_staff['NAME'].str[0]
    inep_staff['full_name'] = inep_staff['first_name'].map(str) + ' ' + inep_staff['last_name'].map(str)
    cphp_staff = pd.read_excel(fy22_inep_staff, sheet_name='CPHP Staff List', header=0).rename(
        columns={'Last Name': 'last_name',
                 'First Name': 'first_name',
                 'Email Address': 'email'})
    cphp_staff['full_name'] = cphp_staff['first_name'].map(str) + ' ' + cphp_staff['last_name'].map(str)
    staff = inep_staff.drop(columns='NAME').append(
        cphp_staff.loc[~cphp_staff['email'].isnull(), ['email', 'first_name', 'last_name', 'full_name']],
        ignore_index=True).drop_duplicates()

    # Create lookup table for unit to regional educators
    re_lookup = pd.read_excel(fy22_inep_staff, sheet_name="RE's and CD's")[
        ['UNIT #', 'REGIONAL EDUCATOR', 'RE E-MAIL']]
    re_lookup['REGIONAL EDUCATOR'] = re_lookup['REGIONAL EDUCATOR'].str.replace(', Interim', '')
    re_lookup = re_lookup.drop_duplicates()
    re_lookup = utils.reorder_name(re_lookup, 'REGIONAL EDUCATOR', 'REGIONAL EDUCATOR', drop_substr_fields=True)
    re_lookup['UNIT #'] = re_lookup['UNIT #'].astype(str)

    # Import lookup table for counties to unit
    unit_counties = pd.read_excel(unit_counties)
    unit_counties['Unit #'] = unit_counties['Unit #'].astype(str)

    # Coalition Surveys Data Cleaning

    # Coalitions

    coa_data = utils.counties_to_units(coa_data, unit_field='coalition_unit', unit_counties=unit_counties)
    coa_data = utils.select_pears_data(coa_data,
                                       record_name_field='coalition_name',
                                       columns=['coalition_id',
                                                'coalition_name',
                                                'reported_by_email',
                                                'coalition_unit',
                                                'program_area',
                                                'relationship_depth',
                                                'created',
                                                'modified',
                                                'on_hiatus']).rename(columns={'coalition_unit': 'unit'})

    coa_data['UPDATES'] = np.nan
    coa_data.loc[(coa_data['relationship_depth'].isin(['Coalition', 'Collaboration', 'Coordination']))
                 & (~coa_data['coalition_id'].isin(coa_surveys['coalition_id']))
                 & (coa_data['on_hiatus'] != 'Yes'),
                 'UPDATES'] = utils.get_update_note(update_notes, module='Coalitions',
                                                    update='UPDATES',
                                                    notification='Notification')

    coa_corrections = coa_data.loc[(coa_data['UPDATES'].notnull())].drop(columns=['program_area',
                                                                                  'created',
                                                                                  'modified']).fillna('')
    # Send to corrections report and email

    # Coalition Surveys

    # How do staff update their survey responses?
    # Make all staff collaborators on statewide PA?

    # Data Validation:
    # 'What is the coalition_id from the PEARS Coalition module that corresponds to this survey?' == numeric only

    coa_surveys['EVALUATION TAB UPDATES'] = np.nan
    coa_surveys.loc[~coa_surveys['coalition_id'].isin(coa_data['coalition_id']),
                    'EVALUATION TAB UPDATES'] = utils.get_update_note(update_notes,
                                                                      module='Program Activities',
                                                                      update='EVALUATION TAB UPDATES',
                                                                      notification='Notification')

    coa_survey_corrections = coa_surveys.loc[coa_surveys['EVALUATION TAB UPDATES'].notnull()].set_index(
        'program_id').fillna('')
    # Send to corrections report
    coa_survey_corrections_email = coa_survey_corrections.drop(columns='response_id')
    # Send to corrections email

    # Corrections Report

    # Summarize and concatenate module corrections
    corrections_dict = {
        'Coalitions': coa_corrections,
        'Program Activities': coa_survey_corrections,
    }

    module_sums = [utils.corrections_sum(corrections,
                                         module,
                                         total=False) for module, corrections in corrections_dict.items()]

    corrections_sums = pd.concat(module_sums, ignore_index=True)

    corrections_sums.insert(0, 'Module', corrections_sums.pop('Module'))

    corrections_sums = pd.merge(corrections_sums, update_notes, how='left', on=['Module', 'Update'])

    report_filename = 'Quarterly Coalition Survey Entry ' + fq + '.xlsx'
    report_file_path = output_dir + report_filename

    utils.write_report(report_file_path, report_dict={
        'Corrections Summary': corrections_sums,
        'Coalitions': coa_corrections,
        'Coalition Surveys': coa_survey_corrections
    })

    # Email Survey Notifications

    if send_emails:

        deadline_date = pd.to_datetime("today").replace(day=19).strftime('%A %b %d, %Y')

        notification_html = """<html>
          <head></head>
        <body>
                    <p>
                    Hello {0},<br><br>
                    You are receiving this email because you need to submit or update quarterly Coalition Surveys.
                    Please update the entries listed in the table(s) below by <b>5:00pm {1}</b>.
                    <ul>
                      <li>Coalition Surveys are required for any Coalition in the Coordination, Coalition,
                      or Collaboration stage of development.</li>
                      <li>Use the following link to submit <b>new</b> Coalition Surveys for each Coalition listed below.
                       <a href="https://bit.ly/3qXvAAO">https://bit.ly/3qXvAAO</a></li>
                      <li>For each entry listed, please make the edit(s) displayed in the columns labeled <b>UPDATE</b>
                      in the column heading.</li>
                      <li>You can locate entries in PEARS by entering their IDs into the search filter.</li>
                      <li>As a friendly reminder – following the Cheat Sheets
                      <a href="https://uofi.app.box.com/folder/49632670918?s=wwymjgjd48tyl0ow20vluj196ztbizlw">
                      [Located Here]</a>
                      will help to prevent future PEARS corrections.</li>
                  </ul>
                  {2}
                    <br>{3}<br>
                    {4}<br>
                    </p>
          </body>
        </html>
        """

        # Create dataframe of staff to notify
        notify_staff = coa_corrections[['reported_by_email',
                                        'unit']].append(coa_survey_corrections_email[['reported_by_email']],
                                                        ignore_index=True).drop_duplicates(subset='reported_by_email',
                                                                                           keep='first')
        # notify_staff = coa_survey_corrections_email[['reported_by_email']].drop_duplicates()

        # Subset current staff using the staff list
        current_staff = notify_staff.loc[notify_staff['reported_by_email'].isin(staff['email']),
                                         ['reported_by_email', 'unit']]
        current_staff = current_staff.values.tolist()

        # If email fails to send, the recipient is added to this list
        failed_recipients = []

        # Email Update Notifications to current staff

        for x in current_staff:
            recipient = x[0]
            unit = x[1]

            staff_name = staff.loc[staff['email'] == recipient, 'full_name'].item()

            notification_subject = 'Coalition Survey Entry ' + fq + ', ' + staff_name

            response_tag = """If you have any questions or need help please reply to this email and a member of the
            FCS Evaluation Team will reach out soon.
                    <br>Thanks and have a great day!<br>
                    <br> <b> FCS Evaluation Team </b> <br>
                    <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
            """

            new_notification_cc = notification_cc

            if (unit in re_lookup["UNIT #"].tolist()) \
                and (recipient not in state_staff['E-MAIL'].tolist()) \
                    and ('@uic.edu' not in recipient):
                response_tag = 'If you have any questions or need help please contact your Regional Specialist,' \
                               ' <b>{0}</b> (<a href = "mailto: {1} ">{1}</a>).'
                re_name = re_lookup.loc[re_lookup['UNIT #'] == unit, 'REGIONAL EDUCATOR'].item()
                re_email = re_lookup.loc[re_lookup['UNIT #'] == unit, 'NETID/E-MAIL'].item()
                response_tag = response_tag.format(*[re_name, re_email])
                new_notification_cc = notification_cc + ', ' + re_email

            notification_dfs = {'Coalitions': utils.staff_corrections(coa_corrections,
                                                                      former=False,
                                                                      staff_email=recipient),
                                'Coalition Surveys': utils.staff_corrections(coa_survey_corrections_email,
                                                                             former=False,
                                                                             staff_email=recipient)}

            y = [staff.loc[staff['email'] == recipient, 'first_name'].item(), deadline_date, response_tag]

            utils.insert_dfs(notification_dfs, y)
            new_notification_html = notification_html.format(*y)

            # Try to send the email, otherwise add the recipient to failed_recipients
            try:
                utils.send_mail(send_from=creds['admin_send_from'],
                                send_to=recipient,
                                cc=new_notification_cc,
                                subject=notification_subject,
                                html=new_notification_html,
                                username=creds['admin_username'],
                                password=creds['admin_password'],
                                wb=False,
                                is_tls=True)
            except smtplib.SMTPException:
                failed_recipients.append([staff_name, x])

        # Email Update Notifications for former staff

        # Subset former staff using the staff list
        former_staff = notify_staff.loc[~notify_staff['reported_by_email'].isin(staff['email'])]

        coa_df = utils.staff_corrections(coa_corrections, former=True, former_staff=former_staff)
        pa_df = utils.staff_corrections(coa_survey_corrections_email, former=True, former_staff=former_staff)

        former_staff_subject = 'Former Staff Coalition Survey Entry ' + fq

        former_staff_filename = former_staff_subject + '.xlsx'
        former_staff_file_path = output_dir + former_staff_filename

        # Export former staff corrections as an Excel file
        utils.write_report(former_staff_file_path, report_dict={'Coalitions': coa_df, 'Coalition Surveys': pa_df})

        # Send former staff updates email

        former_staff_report_recipients = 'recipient@domain.com'

        former_staff_html = """<html>
          <head></head>
        <body>
                    <p>
                    Hello DATA ENTRY SUPPORT et al,<br><br>
                    The attached Excel workbook compiles Coalition entries created by former staff that require
                    Coalition Surveys and surveys that require updates.
                    Please complete the updates for each record by <b>5:00pm {0}</b>.
                    <ul>
                      <li>Use the following link to submit <b>new</b> Coalition Surveys for each Coalition listed below.
                       <a href="https://bit.ly/3qXvAAO">https://bit.ly/3qXvAAO</a></li>
                      <li>For each entry listed, please make the edit(s) written in the columns labeled <b>UPDATE</b>
                      in the column heading.</li>
                      <li>You can locate entries in PEARS by entering their IDs into the search filter.</li>
                    </ul>
                  If you have any questions or need help please reply to this email and a member of the
                  FCS Evaluation Team will reach out soon.
                    <br>Thanks and have a great day!<br>
                    <br> <b> FCS Evaluation Team </b> <br>
                    <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
                    </p>
          </body>
        </html>
        """
        y = [deadline_date]

        new_former_staff_html = former_staff_html.format(*y)

        try:
            if not coa_df.empty:
                utils.send_mail(send_from=creds['admin_send_from'],
                                send_to=former_staff_report_recipients,
                                cc=notification_cc,
                                subject=former_staff_subject,
                                html=new_former_staff_html,
                                username=creds['admin_username'],
                                password=creds['admin_password'],
                                wb=True,
                                file_path=former_staff_file_path,
                                filename=former_staff_filename,
                                is_tls=True)
        except smtplib.SMTPException:
            failed_recipients.append(['DATA ENTRY SUPPORT NAME', former_staff_report_recipients])

        report_recipients = 'list@domain.com, of_recipients@domain.com'

        report_subject = 'Quarterly Coalition Survey Entry Q2 ' + fq

        report_html = """<html>
          <head></head>
        <body>
                    <p>
                    Hello everyone,<br><br>
                    The attached reported compiles the most recent round of quarterly Coalition Survey entry.
                    If you have any questions, please reply to this email and a member of the FCS Evaluation Team
                    will reach out soon.<br>
                    <br>Thanks and have a great day!<br>
                    <br> <b> FCS Evaluation Team </b> <br>
                    <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
                    </p>
          </body>
        </html>
        """

        try:
            utils.send_mail(send_from=creds['admin_send_from'],
                            send_to=report_recipients,
                            cc=report_cc,
                            subject=report_subject,
                            html=report_html,
                            username=creds['admin_username'],
                            password=creds['admin_password'],
                            wb=True,
                            file_path=report_file_path,
                            filename=report_filename,
                            is_tls=True)
        except smtplib.SMTPException:
            print("Failed to send report to Regional Specialists.")

            # Notify admin of any failed attempts to send an email
            utils.send_failure_notice(failed_recipients=failed_recipients,
                                      send_from=creds['admin_send_from'],
                                      send_to=creds['admin_send_from'],
                                      username=creds['admin_username'],
                                      password=creds['admin_password'],
                                      fail_subject=report_subject + ' Failure Notice',
                                      success_msg='Data cleaning notifications sent successfully.')
