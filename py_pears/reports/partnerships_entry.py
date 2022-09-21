import pandas as pd
import numpy as np
import py_pears.utils as utils


def main(creds,
         export_dir,
         output_dir,
         staff_list,
         unit_counties,
         prev_year_part_export,
         send_emails=False,
         report_cc='',
         report_recipients=''):

    # Download required PEARS exports from S3
    utils.download_s3_exports(profile=creds['aws_profile'],
                              org=creds['s3_organization'],
                              dst=export_dir,
                              modules=['User',
                                       'Site',
                                       'Program_Activities',
                                       'Indirect_Activity',
                                       'Partnership'])

    # Custom fields that require reformatting
    # Only needed for multi-select dropdowns
    custom_field_labels = ['fcs_program_team', 'snap_ed_grant_goals', 'fcs_grant_goals', 'fcs_special_projects',
                           'snap_ed_special_projects']

    pa_export = pd.ExcelFile(export_dir + "Program_Activities_Export.xlsx")
    pa_data = pd.read_excel(pa_export, 'Program Activity Data')
    pa_data = utils.reformat(pa_data, custom_field_labels)
    pa_data = pa_data.loc[pa_data['program_areas'] == 'SNAP-Ed']

    ia_export = pd.ExcelFile(export_dir + "Indirect_Activity_Export.xlsx")
    ia_data = pd.read_excel(ia_export, 'Indirect Activity Data')
    ia_data = utils.reformat(ia_data, custom_field_labels)
    ia_data = ia_data.loc[ia_data['program_area'] == 'SNAP-Ed']
    ia_ic = pd.read_excel(ia_export, 'Intervention Channels')
    ia_ic = utils.select_pears_data(ia_ic, record_name_field='activity')
    ia_ic_data = pd.merge(ia_data, ia_ic, how='inner', on='activity_id')

    sites = pd.read_excel(export_dir + "Site_Export.xlsx", sheet_name='Site Data')
    sites = sites.loc[sites['is_active'] == 1]

    part_export = pd.ExcelFile(export_dir + "Partnership_Export.xlsx")
    part_data = pd.read_excel(part_export, 'Partnership Data')
    part_data = utils.reformat(part_data, custom_field_labels)
    part_data = part_data.loc[part_data['program_area'] == 'SNAP-Ed']

    part_data_2021 = pd.read_excel(prev_year_part_export,
                                   sheet_name='Partnership Data')

    fy22_inep_staff = pd.read_excel(staff_list,
                                    sheet_name='SNAP-Ed Staff List')
    user_export = pd.read_excel(export_dir + "User_Export.xlsx", sheet_name='User Data')

    # Import lookup table for counties to unit
    unit_counties = pd.read_excel(unit_counties,
                                  sheet_name='PEARS Units')

    # Partnerships Data Entry Report

    part_data = utils.select_pears_data(part_data,
                                        record_name_field='partnership_name',
                                        columns=['partnership_id',
                                                 'partnership_name',
                                                 'reported_by',
                                                 'reported_by_email',
                                                 'partnership_unit',
                                                 'site_id',
                                                 'site_name',
                                                 'site_address',
                                                 'site_city',
                                                 'site_state',
                                                 'site_zip',
                                                 'created',
                                                 'modified'])

    exclude_sites = ['abc placeholder', 'U of I Extension', 'University of Illinois Extension']
    pa_data = utils.select_pears_data(pa_data,
                                      record_name_field='name',
                                      exclude_sites=exclude_sites,
                                      columns=['program_id',
                                               'program_areas',
                                               'comments',
                                               'unit',
                                               'site_id',
                                               'site_name',
                                               'site_address',
                                               'site_city',
                                               'site_state',
                                               'site_zip',
                                               'snap_ed_grant_goals',
                                               'snap_ed_special_projects',
                                               'reported_by_email']).rename(columns={'program_areas': 'program_area'})
    pa_data['id'] = 'pa' + pa_data['program_id'].astype('str')
    # If Program Activity is for a Parent Site:
    # Create Program Activity record for each child site
    pa_data = pd.merge(pa_data,
                       sites[['parent_site_name', 'site_id', 'site_name', 'address', 'city', 'state', 'zip_code']],
                       how='left', left_on='site_name', right_on='parent_site_name', suffixes=('', '_child'))

    pa_data.loc[pa_data['site_id_child'].notnull(), 'site_id'] = pa_data['site_id_child']
    pa_data.loc[pa_data['site_id_child'].notnull(), ['partnership_name', 'site_name']] = pa_data['site_name_child']
    pa_data.loc[pa_data['site_id_child'].notnull(), 'site_address'] = pa_data['address']
    pa_data.loc[pa_data['site_id_child'].notnull(), 'site_city'] = pa_data['city']
    pa_data.loc[pa_data['site_id_child'].notnull(), 'site_zip'] = pa_data['zip_code']
    pa_data.loc[pa_data['site_id_child'].notnull(), 'site_state'] = pa_data['state']
    pa_data = pa_data.drop(columns=['site_name_child', 'address', 'city', 'zip_code', 'state'])

    ia_ic_data['id'] = 'ia' + ia_ic_data['activity_id'].astype('str')
    ia_ic_data = utils.select_pears_data(ia_ic_data,
                                         record_name_field='title',
                                         exclude_sites=exclude_sites,
                                         columns=['id',
                                                  'program_area',
                                                  'unit',
                                                  'site_id',
                                                  'site_name',
                                                  'site_address',
                                                  'site_city',
                                                  'site_state',
                                                  'site_zip',
                                                  'reported_by_email'])

    part_entry = pa_data.append(ia_ic_data)
    # Unique Partnerships to enter based on Program Activities and Indirect Activities
    part_entry = part_entry.loc[~part_entry['site_id'].isin(part_data['site_id'])].drop_duplicates(
        subset='site_id', keep='first').rename(columns={'unit': 'partnership_unit',
                                                        'comments': 'program_activity_comments'})

    # Set default Partnership field values for General Information Tab
    part_entry['partnership_name'] = part_entry['site_name']
    part_entry.insert(0, 'partnership_name', part_entry.pop('partnership_name'))
    part_entry['action_plan_name'] = 'Health: Chronic Disease Prevention and Management (State - 2020-2021)'
    part_entry.insert(3, 'action_plan_name', part_entry.pop('action_plan_name'))
    # part_entry['assistance_received_recruitment'] = 1
    # part_entry['assistance_received_space'] = 1
    part_entry['assistance_received'] = 'Recruitment (includes program outreach), ' \
                                        'Space (e.g., facility or room where programs take place)'
    # part_entry['assistance_provided_human_resources'] = 1
    # part_entry['assistance_provided_program_implementation'] = 1
    part_entry['assistance_provided'] = 'Human resources (*staff or staff time), ' \
                                        'Program implementation (e.g. food and beverage standards)'
    part_entry['assistance_received_funding'] = 'No'
    part_entry.loc[part_entry['id'].str.contains('pa'), 'is_direct_education_intervention'] = 1
    part_entry.loc[part_entry['id'].str.contains('ia'), 'is_direct_education_intervention'] = 0

    # Determine applicable Partnership collaborators
    part_entry['collaborator_unit'] = part_entry['partnership_unit']
    part_entry = pd.merge(part_entry, unit_counties, how='left', left_on='partnership_unit', right_on='County')
    part_entry.loc[part_entry['partnership_unit'].isin(unit_counties['County']),
                   'collaborator_unit'] = part_entry['Unit']
    part_entry = part_entry.drop(columns={'Unit', 'County'})
    staff_nulls = ('N/A', 'NEW', 'OPEN', np.nan)
    collaborators = fy22_inep_staff.loc[(~fy22_inep_staff['NAME'].isin(staff_nulls))
                                        & (fy22_inep_staff['JOB CLASS'].isin(['EPC', 'UE'])), 'E-MAIL']
    collaborators = pd.merge(collaborators,
                             user_export[['full_name', 'email', 'unit', 'viewable_units']],
                             how='inner',
                             left_on='E-MAIL',
                             right_on='email').drop(
        columns=['E-MAIL', 'email']).rename(
        columns={'full_name': 'collaborators'}).drop_duplicates()
    collaborators['viewable_units'] = collaborators['viewable_units'].str.split(", ")
    collaborators.loc[collaborators['viewable_units'].isnull(), 'viewable_units'] = ""
    collaborators.loc[collaborators.viewable_units.map(len) > 1, 'unit'] = collaborators['viewable_units']
    collaborators = collaborators.explode('unit').drop(columns=['viewable_units'])
    part_collaborators = pd.merge(part_entry[['partnership_name', 'collaborator_unit']], collaborators, how='left',
                                  left_on='collaborator_unit', right_on='unit')
    part_collaborators = part_collaborators.groupby('partnership_name').agg(lambda x: x.dropna().unique().tolist())
    part_collaborators = part_collaborators.drop(columns={'collaborator_unit', 'unit'})
    part_entry = pd.merge(part_entry, part_collaborators, how='left', on='partnership_name').drop(
        columns=['collaborator_unit'])
    part_entry['collaborators'] = [', '.join(map(str, collab_list)) for collab_list in part_entry['collaborators']]

    # Set default field values for Evaluation Tab
    part_entry['relationship_depth'] = 'Cooperator'
    part_entry['assessment_tool'] = 'None'
    part_entry['accomplishments'] = 'N/A'
    part_entry['lessons_learned'] = 'N/A'

    # Subset Partnerships to copy forward from previous report year
    c_parts_site_id = pd.merge(part_entry,
                               part_data_2021[['partnership_id',
                                               'partnership_name',
                                               'site_id',
                                               'site_name',
                                               'site_zip']], how='left', on='site_id',
                               suffixes=('', '_copy')).rename(columns={'partnership_id': 'partnership_id_copy'})
    c_parts_site_id = c_parts_site_id.loc[c_parts_site_id['partnership_id_copy'].notnull()]
    c_parts_site_id = c_parts_site_id[['id',
                                       'partnership_id_copy',
                                       'partnership_name_copy',
                                       'program_area',
                                       'action_plan_name',
                                       'site_id',
                                       'site_name_copy',
                                       'site_address',
                                       'site_city',
                                       'site_state',
                                       'site_zip',
                                       'partnership_unit',
                                       'assistance_received',
                                       'assistance_provided',
                                       'assistance_received_funding',
                                       'is_direct_education_intervention',
                                       'collaborators',
                                       'snap_ed_grant_goals',
                                       'snap_ed_special_projects',
                                       'relationship_depth',
                                       'parent_site_name',
                                       'program_activity_comments',
                                       'reported_by_email']]

    # Subset new Partnerships to create
    new_parts = part_entry.loc[~part_entry['site_id'].isin(c_parts_site_id['site_id']),
                               ['partnership_name',
                                'id',
                                'program_area',
                                'action_plan_name',
                                'program_activity_comments',
                                'partnership_unit',
                                'site_id',
                                'site_name',
                                'site_address',
                                'site_city',
                                'site_state',
                                'site_zip',
                                'parent_site_name',
                                'site_id_child',
                                'assistance_received',
                                'assistance_provided',
                                'assistance_received_funding',
                                'is_direct_education_intervention',
                                'collaborators',
                                'snap_ed_grant_goals',
                                'snap_ed_special_projects',
                                'relationship_depth',
                                'assessment_tool',
                                'accomplishments',
                                'lessons_learned',
                                'reported_by_email']]
    new_parts = new_parts.drop(columns='site_id_child')
    # Create utils function for insert - pop method
    new_parts.insert((len(new_parts.columns) - 1), 'program_activity_comments',
                     new_parts.pop('program_activity_comments'))

    # Create utils function for prev_month()
    prev_month = (pd.to_datetime("today") - pd.DateOffset(months=1)).to_period('M')

    # SNAP-Ed Workbook

    snap_ed_c_parts_site_id = c_parts_site_id.loc[c_parts_site_id['partnership_unit'] != 'CPHP (District)']
    snap_ed_new_parts = new_parts.loc[new_parts['partnership_unit'] != 'CPHP (District)']

    fcs_dfs = {'New Partnerships': snap_ed_new_parts, 'Copy Forward - Site ID Matches': snap_ed_c_parts_site_id}

    fcs_filename = 'SNAP-Ed Partnerships Data Entry ' + prev_month.strftime('%Y-%m') + '.xlsx'

    fcs_file_path = output_dir + fcs_filename

    utils.write_report(fcs_file_path, report_dict=fcs_dfs)

    # CPHP Workbook

    cphp_c_parts_site_id = c_parts_site_id.loc[c_parts_site_id['partnership_unit'] == 'CPHP (District)']
    cphp_new_parts = new_parts.loc[new_parts['partnership_unit'] == 'CPHP (District)']

    cphp_dfs = {'New Partnerships': cphp_new_parts, 'Copy Forward - Site ID Matches': cphp_c_parts_site_id}

    cphp_filename = 'CPHP Partnerships Data Entry ' + prev_month.strftime('%Y-%m') + '.xlsx'

    cphp_file_path = output_dir + cphp_filename

    utils.write_report(cphp_file_path, report_dict=cphp_dfs)

    # Email Data Entry Report

    if send_emails:

        fcs_report_subject = 'SNAP-Ed Partnerships Data Entry ' + prev_month.strftime('%Y-%m')

        report_html = """<html>
          <head></head>
        <body>
                    <p>
                   Hello DATA ENTRY SUPPORT,<br><br>
                    The attached data is for Direct/Indirect Education partners that require Partnership Module entries.
                    Could you please enter them into PEARS? Should you need it, the Partnerships Cheat Sheet is located
                    <a href="https://uofi.app.box.com/folder/49632670918?s=wwymjgjd48tyl0ow20vluj196ztbizlw">here</a>.
                    <ul>
                      <li>New Partnerships for Direct Education contain 'pa' in the id field and whereas the id for
                      Indirect Education Partnerships contain 'ia'.</li>
                      <li>If the Partnership Unit is set to
                      'Illinois - University of Illinois Extension (Implementing Agency)',
                      please select a more appropriate unit.</li>
                      <li>When copying forward Partnerships from a previous year, make sure the new entry matches the
                      data in this spreadsheet.</li>
                      <li>Copied Partnerships should only display '(Copied)' in the title once.</li>
                      <li>District-level Direct Education requires an individual Partnership
                      for each Site in attendance.</li>
                      <li>If the Parent Site column is not empty, please verify that all sites listed in the
                      Program Activity Comments have corresponding Site and Partnership entries.</li>
                      <li>If the SNAP-Ed Grant Goals or SNAP-Ed Special Projects fields are empty,
                      contact staff who created the original record (in the ID field) for the appropriate values.</li>
                    </ul>
                  If you have any questions, please reply to this email and I will respond at my
                  earliest opportunity.<br>
                    <br>Thanks and have a great day!<br>
                    <br> <b> FCS Evaluation Team </b> <br>
                    <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
                    </p>
          </body>
        </html>
        """

        utils.send_mail(send_from=creds['admin_send_from'],
                        send_to=report_recipients,
                        cc=report_cc,
                        subject=fcs_report_subject,
                        html=report_html,
                        username=creds['admin_username'],
                        password=creds['admin_password'],
                        wb=True,
                        is_tls=True,
                        file_path=fcs_file_path,
                        filename=fcs_filename)

        cphp_report_subject = 'CPHP Partnerships Data Entry ' + prev_month.strftime('%Y-%m')

        if any(x.empty is False for x in cphp_dfs.values()):
            utils.send_mail(send_from=creds['admin_send_from'],
                            send_to=report_recipients,
                            cc=report_cc,
                            subject=cphp_report_subject,
                            html=report_html,
                            username=creds['admin_username'],
                            password=creds['admin_password'],
                            wb=True,
                            is_tls=True,
                            file_path=cphp_file_path,
                            filename=cphp_filename)
