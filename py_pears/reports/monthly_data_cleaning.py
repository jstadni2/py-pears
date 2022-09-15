import pandas as pd
import numpy as np
import smtplib
import py_pears.utils as utils


# functions for reordering comma-separated name
# df: dataframe of staff list
# name_field: column label of name field
# reordered_name_field: column label of reorded name field
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


# Function to drop duplicate recoorts from merging module data with child records
# df: dataframe of module corrections
# c_updates: list of labels of child record update columns
# parent_id: column label for the module ID
# child_id: column label for the child record ID
def drop_child_dupes(df, c_updates, parent_id, child_id):
    df_dupes = df.loc[df[c_updates].isnull().all(1) & df.duplicated(subset=[parent_id], keep=False)]
    df_dupes_filtered = df_dupes.drop_duplicates(subset=[parent_id], keep='first')
    df = df.loc[~df[child_id].isin(df_dupes[child_id])]
    df = df.append(df_dupes_filtered)
    return df


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


# Concatenate update columns into a single column of newline-separated update notifications
# data: dataframe of PEARS module data
# concat_col: string for the label of the concatenated column
# update_cols: list of strings of update columns to concatenate
def concat_updates(data, concat_col, update_cols):
    out_data = data.copy()
    # Use a lambda/list comprehension?
    out_data[concat_col] = out_data[update_cols].apply(lambda x: '\n'.join(x.fillna('').values.tolist()), axis=1)
    out_data.loc[out_data[concat_col].str.isspace(), concat_col] = np.nan
    out_data[concat_col] = out_data[concat_col].str.strip()
    if len(update_cols) > 2:
        out_data[concat_col] = out_data[concat_col].str.replace(r'\n+', '', regex=True)
    return out_data


# Format corrections for the HTML tables in update notification emails
# corrections: dataframe of PEARS module corrections
# cols: list of strings for the columns of corrections to include in the notification table
# index: string for the column label of the PEARS module index
# int_cols: list of strings for the columns to downcast to integers
# rename_cols: dict of columns to rename, key of the original column and value for the renamed column
# update_cols: update columns that will have newlines replaced with spaces
# date_cols: datetime columns to convert to strings in '%m-%d-%Y' format
# datetime_cols: datetime columns to convert to strings in '%m-%d-%Y %r' format
def corrections_email_format(corrections, cols, index,
                             int_cols=None,
                             rename_cols=None,
                             update_cols=None,
                             date_cols=None,
                             datetime_cols=None):
    email_corrections = corrections.copy()
    email_corrections = email_corrections[cols].set_index(index)

    if int_cols:
        for int_col in int_cols:  # Use lambda?
            email_corrections[int_col] = pd.to_numeric(email_corrections[int_col], downcast='integer')

    if rename_cols:
        email_corrections = email_corrections.rename(columns=rename_cols)

    if update_cols:
        for update_col in update_cols:
            email_corrections[update_col] = email_corrections[update_col].str.replace('\n', ' ')

    if date_cols:
        for date_col in date_cols:
            email_corrections[date_col] = pd.to_datetime(email_corrections[date_col]).dt.strftime('%m-%d-%Y')

    if datetime_cols:
        for datetime_col in datetime_cols:
            email_corrections[datetime_col] = pd.to_datetime(email_corrections[datetime_col]).dt.strftime('%m-%d-%Y %r')

    email_corrections = email_corrections.fillna('')
    return email_corrections


# Function to calculate total records for each module and update.
# df: dataframe of module corrections
# module: string value of module name
def corrections_sum(df, module):
    df_sum = df.count().to_frame(name="# of Entries").reset_index().rename(columns={'index': 'Update'})
    df_sum = df_sum.loc[df_sum['Update'].str.contains('UPDATE')]
    df_total = {'Update': 'Total', '# of Entries': len(df)}
    df_sum = df_sum.append(df_total, ignore_index=True)
    df_sum['Module'] = module
    return df_sum


# Export the corrections report as a xlsx
# report_dict: dict of sheet names to dataframes of corrections data
# file_path: string for the output directory and filename
def write_corrections_report(report_dict, file_path):
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    for sheet_name, df in report_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False, freeze_panes=(1, 0))
        worksheet = writer.sheets[sheet_name]
        workbook = writer.book
        worksheet.autofilter(0, 0, 0, len(df.columns) - 1)
        # Conditional format for Corrections Summary
        blue_bold = workbook.add_format({'bold': True, 'bg_color': '#DEEAF0', 'font_color': '#000000'})
        worksheet.conditional_format(0, 0, len(df), 2,
                                     {'type': 'formula', 'criteria': '=$B1="Total"', 'format': blue_bold})
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),
                len(str(series.name))
            )) + 1
            worksheet.set_column(idx, idx, max_len)
    writer.close()

# Function to subset module corrections for a specific staff member
# df: dataframe of module corrections
# former: boolean, True if subsetting corrections for a former staff member
# staff_email: string for the staff member's email
def staff_corrections(df, former=True, staff_email='', former_staff=pd.DataFrame()):
    if former:
        return df.loc[df['reported_by_email'].isin(former_staff['reported_by_email'])].reset_index()
    else:
        return df.loc[df['reported_by_email'] == staff_email].drop(columns=['reported_by', 'reported_by_email', 'unit'])


# Function to insert a staff member's corrections into an html email template
# dfs: dicts of module names to staff members' corrections dataframes for that module
# strs: list of strings that will be appened to the html email template string
def insert_dfs(dfs, strs):
    for heading, df in dfs.items():
        if df.empty is False:
            strs.append('<h1> ' + heading + ' </h1>' + df.to_html(border=2, justify='center'))
        else:
            strs.append('')


# Run the Monthly Data Cleaning
# creds: dict of credentials loaded from credentials.json
# export_dir: directory where PEARS exports are downloaded to
# output_dir: directory where report outputs are saved
# staff_list: path to the staff list Excel workbook
# names_list: path to a text file containing names used to flag the Name field of Coalition Member records
# unit_counties: path to a workbook that maps counties to Extension units
# update_notifications: path to a workbook that compiles the update notifications
# send_emails: boolean for sending emails associated with this report (default: False)
# notification_cc: list-like string of email addresses to cc on unauthorized site creation notifications
# former_staff_recipients: list-like string of email addresses for recipients of the former staff corrections email
# report_cc: list-like string of email addresses to cc on the report email
# report_recipients: list-like string of email addresses for recipients of the report email
def main(creds,
         export_dir,
         output_dir,
         staff_list,
         names_list,
         unit_counties,
         update_notifications,
         send_emails=False,
         notification_cc='',
         former_staff_recipients='',
         report_cc='',
         report_recipients=''):

    # Import and consolidate staff lists
    # Data cleaning is only conducted on records related to SNAP-Ed and Family Consumer Science programming

    fy22_inep_staff = pd.ExcelFile(staff_list)
    # Alternatively, use the absolute path to the staff list
    # fy22_inep_staff = pd.ExcelFile(r"C:\Users\netid\Box\INEP Staff Lists\FY22 INEP Staff List.xlsx")
    # Adjust header argument in following lines for actual staff list
    snap_ed_staff = pd.read_excel(fy22_inep_staff, sheet_name='SNAP-Ed Staff List', header=0)
    heat_staff = pd.read_excel(fy22_inep_staff, sheet_name='HEAT Project Staff', header=0)
    state_staff = pd.read_excel(fy22_inep_staff, sheet_name='FCS State Office', header=0)
    staff_cols = ['NAME', 'E-MAIL']
    staff_dfs = [snap_ed_staff[staff_cols], heat_staff[staff_cols], state_staff[staff_cols]]
    inep_staff = pd.concat(staff_dfs, ignore_index=True).rename(columns={'E-MAIL': 'email'})
    inep_staff = inep_staff.loc[~inep_staff.isnull().any(1)]
    inep_staff = reorder_name(inep_staff, 'NAME', 'full_name')
    cphp_staff = pd.read_excel(fy22_inep_staff, sheet_name='CPHP Staff List', header=0).rename(
        columns={'Last Name': 'last_name',
                 'First Name': 'first_name',
                 'Email Address': 'email'})
    cphp_staff['full_name'] = cphp_staff['first_name'].map(str) + ' ' + cphp_staff['last_name'].map(str)
    staff = inep_staff.drop(columns='NAME').append(cphp_staff.loc[~cphp_staff['email'].isnull(),
                                                                  ['email', 'first_name', 'last_name', 'full_name']],
                                                   ignore_index=True).drop_duplicates()

    # Create lookup table for unit to regional educators
    re_lookup = pd.read_excel(fy22_inep_staff, sheet_name="RE's and CD's")[['UNIT #', 'REGIONAL EDUCATOR', 'RE E-MAIL']]
    re_lookup['REGIONAL EDUCATOR'] = re_lookup['REGIONAL EDUCATOR'].str.replace(', Interim', '')
    re_lookup = re_lookup.drop_duplicates()
    re_lookup = reorder_name(re_lookup, 'REGIONAL EDUCATOR', 'REGIONAL EDUCATOR', drop_substr_fields=True)
    re_lookup['UNIT #'] = re_lookup['UNIT #'].astype(str)

    # Import list of former staff
    # Used to send former staff's updates to evaluation team
    former_snap_ed_staff = pd.read_excel(fy22_inep_staff, sheet_name='Former Staff')
    former_snap_ed_staff['email'] = former_snap_ed_staff['NETID'].map(str) + '@illinois.edu'

    # Import lookup table for counties to unit
    unit_counties = pd.read_excel(unit_counties)
    unit_counties['Unit #'] = unit_counties['Unit #'].astype(str)

    # Custom fields that require reformatting
    # Only needed for multi-select dropdowns
    custom_field_labels = ['fcs_program_team', 'snap_ed_grant_goals', 'fcs_grant_goals', 'fcs_special_projects',
                           'snap_ed_special_projects']

    # Import Coalitions data and Coalition Members
    coalitions_export = pd.ExcelFile(export_dir + '/' + "Coalition_Export.xlsx")
    coa_data = pd.read_excel(coalitions_export, 'Coalition Data')
    coa_data = utils.reformat(coa_data, custom_field_labels)
    # Only data clean records for SNAP-Ed
    # SNAP-Ed staff occasionally select the wrong program_area for Coalitions
    coa_data = coa_data.loc[(coa_data['program_area'] == 'SNAP-Ed') |
                            (coa_data['reported_by_email'].isin(snap_ed_staff['E-MAIL'])) |
                            (coa_data['reported_by_email'].isin(
                                former_snap_ed_staff['email']))]  # Filtering for former staff will include transfers
    coa_members = pd.read_excel(coalitions_export, 'Members')

    # Import list of Illinois names, used to flag Coalition Members with individual's names
    # Source: https://www.ssa.gov/oact/babynames/state/
    il_names = pd.read_csv(names_list,
                           delimiter=",",
                           names=['state', 'sex', 'year', 'name', 'frequency'])
    il_names = il_names['name'].drop_duplicates()
    il_names = il_names.astype(str) + ' '

    # Import Indirect Activity data and Intervention Channels
    indirect_activities_export = pd.ExcelFile(export_dir + "Indirect_Activity_Export.xlsx")
    ia_data = pd.read_excel(indirect_activities_export, 'Indirect Activity Data')
    ia_data = utils.reformat(ia_data, custom_field_labels)
    # Only data clean records for SNAP-Ed
    ia_data = ia_data.loc[ia_data['program_area'] == 'SNAP-Ed']
    ia_ic = pd.read_excel(indirect_activities_export, 'Intervention Channels')

    # Import Partnerships data
    partnerships_export = pd.ExcelFile(export_dir + "Partnership_Export.xlsx")
    part_data = pd.read_excel(partnerships_export, 'Partnership Data')
    part_data = utils.reformat(part_data, custom_field_labels)
    # Only data clean records for SNAP-Ed
    # SNAP-Ed staff occasionally select the wrong program_area for Partnerships
    part_data = part_data.loc[(part_data['program_area'] == 'SNAP-Ed') |
                              (part_data['reported_by_email'].isin(snap_ed_staff['E-MAIL'])) |
                              (part_data['reported_by_email'].isin(
                                  former_snap_ed_staff['email']))]  # Filtering for former staff will include transfers

    # Import Program Activity data and Sessions
    program_activities_export = pd.ExcelFile(export_dir + "program_activities_export.xlsx")
    pa_data = pd.read_excel(program_activities_export, 'Program Activity Data')
    pa_data = utils.reformat(pa_data, custom_field_labels)
    # Subset Program Activities for Family Consumer Science
    pa_data_fcs = pa_data.loc[pa_data['program_areas'].str.contains('Family Consumer Science')]
    # Subset Program Activities for SNAP-Ed
    pa_data = pa_data.loc[pa_data['program_areas'].str.contains('SNAP-Ed')]
    pa_sessions = pd.read_excel(program_activities_export, 'Sessions')

    # Import PSE Site Activity data, Needs, Readiness, Effectiveness, and Changes
    pse_site_activities_export = pd.ExcelFile(export_dir + "PSE_Site_Activity_Export.xlsx")
    pse_data = pd.read_excel(pse_site_activities_export, 'PSE Data')
    pse_data = utils.reformat(pse_data, custom_field_labels)
    pse_nre = pd.read_excel(pse_site_activities_export, 'Needs, Readiness, Effectiveness')
    pse_changes = pd.read_excel(pse_site_activities_export, 'Changes')

    # Import Update Notifications, used for the Corrections Report
    update_notes = pd.read_excel(update_notifications,
                                 sheet_name='Monthly Data Cleaning').drop(columns=['Tab'])

    # Monthly PEARS Data Cleaning

    # Timestamp and report year bounds used to filter data to clean
    ts = pd.to_datetime("today")
    report_year_start = '10/01/2021'
    report_year_end = '09/30/2022'

    # Coalitions

    # Convert counties to units for use in update notification email
    coa_data = counties_to_units(data=coa_data, unit_field='coalition_unit', unit_counties=unit_counties)

    # Filter out test records, select relevant columns
    coa_data = coa_data.loc[~coa_data['coalition_name'].str.contains('(?i)TEST', regex=True),
                            ['coalition_id',
                             'coalition_name',
                             'reported_by',
                             'reported_by_email',
                             'created',
                             'modified',
                             'coalition_unit',
                             'action_plan_name',
                             'program_area',
                             'snap_ed_grant_goals']]

    # Set Coalition data cleaning flags

    coa_data['GENERAL INFORMATION TAB UPDATES'] = np.nan

    coa_data['GI UPDATE1'] = np.nan
    coa_data.loc[coa_data['action_plan_name'].isnull(),
                 'GI UPDATE1'] = get_update_note(update_notes, module='Coalitions', update='GI UPDATE1')

    coa_data['GI UPDATE2'] = np.nan
    coa_data.loc[coa_data['program_area'] != 'SNAP-Ed',
                 'GI UPDATE2'] = get_update_note(update_notes, module='Coalitions', update='GI UPDATE2')

    # Concatenate General Information tab updates
    coa_data = concat_updates(coa_data,
                              concat_col='GENERAL INFORMATION TAB UPDATES',
                              update_cols=['GI UPDATE1', 'GI UPDATE2'])

    coa_data['CUSTOM DATA TAB UPDATES'] = np.nan
    coa_data.loc[coa_data['snap_ed_grant_goals'].isnull(),
                 'CUSTOM DATA TAB UPDATES'] = get_update_note(update_notes,
                                                              module='Coalitions',
                                                              update='CUSTOM DATA TAB UPDATES')

    coa_data['COALITION MEMBERS TAB UPDATES'] = np.nan

    # Count Coalition Members of each Coalition, flag Coalitions that have none
    coa_data['CM UPDATE1'] = np.nan
    coa_members_count = coa_members.groupby('coalition_id')['member_id'].count().reset_index(name='# of Members')
    coa_data = pd.merge(coa_data, coa_members_count, how='left', on='coalition_id')
    coa_data.loc[(coa_data['# of Members'].isnull()) | (coa_data['# of Members'] == 0),
                 'CM UPDATE1'] = get_update_note(update_notes, module='Coalitions', update='CM UPDATE1')

    # Subsequent updates require Members data
    coa_members_data = pd.merge(coa_data, coa_members, how='left', on='coalition_id').rename(
        columns={'name': 'member_name'})

    coa_members_data['CM UPDATE2'] = np.nan
    coa_members_data.loc[
        (coa_members_data['type'] != 'Community members/individuals') & (coa_members_data['site_id'].isnull()),
        'CM UPDATE2'] = get_update_note(update_notes, module='Coalitions', update='CM UPDATE2')

    # Flag Coalition Members that contain individuals' names
    coa_members_data['CM UPDATE3'] = np.nan
    # Terms indicating false positives
    exclude_terms = ['University', 'Hospital', 'YMCA', 'Center', 'County', 'Elementary', 'Foundation', 'Church', 'Club',
                     'Daycare', 'Housing', 'SNAP-Ed']
    coa_members_data.loc[(coa_members_data['member_name'].str.contains('|'.join(il_names), na=False)) &
                         (coa_members_data['member_name'].str.count(' ') == 1) &
                         (~coa_members_data['member_name'].str.contains('|'.join(exclude_terms), na=False)),
                         'CM UPDATE3'] = get_update_note(update_notes, module='Coalitions', update='CM UPDATE3')

    # Concatenate Coalition Members tab updates
    coa_members_data = concat_updates(coa_members_data,
                                      concat_col='COALITION MEMBERS TAB UPDATES',
                                      update_cols=['CM UPDATE1', 'CM UPDATE2', 'CM UPDATE3'])

    # Subset records that require updates
    coa_corrections = coa_members_data.loc[coa_members_data.filter(like='UPDATE').notnull().any(1)]
    coa_corrections = drop_child_dupes(coa_corrections,
                                       c_updates=['CM UPDATE2', 'CM UPDATE3'],
                                       parent_id='coalition_id',
                                       child_id='member_id')
    # coa_corrections is exported in the Corrections Report

    # Reformat coa_corrections for the update notification email body
    coa_corrections_email = corrections_email_format(coa_corrections,
                                                     cols=['coalition_id',
                                                           'coalition_name',
                                                           'reported_by',
                                                           'reported_by_email',
                                                           'coalition_unit',
                                                           'GENERAL INFORMATION TAB UPDATES',
                                                           'action_plan_name',
                                                           'program_area',
                                                           'CUSTOM DATA TAB UPDATES',
                                                           'COALITION MEMBERS TAB UPDATES',
                                                           '# of Members',
                                                           'member_name',
                                                           'site_id'],
                                                     index=['coalition_id'],
                                                     int_cols=['site_id'],
                                                     rename_cols={'coalition_unit': 'unit'},
                                                     update_cols=['GENERAL INFORMATION TAB UPDATES',
                                                                  'COALITION MEMBERS TAB UPDATES'])

    # Indirect Activities

    # Set Indirect Activity data cleaning flags

    # Convert counties to units for use in update notification email
    ia_data = counties_to_units(data=ia_data, unit_field='unit', unit_counties=unit_counties)

    # Filter out test records, select relevant columns
    ia_data = ia_data.loc[~ia_data['title'].str.contains('(?i)TEST', regex=True),
                          ['activity_id',
                           'title',
                           'reported_by',
                           'reported_by_email',
                           'created',
                           'modified',
                           'start_date',
                           'end_date',
                           'unit',
                           'type',
                           'snap_ed_grant_goals']]

    # Set Indirect Activity data cleaning flags

    ia_data['CUSTOM DATA TAB UPDATES'] = np.nan

    ia_data.loc[ia_data.duplicated(subset=['reported_by_email', 'type'], keep=False),
                'CUSTOM DATA TAB UPDATES'] = get_update_note(update_notes,
                                                             module='Indirect Activities',
                                                             update='CUSTOM DATA TAB UPDATES',
                                                             notification='Notification1')
    ia_data.loc[ia_data['snap_ed_grant_goals'].isnull(),
                'CUSTOM DATA TAB UPDATES'] = get_update_note(update_notes,
                                                             module='Indirect Activities',
                                                             update='CUSTOM DATA TAB UPDATES',
                                                             notification='Notification2')

    # Filter out test records, select relevant columns
    ia_ic = ia_ic.loc[~ia_ic['activity'].str.contains('(?i)TEST', regex=True),
                      ['activity_id',
                       'activity',
                       'channel_id',
                       'channel',
                       'description',
                       'site_id',
                       'site_name',
                       'reach',
                       'newly_reached']]
    ia_ic['INTERVENTION CHANNELS AND REACH TAB UPDATES'] = np.nan

    # Subsequent updates require Intervention Channels data
    ia_ic_data = pd.merge(ia_data, ia_ic, how='left', on='activity_id')

    # Flag Intervention Channels that don't contain a date in their description
    ia_ic_data['IC UPDATE1'] = np.nan
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    ia_ic_data['description'] = ia_ic_data['description'].astype(str)
    ia_ic_data.loc[ia_ic_data['description'] == 'nan', 'description'] = ''
    ia_ic_data.loc[~ia_ic_data['description'].str.contains('|'.join(months)) & ~ia_ic_data['description'].str.contains(
        r'\d+/|-|.{1}\d{2,4}'),
                   'IC UPDATE1'] = get_update_note(update_notes,
                                                   module='Indirect Activities',
                                                   update='Indirect Activities')

    ia_ic_data['IC UPDATE2'] = np.nan
    # ia_ic_data.loc[(ia_ic_data['site_name'].isnull()) | (ia_ic_data['site_name'] == 'abc placeholder'),
    #                'IC UPDATE2'] = get_update_note(update_notes, module='Indirect Activities', update='IC UPDATE2')
    # As of 3/18/22, Indirect Activity site is set to a required field in PEARS for Illinois Extension

    ia_ic_data['IC UPDATE3'] = np.nan
    ia_ic_data.loc[(ia_ic_data['newly_reached'].notnull())
                   & (ia_ic_data['newly_reached'] != 0),
                   'IC UPDATE3'] = get_update_note(update_notes, module='Indirect Activities', update='IC UPDATE3')

    ia_ic_data['IC UPDATE4'] = np.nan
    ia_ic_data.loc[(ia_ic_data['reach'].isnull()) | (ia_ic_data['reach'] == 0),
                   'IC UPDATE4'] = get_update_note(update_notes, module='Indirect Activities', update='IC UPDATE4')
    # How are null reach values possible?
    # Reach is null if there are no Intervention Channels for the Indirect Activity

    ia_ic_data['IC UPDATE5'] = np.nan
    ia_ic_data.loc[ia_ic_data.duplicated(subset=['activity_id', 'description', 'site_id'], keep=False),
                   'IC UPDATE5'] = get_update_note(update_notes, module='Indirect Activities', update='IC UPDATE5')

    ia_ic_data['IC UPDATE6'] = np.nan
    hard_copy_materials = "Hard copy materials (e.g. flyers, pamphlets, activity books, posters, banners, postcards, " \
                          "recipe cards, or newsletters for mailings)"
    ia_ic_data.loc[ia_ic_data['channel'] == hard_copy_materials,
                   'IC UPDATE6'] = get_update_note(update_notes, module='Indirect Activities', update='IC UPDATE6')

    # Concatenate Intervention Channels and Reach tab updates
    ia_ic_data = concat_updates(ia_ic_data,
                                concat_col='INTERVENTION CHANNELS AND REACH TAB UPDATES',
                                update_cols=['IC UPDATE1',
                                             'IC UPDATE2',
                                             'IC UPDATE3',
                                             'IC UPDATE4',
                                             'IC UPDATE5',
                                             'IC UPDATE6'])

    # Subset records that require updates
    ia_corrections = ia_ic_data.loc[ia_ic_data.filter(like='UPDATE').notnull().any(1)]
    ia_corrections = drop_child_dupes(ia_corrections,
                                      c_updates=ia_ic_data.columns[ia_ic_data.columns.str.contains('IC')].tolist(),
                                      parent_id='activity_id',
                                      child_id='channel_id')
    # ia_corrections is exported in the Corrections Report

    # Reformat ia_corrections for the update notification email body
    ia_corrections_email = corrections_email_format(ia_corrections,
                                                    cols=['activity_id',
                                                          'title',
                                                          'reported_by',
                                                          'reported_by_email',
                                                          'unit',
                                                          'CUSTOM DATA TAB UPDATES',
                                                          'type',
                                                          'INTERVENTION CHANNELS AND REACH TAB UPDATES',
                                                          'channel_id',
                                                          'channel',
                                                          'description',
                                                          'site_name',
                                                          'newly_reached',
                                                          'reach'],
                                                    index='activity_id',
                                                    int_cols=['newly_reached', 'channel_id'],
                                                    update_cols=['INTERVENTION CHANNELS AND REACH TAB UPDATES'])

    # Partnerships

    # Convert counties to units for use in update notification email
    part_data = counties_to_units(data=part_data, unit_field='partnership_unit', unit_counties=unit_counties)

    # Filter out test records, select relevant columns
    part_data = part_data.loc[~part_data['partnership_name'].str.contains('(?i)TEST', regex=True),
                              ['partnership_id',
                               'partnership_name',
                               'reported_by',
                               'reported_by_email',
                               'created',
                               'modified',
                               'partnership_unit',
                               'action_plan_name',
                               'is_direct_education_intervention',
                               'program_area',
                               'is_pse_intervention',
                               'relationship_depth',
                               'snap_ed_grant_goals']]

    # Set Partnerships data cleaning flags

    part_data['GENERAL INFORMATION TAB UPDATES'] = np.nan

    part_data['GI UPDATE1'] = np.nan
    part_data.loc[part_data['action_plan_name'].isnull(),
                  'GI UPDATE1'] = get_update_note(update_notes, module='Partnerships', update='GI UPDATE1')

    part_data['GI UPDATE2'] = np.nan
    part_data.loc[(part_data['is_direct_education_intervention'] == 0) & (part_data['is_pse_intervention'] == 0),
                  'GI UPDATE2'] = get_update_note(update_notes, module='Partnerships', update='GI UPDATE2')

    part_data['GI UPDATE3'] = np.nan
    part_data.loc[part_data['program_area'] != 'SNAP-Ed',
                  'GI UPDATE3'] = get_update_note(update_notes, module='Partnerships', update='GI UPDATE3')

    # Concatenate General Information tab updates
    part_data = concat_updates(part_data,
                               concat_col='GENERAL INFORMATION TAB UPDATES',
                               update_cols=['GI UPDATE1', 'GI UPDATE2', 'GI UPDATE3'])

    part_data['CUSTOM DATA TAB UPDATES'] = np.nan
    part_data.loc[part_data['snap_ed_grant_goals'].isnull(),
                  'CUSTOM DATA TAB UPDATES'] = get_update_note(update_notes,
                                                               module='Partnerships',
                                                               update='CUSTOM DATA TAB UPDATES')

    part_data['EVALUATION TAB UPDATES'] = np.nan
    part_data.loc[part_data['relationship_depth'].isnull(),
                  'EVALUATION TAB UPDATES'] = get_update_note(update_notes,
                                                              module='Partnerships',
                                                              update='EVALUATION TAB UPDATES')

    # Subset records that require updates
    part_corrections = part_data.loc[part_data.filter(like='UPDATE').notnull().any(1)]
    # part_corrections is exported in the Corrections Report

    # Reformat part_corrections for the update notification email body
    part_corrections_email = corrections_email_format(part_corrections,
                                                      cols=['partnership_id',
                                                            'partnership_name',
                                                            'reported_by',
                                                            'reported_by_email',
                                                            'partnership_unit',
                                                            'GENERAL INFORMATION TAB UPDATES',
                                                            'action_plan_name',
                                                            'program_area',
                                                            'CUSTOM DATA TAB UPDATES',
                                                            'EVALUATION TAB UPDATES',
                                                            'relationship_depth'],
                                                      index='partnership_id',
                                                      rename_cols={'partnership_unit': 'unit'},
                                                      update_cols=['GENERAL INFORMATION TAB UPDATES'])

    # Program Activities

    # Set Program Activities data cleaning flags

    # Select relevant columns
    pa_sessions = pa_sessions.loc[:, ['session_id',
                                      'program_id',
                                      'start_date',
                                      'start_date_with_time',
                                      'length',
                                      'num_participants']]

    pa_sessions['GENERAL INFORMATION TAB UPDATES'] = np.nan

    pa_sessions['GI UPDATE1'] = np.nan
    pa_sessions['start_date'] = pd.to_datetime(pa_sessions['start_date'])
    pa_sessions.loc[(pa_sessions['start_date'] < report_year_start) | (pa_sessions['start_date'] > report_year_end),
                    'GI UPDATE1'] = get_update_note(update_notes, module='Program Activities', update='GI UPDATE1')

    pa_sessions['GI UPDATE2'] = np.nan
    pa_sessions.loc[(pa_sessions['start_date_with_time'] < ts)
                    & (pa_sessions['num_participants'].isnull()),
                    'GI UPDATE2'] = get_update_note(update_notes,
                                                    module='Program Activities',
                                                    update='GI UPDATE2',
                                                    notification='Notification1')
    pa_sessions.loc[(pa_sessions['start_date_with_time'] < ts)
                    & (pa_sessions['num_participants'] == 0),
                    'GI UPDATE2'] = get_update_note(update_notes,
                                                    module='Program Activities',
                                                    update='GI UPDATE2',
                                                    notification='Notification2')

    pa_sessions['GI UPDATE3'] = np.nan
    pa_sessions.loc[pa_sessions.duplicated(subset=['program_id', 'start_date_with_time'], keep=False),
                    'GI UPDATE3'] = get_update_note(update_notes, module='Program Activities', update='GI UPDATE3')

    # Convert counties to units for use in update notification email
    pa_data = counties_to_units(data=pa_data, unit_field='unit', unit_counties=unit_counties)

    # Filter out test records, select relevant columns
    pa_data = pa_data.loc[(~pa_data['name'].str.contains('(?i)TEST', regex=True))
                          & (pa_data['name'] != 'abc placeholder'),
                          ['program_id',
                           'reported_by',
                           'reported_by_email',
                           'created',
                           'modified',
                           'is_complete',
                           'name',
                           'unit',
                           'start_date',
                           'end_date',
                           'intervention',
                           'setting',
                           'primary_curriculum',
                           'participants_total',
                           'comments',
                           'snap_ed_grant_goals',
                           'snap_ed_special_projects']]

    # Subsequent updates require Program Activity data
    pa_sessions_data = pd.merge(pa_data, pa_sessions, how='left', on='program_id', suffixes=('_PA', '_Session'))

    pa_sessions_data['GI UPDATE4'] = np.nan
    pa_sessions_data.loc[(pa_sessions_data['length'] < 20) | (pa_sessions_data['length'].isnull()),
                         'GI UPDATE4'] = get_update_note(update_notes, module='Program Activities', update='GI UPDATE4')

    # Concatenate General Information tab updates
    pa_sessions_data = concat_updates(pa_sessions_data,
                                      concat_col='GENERAL INFORMATION TAB UPDATES',
                                      update_cols=['GI UPDATE1', 'GI UPDATE2', 'GI UPDATE3', 'GI UPDATE4'])

    pa_sessions_data['CUSTOM DATA TAB UPDATES'] = np.nan
    pa_sessions_data.loc[pa_sessions_data['snap_ed_grant_goals'].isnull(),
                         'CUSTOM DATA TAB UPDATES'] = get_update_note(update_notes,
                                                                      module='Program Activities',
                                                                      update='CUSTOM DATA TAB UPDATES',
                                                                      notification='Notification1')
    pa_sessions_data.loc[(pa_sessions_data['snap_ed_special_projects'].str.contains('None'))
                         & (pa_sessions_data['snap_ed_special_projects'] != 'None'),
                         'CUSTOM DATA TAB UPDATES'] = get_update_note(update_notes,
                                                                      module='Program Activities',
                                                                      update='CUSTOM DATA TAB UPDATES',
                                                                      notification='Notification2')

    pa_sessions_data['SNAP-ED CUSTOM DATA TAB UPDATES'] = np.nan

    pa_sessions_data['SCD UPDATE1'] = np.nan
    pa_sessions_data.loc[pa_sessions_data['intervention'].isnull(),
                         'SCD UPDATE1'] = get_update_note(update_notes,
                                                          module='Program Activities',
                                                          update='SCD UPDATE1',
                                                          notification='Notification1')
    pa_sessions_data.loc[pa_sessions_data['intervention'] != 'SNAP-Ed Community Network',
                         'SCD UPDATE1'] = get_update_note(update_notes,
                                                          module='Program Activities',
                                                          update='SCD UPDATE1',
                                                          notification='Notification2')

    pa_sessions_data['SCD UPDATE2'] = np.nan
    settings_other = ['Other places people', 'Other settings people']
    pa_sessions_data.loc[pa_sessions_data['setting'].str.contains('|'.join(settings_other), na=False),
                         'SCD UPDATE2'] = get_update_note(update_notes,
                                                          module='Program Activities',
                                                          update='SCD UPDATE2')

    # Concatenate Snap-Ed Custom Data tab updates
    pa_sessions_data = concat_updates(pa_sessions_data,
                                      concat_col='SNAP-ED CUSTOM DATA TAB UPDATES',
                                      update_cols=['SCD UPDATE1', 'SCD UPDATE2'])

    # Flag Program Activities where the unique participants is equal to the sum of session participants
    pa_sessions_data['DEMOGRAPHICS TAB UPDATES'] = np.nan
    pa_sessions_metrics = pa_sessions.groupby('program_id').agg(
        {'session_id': 'count',
         'num_participants': 'sum'}).reset_index().rename(
        columns={'session_id': '# of Sessions',
                 'num_participants': 'Total Session Participants'})
    pa_sessions_data = pd.merge(pa_sessions_data, pa_sessions_metrics, how='left', on='program_id')
    pa_sessions_data.loc[(pa_sessions_data['# of Sessions'] > 1) &
                         (pa_sessions_data['Total Session Participants'] == pa_sessions_data['participants_total']),
                         'DEMOGRAPHICS TAB UPDATES'] = get_update_note(update_notes,
                                                                       module='Program Activities',
                                                                       update='DEMOGRAPHICS TAB UPDATES')
    # End of year:
    # For entries with only 1 session, the total # of session participants should = total # of unique participants.

    # Data clean FCS Program Activities
    pa_data_fcs['CUSTOM DATA TAB UPDATES'] = np.nan
    pa_data_fcs.loc[(pa_data_fcs['fcs_program_team'].str.contains('SNAP-Ed'))
                    & (pa_data_fcs['fcs_grant_goals'].isnull()),
                    'CUSTOM DATA TAB UPDATES'] = get_update_note(update_notes,
                                                                 module='Program Activities',
                                                                 update='CUSTOM DATA TAB UPDATES')

    # Append FCS Program Activities to SNAP-Ed Program Activities
    add_cols = pa_sessions_data.columns[~pa_sessions_data.columns.isin(pa_data_fcs.columns)].tolist()
    pa_data_fcs = pd.concat([pa_data_fcs, pd.DataFrame(columns=add_cols)])  # turns program_id into float
    pa_data_fcs['program_id'] = pa_data_fcs['program_id'].astype(int)
    sub_cols = pa_sessions_data.columns[pa_sessions_data.columns.isin(pa_data_fcs.columns)].tolist()
    pa_sessions_data = pa_sessions_data.append(pa_data_fcs[sub_cols], ignore_index=True)
    # possible dupes added?

    # Subset records that require updates
    pa_corrections = pa_sessions_data.loc[pa_sessions_data.filter(like='UPDATE').notnull().any(1)]
    pa_corrections = drop_child_dupes(pa_corrections,
                                      c_updates=['GI UPDATE1', 'GI UPDATE2', 'GI UPDATE4'],
                                      parent_id='program_id',
                                      child_id='session_id')
    # pa_corrections is exported in the Corrections Report

    # Reformat pa_corrections for the update notification email body
    pa_corrections_email = corrections_email_format(pa_corrections,
                                                    cols=['program_id',
                                                          'name',
                                                          'reported_by',
                                                          'reported_by_email',
                                                          'unit',
                                                          'GENERAL INFORMATION TAB UPDATES',
                                                          'session_id',
                                                          'start_date_with_time',
                                                          'length',
                                                          'num_participants',
                                                          'CUSTOM DATA TAB UPDATES',
                                                          'SNAP-ED CUSTOM DATA TAB UPDATES',
                                                          'intervention',
                                                          'setting',
                                                          'DEMOGRAPHICS TAB UPDATES',
                                                          'primary_curriculum'],
                                                    index='program_id',
                                                    int_cols=['num_participants'],
                                                    update_cols=['GENERAL INFORMATION TAB UPDATES',
                                                                 'SNAP-ED CUSTOM DATA TAB UPDATES'],
                                                    datetime_cols=['start_date_with_time'])

    # PSE Site Activities

    # Convert counties to units for use in update notification email
    pse_data = counties_to_units(data=pse_data, unit_field='pse_unit', unit_counties=unit_counties)

    # Filter out test records, select relevant columns
    pse_data['name'] = pse_data['name'].astype(str)
    pse_data = pse_data.loc[
        (~pse_data['name'].str.contains('(?i)TEST', regex=True, na=False))
        & (pse_data['site_name'] != 'abc placeholder'),
        ['pse_id',
         'site_id',
         'site_name',
         'name',
         'reported_by',
         'reported_by_email',
         'created',
         'modified',
         'setting',
         'start_fiscal_year',
         'planning_stage_sites_contacted_and_agreed_to_participate',
         'total_reach',
         'pse_unit',
         'program_area',
         'intervention',
         'snap_ed_grant_goals']]

    # Set PSE data cleaning flags

    pse_data['GENERAL INFORMATION TAB UPDATES'] = np.nan

    pse_data['GI UPDATE1'] = np.nan
    pse_data.loc[(pse_data['start_fiscal_year'] != 2022) & (
            pse_data['planning_stage_sites_contacted_and_agreed_to_participate'] == 1),
                 'GI UPDATE1'] = get_update_note(update_notes, module='PSE Site Activities', update='GI UPDATE1')

    pse_data['GI UPDATE2'] = np.nan
    pse_data.loc[pse_data['program_area'] == 'Family Consumer Science',
                 'GI UPDATE2'] = get_update_note(update_notes, module='PSE Site Activities', update='GI UPDATE2')

    pse_data['GI UPDATE3'] = np.nan
    pse_data.loc[(pse_data['site_id'].duplicated(keep=False))
                 & (pse_data['pse_unit'] != 'CPHP'),
                 'GI UPDATE3'] = get_update_note(update_notes, module='PSE Site Activities', update='GI UPDATE3')

    pse_data['GI UPDATE4'] = np.nan
    pse_data.loc[pse_data['intervention'] != 'SNAP-Ed Community Network',
                 'GI UPDATE4'] = get_update_note(update_notes, module='PSE Site Activities', update='GI UPDATE4')

    # Concatenate General Information tab updates
    pse_data = concat_updates(pse_data,
                              concat_col='GENERAL INFORMATION TAB UPDATES',
                              update_cols=['GI UPDATE1', 'GI UPDATE2', 'GI UPDATE3', 'GI UPDATE4'])

    pse_data['CUSTOM DATA TAB UPDATES'] = np.nan
    pse_data.loc[pse_data['snap_ed_grant_goals'].isnull(),
                 'CUSTOM DATA TAB UPDATES'] = get_update_note(update_notes,
                                                              module='PSE Site Activities',
                                                              update='CUSTOM DATA TAB UPDATES')

    # Select relevant Needs, Readiness, Effectiveness columns
    pse_nre = pse_nre.loc[:, ['pse_id',
                              'assessment_id',
                              'assessment_type',
                              'assessment_tool',
                              'baseline_score',
                              'baseline_date',
                              'follow_up_date',
                              'follow_up_score']]

    # Subsequent updates require Needs, Readiness, Effectiveness data
    pse_nre_data = pd.merge(pse_data, pse_nre, how='left', on='pse_id')

    pse_nre_data['NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'] = np.nan

    pse_nre_data['NRE UPDATE1'] = np.nan
    pse_nre_data.loc[(pse_nre_data['assessment_type'] == 'Needs assessment/environmental scan')
                     & (pse_nre_data['baseline_score'].isnull())
                     & (~pse_nre_data['assessment_tool'].str.contains('SLAQ', na=False)),
                     'NRE UPDATE1'] = get_update_note(update_notes, module='PSE Site Activities', update='NRE UPDATE1')

    pse_nre_data['NRE UPDATE2'] = np.nan
    pse_nre_data['baseline_date'] = pd.to_datetime(pse_nre_data['baseline_date'])
    pse_nre_data.loc[(pse_nre_data['assessment_type'] == 'Needs assessment/environmental scan')
                     & (pse_nre_data['baseline_date'].isnull()),
                     'NRE UPDATE2'] = get_update_note(update_notes, module='PSE Site Activities', update='NRE UPDATE2')

    pse_nre_data['NRE UPDATE3'] = np.nan
    pse_nre_data['follow_up_date'] = pd.to_datetime(pse_nre_data['follow_up_date'])
    pse_nre_data.loc[(pse_nre_data['assessment_type'] == 'Needs assessment/environmental scan')
                     & (pse_nre_data['follow_up_date'].notnull())
                     & (pse_nre_data['follow_up_score'].isnull()),
                     'NRE UPDATE3'] = get_update_note(update_notes, module='PSE Site Activities', update='NRE UPDATE3')

    pse_nre_data['NRE UPDATE4'] = np.nan
    pse_nre_data.loc[(pse_nre_data['assessment_type'] == 'Needs assessment/environmental scan')
                     & (pse_nre_data['follow_up_date'].isnull())
                     & (pse_nre_data['follow_up_score'].notnull()),
                     'NRE UPDATE4'] = get_update_note(update_notes, module='PSE Site Activities', update='NRE UPDATE4')

    # Concatenate Needs, Readiness, Effectiveness tab updates
    pse_nre_data = concat_updates(pse_nre_data,
                                  concat_col='NEEDS, READINESS & EFFECTIVENESS TAB UPDATES',
                                  update_cols=['NRE UPDATE1', 'NRE UPDATE2', 'NRE UPDATE3', 'NRE UPDATE4'])

    # Subsequent updates require Changes Adopted data
    pse_nre_changes_data = pd.merge(pse_nre_data, pse_changes[['pse_id', 'change_id']], how='left',
                                    on='pse_id').drop_duplicates(subset=['pse_id', 'assessment_id'])

    pse_nre_changes_data['CHANGES ADOPTED TAB UPDATES'] = np.nan
    pse_nre_changes_data.loc[(pse_nre_changes_data['change_id'].notnull())
                             & (pse_nre_changes_data['total_reach'].isnull()),
                             'CHANGES ADOPTED TAB UPDATES'] = get_update_note(update_notes,
                                                                              module='PSE Site Activities',
                                                                              update='CHANGES ADOPTED TAB UPDATES')

    # Subset records that require updates
    pse_corrections = pse_nre_changes_data.loc[pse_nre_changes_data.filter(like='UPDATE').notnull().any(1)]
    # pse_corrections is exported in the Corrections Report

    # Reformat pse_corrections for the update notification email body
    pse_corrections_email = corrections_email_format(pse_corrections,
                                                     cols=['pse_id',
                                                           'site_name',
                                                           'reported_by',
                                                           'reported_by_email',
                                                           'pse_unit',
                                                           'GENERAL INFORMATION TAB UPDATES',
                                                           'start_fiscal_year',
                                                           'planning_stage_sites_contacted_and_agreed_to_participate',
                                                           'program_area',
                                                           'site_id',
                                                           'intervention',
                                                           'CUSTOM DATA TAB UPDATES',
                                                           'NEEDS, READINESS & EFFECTIVENESS TAB UPDATES',
                                                           'baseline_score',
                                                           'baseline_date',
                                                           'follow_up_date',
                                                           'follow_up_score',
                                                           'CHANGES ADOPTED TAB UPDATES',
                                                           'total_reach'],
                                                     index='pse_id',
                                                     rename_cols={'pse_unit': 'unit'},
                                                     update_cols=['GENERAL INFORMATION TAB UPDATES',
                                                                  'NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'],
                                                     date_cols=['baseline_date', 'follow_up_date'])

    # Corrections Report

    # Summarize and concatenate module corrections
    corrections_dict = {
        'Coalitions': coa_corrections,
        'Indirect Activities': ia_corrections,
        'Partnerships': part_corrections,
        'Program Activities': pa_corrections,
        'PSE Site Activities': pse_corrections
    }

    module_sums = [corrections_sum(corrections, module) for module, corrections in corrections_dict.items()]

    corrections_sums = pd.concat(module_sums, ignore_index=True)
    corrections_sums.insert(0, 'Module', corrections_sums.pop('Module'))
    corrections_sums = pd.merge(corrections_sums, update_notes, how='left', on=['Module', 'Update'])

    # Calculate the month for this report
    prev_month = (ts - pd.DateOffset(months=1)).to_period('M')

    # Export the Corrections Report as an Excel file

    corrections_report_filename = 'Monthly PEARS Corrections ' + prev_month.strftime('%Y-%m') + '.xlsx'
    corrections_report_path = output_dir + '/' + corrections_report_filename

    dfs = {'Corrections Summary': corrections_sums,
           'Coalitions': coa_corrections,
           'Indirect Activities': ia_corrections,
           'Partnerships': part_corrections,
           'Program Activities': pa_corrections,
           'PSE': pse_corrections}

    # Create function for write_corrections_report
    writer = pd.ExcelWriter(file_path1, engine='xlsxwriter')
    for sheetname, df in dfs.items():  # loop through `dict` of dataframes
        df.to_excel(writer, sheet_name=sheetname, index=False, freeze_panes=(1, 0))  # send df to writer
        worksheet = writer.sheets[sheetname]  # pull worksheet object
        workbook = writer.book
        worksheet.autofilter(0, 0, 0, len(df.columns) - 1)
        blue_bold = workbook.add_format({'bold': True, 'bg_color': '#DEEAF0', 'font_color': '#000000'})
        worksheet.conditional_format(0, 0, len(df), 2, {'type': 'formula', 'criteria': '=$B1="Total"', 'format': blue_bold})
        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)
    writer.close()

    # Email Update Notifications

    if send_emails:

        # Set the deadline for when updates are due
        deadline_date = ts.date().replace(day=19).strftime('%A %b %d, %Y')

        # Update Notification email body template
        html = """<html>
          <head></head>
        <body>
                    <p>
                    Hello {0},<br><br>

                    A few of your PEARS entries need edits. Please update the entries listed in the table(s) below by <b>5:00pm {1}</b>.
                    Records not corrected by then will continue to show up on monthly PEARS notifications until they are resolved.
                    <ul>
                      <li>For each entry listed, please make the edit(s) written in the columns labeled <b>UPDATE</b> in the column heading.</li>
                      <li>You can locate entries in PEARS by entering their IDs into the search filter.</li>
                      <li>To edit a PEARS entry previously marked as “complete,” you can mark the entry as “incomplete,”
                          edit the record, and then mark as “complete” again.</li>
                      <li>As a friendly reminder – following the Cheat Sheets
                          <a href="https://uofi.app.box.com/folder/49632670918?s=wwymjgjd48tyl0ow20vluj196ztbizlw">[Located Here]</a>
                          will help to prevent future PEARS corrections.</li>
                  </ul>

                    {2}
                    <br>{3}<br>
                    {4}<br>
                    {5}<br>
                    {6}<br>
                    {7}<br>
                    </p>
          </body>
        </html>
        """

        # Create dataframe of staff to notify
        Module_Corrections2 = [coa_corrections_email, ia_corrections_email, part_corrections_email, pa_corrections_email, pse_corrections_email]
        notify_staff = pd.DataFrame()

        for df in Module_Corrections2:
            notify_staff = notify_staff.append(df[['reported_by', 'reported_by_email', 'unit']], ignore_index=True)

        notify_staff = notify_staff.sort_values(['reported_by', 'unit']).drop_duplicates(
            subset=['reported_by', 'reported_by_email'], keep='first').reset_index(drop=True)

        # Subset current staff using the staff list
        current_staff = notify_staff.loc[
            notify_staff['reported_by_email'].isin(staff['email']), ['reported_by', 'reported_by_email', 'unit']]
        current_staff = current_staff.values.tolist()


        # Verify emails?

        # If email fails to send, the recipient is added to this list
        failed_recipients = []

        # Email Update Notifications to current staff

        for x in current_staff:

            Coa_df = staff_corrections(coa_corrections_email, former=False, staff_email=x[1])
            IA_df = staff_corrections(ia_corrections_email, former=False, staff_email=x[1])
            Part_df = staff_corrections(part_corrections_email, former=False, staff_email=x[1])
            PA_df = staff_corrections(pa_corrections_email, former=False, staff_email=x[1])
            PSE_df = staff_corrections(pse_corrections_email, former=False, staff_email=x[1])

            staff_name = x[0]
            send_to = x[1]
            unit = x[2]

            # If the recipient's unit is an INEP unit, they are directed to contact their Regional Specialist
            # Else, they are directed to contact the FCS Evaluation team

            response_tag = """If you have any questions or need help please reply to this email and a member of the FCS Evaluation Team will reach out soon.
                    <br>Thanks and have a great day!<br>

                    <br> <b> FCS Evaluation Team </b> <br>
                    <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
            """

            new_Cc = notification_cc

            if ((unit in re_lookup["UNIT #"].tolist()) and
                (send_to not in state_staff['E-MAIL'].tolist()) and
                ('@uic.edu' not in send_to) and
                (re_lookup.loc[re_lookup['UNIT #'] == unit].empty == False)):
                response_tag = 'If you have any questions or need help please contact your Regional Specialist, <b>{0}</b> (<a href = "mailto: {1} ">{1}</a>).'
                re_name = re_lookup.loc[re_lookup['UNIT #'] == unit, 'REGIONAL EDUCATOR'].item()
                re_email = re_lookup.loc[re_lookup['UNIT #'] == unit, 'RE E-MAIL'].item()
                response_tag = response_tag.format(*[re_name, re_email])
                new_Cc = notification_cc + ', ' + re_email

            # Staff's first name is used in the email salutation
            first_name = staff.loc[staff['email'] == send_to, 'first_name'].item()

            subject = 'PEARS Entries Updates ' + prev_month.strftime('%b-%Y') + ', Unit ' + unit + ', ' + staff_name

            # Insert the corrections dfs into the email body
            dfs = {'Coalitions': Coa_df, 'Indirect Activities': IA_df, 'Partnerships': Part_df, 'Program Activities': PA_df,
                   'PSE Site Activities': PSE_df}
            y = [first_name, deadline_date, response_tag]
            insert_dfs(dfs, y)
            new_html = html.format(*y)

            # Try to send the email, otherwise add the recpient's email address to failed_recipients
            try:
                utils.send_mail(send_from=creds['admin_send_from'],
                                send_to=send_to,
                                cc=new_Cc,
                                subject=subject,
                                html=new_html,
                                username=creds['admin_username'],
                                password=creds['admin_password'],
                                is_tls=True)
            except smtplib.SMTPException:
                failed_recipients.append(x)

        # Email Update Notifications for former staff

        # Subset former staff using the staff list
        former_staff = notify_staff.loc[~notify_staff['reported_by_email'].isin(staff['email'])]

        Coa_df = staff_corrections(coa_corrections_email)
        IA_df = staff_corrections(ia_corrections_email)
        Part_df = staff_corrections(part_corrections_email)
        PA_df = staff_corrections(pa_corrections_email)
        PSE_df = staff_corrections(pse_corrections_email)

        # Export former staff corrections as an Excel file

        former_staff_dfs = {'Coalitions': Coa_df, 'Indirect Activities': IA_df, 'Partnerships': Part_df,
                            'Program Activities': PA_df, 'PSE': PSE_df}

        filename2 = 'Former Staff PEARS Updates ' + prev_month.strftime('%Y-%m') + '.xlsx'
        file_path2 = output_dir + '/' + filename2

        # UPDATE utils.write_report() TO ACCEPT DICT
        utils.write_report(file_path2, former_staff_dfs.keys, former_staff_dfs.values)

        # Send former staff updates email

        subject2 = 'Former Staff PEARS Updates ' + prev_month.strftime('%Y-%m')

        html2 = """<html>
          <head></head>
        <body>
                    <p>
                    Hello RECIPIENT NAME,<br><br>

                    The table(s) below compile PEARS entries created by former staff that require edits.
                    Please update the entries in each sheet by <b>5:00pm {0}</b>.
                    Records not corrected by then will continue to show up on monthly PEARS notifications until they are resolved.
                    <ul>
                      <li>For each entry listed, please make the edit(s) written in the columns labeled <b>UPDATE</b> in the column heading.</li>
                      <li>You can locate entries in PEARS by entering their IDs into the search filter.</li>
                      <li>To edit a PEARS entry previously marked as “complete,”
                          you can mark the entry as “incomplete,” edit the record, and then mark as “complete” again.</li>
                      <li>As a friendly reminder – following the Cheat Sheets
                          <a href="https://uofi.app.box.com/folder/49632670918?s=wwymjgjd48tyl0ow20vluj196ztbizlw">[Located Here]</a>
                          will help to prevent future PEARS corrections.</li>
                  </ul>
                  If you have any questions or need help please reply to this email and a member of the FCS Evaluation Team will reach out soon.

                    <br>Thanks and have a great day!<br>
                    <br> <b> FCS Evaluation Team </b> <br>
                    <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
                    <br>{1}<br>
                    {2}<br>
                    {3}<br>
                    {4}<br>
                    {5}<br>
                    </p>
          </body>
        </html>
        """
        x = [deadline_date]

        insert_dfs(former_staff_dfs, x)

        new_html2 = html2.format(*x)

        # Try to send the email, otherwise add the recpient's email address to failed_recipients
        try:
            if any(x.empty is False for x in former_staff_dfs.values()):
                utils.send_mail(send_from=creds['admin_send_from'],
                                send_to=former_staff_recipients,  # rename numbered variables
                                cc=notification_cc,
                                subject=subject2,
                                html=new_html2,
                                username=creds['admin_username'],
                                password=creds['admin_password'],
                                is_tls=True,
                                wb=True,
                                file_path=file_path2,
                                filename=filename2)
        except smtplib.SMTPException:
            failed_recipients.append(
                ['RECIPIENT NAME',
                 former_staff_recipients,
                 'Illinois - University of Illinois Extension (Implementing Agency)'])

        # Email the Corrections Report

        report_subject = 'Monthly PEARS Corrections ' + prev_month.strftime('%b-%Y')

        html3 = """<html>
          <head></head>
        <body>
                    <p>
                    Hello everyone,<br><br>

                    The attached reported compiles the most recent round of monthly PEARS corrections.
                    If you have any questions, please reply to this email and a member of the FCS Evaluation Team will reach out soon.<br>

                    <br>Thanks and have a great day!<br>
                    <br> <b> FCS Evaluation Team </b> <br>
                    <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
                    </p>
          </body>
        </html>
        """

        # Try to send the email, otherwise print failure notication to console
        try:
            utils.send_mail(send_from=creds['admin_send_from'],
                            send_to=report_recipients,
                            cc=report_cc,
                            subject=report_subject,
                            html=html3,
                            username=creds['admin_username'],
                            password=creds['admin_password'],
                            is_tls=True,
                            wb=True,
                            file_path=file_path1,
                            filename=filename1)
        except smtplib.SMTPException:
            print("Failed to send Corrections Report.")

        # Notify admin of any failed attempts to send an email
        utils.send_failure_notice(failed_recipients=failed_recipients,
                                  send_from=creds['admin_send_from'],
                                  send_to=creds['admin_send_from'],
                                  username=creds['admin_username'],
                                  password=creds['admin_password'],
                                  fail_subject=report_subject + ' Failure Notice',
                                  success_msg='Data cleaning notifications sent successfully.')

# REFACTOR REPORT TO ENABLE AD HOC USAGE
# Run Monthly Data Cleaning from command line as ad hoc report
# Parse inputs with argparse
# if __name__ == '__main__':
#     main()
