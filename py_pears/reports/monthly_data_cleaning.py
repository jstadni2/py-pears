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
    Coalitions_Export = pd.ExcelFile(export_dir + '/' + "Coalition_Export.xlsx")
    Coa_Data = pd.read_excel(Coalitions_Export, 'Coalition Data')
    Coa_Data = utils.reformat(Coa_Data, custom_field_labels)
    # Only data clean records for SNAP-Ed
    # SNAP-Ed staff occasionally select the wrong program_area for Coalitions
    Coa_Data = Coa_Data.loc[(Coa_Data['program_area'] == 'SNAP-Ed') |
                            (Coa_Data['reported_by_email'].isin(snap_ed_staff['E-MAIL'])) |
                            (Coa_Data['reported_by_email'].isin(
                                former_snap_ed_staff['email']))]  # Filtering for former staff will include transfers
    Coa_Members = pd.read_excel(Coalitions_Export, 'Members')

    # Import list of Illinois names, used to flag Coalition Members with individual's names
    # Source: https://www.ssa.gov/oact/babynames/state/
    IL_names = pd.read_csv(names_list,
                           delimiter=",",
                           names=['state', 'sex', 'year', 'name', 'frequency'])
    IL_names = IL_names['name'].drop_duplicates()
    IL_names = IL_names.astype(str) + ' '

    # Import Indirect Activity data and Intervention Channels
    Indirect_Activities_Export = pd.ExcelFile(export_dir + "Indirect_Activity_Export.xlsx")
    IA_Data = pd.read_excel(Indirect_Activities_Export, 'Indirect Activity Data')
    IA_Data = utils.reformat(IA_Data, custom_field_labels)
    # Only data clean records for SNAP-Ed
    IA_Data = IA_Data.loc[IA_Data['program_area'] == 'SNAP-Ed']
    IA_IC = pd.read_excel(Indirect_Activities_Export, 'Intervention Channels')

    # Import Partnerships data
    Partnerships_Export = pd.ExcelFile(export_dir + "Partnership_Export.xlsx")
    Part_Data = pd.read_excel(Partnerships_Export, 'Partnership Data')
    Part_Data = utils.reformat(Part_Data, custom_field_labels)
    # Only data clean records for SNAP-Ed
    # SNAP-Ed staff occasionally select the wrong program_area for Partnerships
    Part_Data = Part_Data.loc[(Part_Data['program_area'] == 'SNAP-Ed') |
                              (Part_Data['reported_by_email'].isin(snap_ed_staff['E-MAIL'])) |
                              (Part_Data['reported_by_email'].isin(
                                  former_snap_ed_staff['email']))]  # Filtering for former staff will include transfers

    # Import Program Activity data and Sessions
    Program_Activities_Export = pd.ExcelFile(export_dir + "Program_Activities_Export.xlsx")
    PA_Data = pd.read_excel(Program_Activities_Export, 'Program Activity Data')
    PA_Data = utils.reformat(PA_Data, custom_field_labels)
    # Subset Program Activities for Family Consumer Science
    PA_Data_FCS = PA_Data.loc[PA_Data['program_areas'].str.contains('Family Consumer Science')]
    # Subset Program Activities for SNAP-Ed
    PA_Data = PA_Data.loc[PA_Data['program_areas'].str.contains('SNAP-Ed')]
    PA_Sessions = pd.read_excel(Program_Activities_Export, 'Sessions')

    # Import PSE Site Activity data, Needs, Readiness, Effectiveness, and Changes
    PSE_Site_Activities_Export = pd.ExcelFile(export_dir + "PSE_Site_Activity_Export.xlsx")
    PSE_Data = pd.read_excel(PSE_Site_Activities_Export, 'PSE Data')
    PSE_Data = utils.reformat(PSE_Data, custom_field_labels)
    PSE_NRE = pd.read_excel(PSE_Site_Activities_Export, 'Needs, Readiness, Effectiveness')
    PSE_Changes = pd.read_excel(PSE_Site_Activities_Export, 'Changes')

    # Import Update Notifications, used for the Corrections Report
    Update_Notes = pd.read_excel(update_notifications,
                                 sheet_name='Monthly Data Cleaning').drop(columns=['Tab'])

    # Monthly PEARS Data Cleaning

    # Timestamp and report year bounds used to filter data to clean
    ts = pd.to_datetime("today")
    report_year_start = '10/01/2021'
    report_year_end = '09/30/2022'

    # Coaltions

    # Convert counties to units for use in update notification email
    Coa_Data['coalition_unit'] = Coa_Data['coalition_unit'].str.replace('|'.join([' \(County\)', ' \(District\)', 'Unit ']),
                                                                        '', regex=True)
    Coa_Data = pd.merge(Coa_Data, unit_counties, how='left', left_on='coalition_unit', right_on='County')
    Coa_Data.loc[(~Coa_Data['coalition_unit'].isin(unit_counties['Unit #'])) &
                 (Coa_Data['coalition_unit'].isin(unit_counties['County'])), 'coalition_unit'] = Coa_Data['Unit #']

    # Filter out test records, select relevant columns
    Coa_Data = Coa_Data.loc[~Coa_Data['coalition_name'].str.contains('(?i)TEST', regex=True),
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

    Coa_Data['GENERAL INFORMATION TAB UPDATES'] = np.nan

    Coa_Data['GI UPDATE1'] = np.nan
    Coa_Data.loc[Coa_Data[
                     'action_plan_name'].isnull(), 'GI UPDATE1'] = 'Please select \'Health: Chronic Disease Prevention & Management\' for Action Plan Name.'

    Coa_Data['GI UPDATE2'] = np.nan
    Coa_Data.loc[Coa_Data['program_area'] != 'SNAP-Ed', 'GI UPDATE2'] = 'Please select \'SNAP-Ed\' for Program Area.'

    # Concatenate General Information tab updates
    Coa_Data['GENERAL INFORMATION TAB UPDATES'] = Coa_Data['GI UPDATE1'].fillna('') + '\n' + Coa_Data['GI UPDATE2'].fillna(
        '')
    Coa_Data.loc[Coa_Data['GENERAL INFORMATION TAB UPDATES'].str.isspace(), 'GENERAL INFORMATION TAB UPDATES'] = np.nan
    Coa_Data['GENERAL INFORMATION TAB UPDATES'] = Coa_Data['GENERAL INFORMATION TAB UPDATES'].str.strip()

    Coa_Data['CUSTOM DATA TAB UPDATES'] = np.nan
    Coa_Data.loc[Coa_Data[
                     'snap_ed_grant_goals'].isnull(), 'CUSTOM DATA TAB UPDATES'] = 'Please complete the Custom Data tab for this entry.'

    Coa_Data['COALITION MEMBERS TAB UPDATES'] = np.nan

    # Count Coalition Members of each Coalition, flag Coalitions that have none
    Coa_Data['CM UPDATE1'] = np.nan
    Coa_Members_Count = Coa_Members.groupby('coalition_id')['member_id'].count().reset_index(name='# of Members')
    Coa_Data = pd.merge(Coa_Data, Coa_Members_Count, how='left', on='coalition_id')
    Coa_Data.loc[(Coa_Data['# of Members'].isnull()) | (
            Coa_Data['# of Members'] == 0), 'CM UPDATE1'] = 'Please add organizational members to this coalition.'

    # Subsequent updates require Members data
    Coa_Members_Data = pd.merge(Coa_Data, Coa_Members, how='left', on='coalition_id').rename(
        columns={'name': 'member_name'})

    Coa_Members_Data['CM UPDATE2'] = np.nan
    Coa_Members_Data.loc[
        (Coa_Members_Data['type'] != 'Community members/individuals') & (Coa_Members_Data['site_id'].isnull()),
        'CM UPDATE2'] = 'Please add the Site to the Coalition Member(s) of this coalition.'

    # Flag Coalition Members that contain individuals' names
    Coa_Members_Data['CM UPDATE3'] = np.nan
    # Terms indicating false positives
    exclude_terms = ['University', 'Hospital', 'YMCA', 'Center', 'County', 'Elementary', 'Foundation', 'Church', 'Club',
                     'Daycare', 'Housing', 'SNAP-Ed']
    Coa_Members_Data.loc[(Coa_Members_Data['member_name'].str.contains('|'.join(IL_names), na=False)) &
                         (Coa_Members_Data['member_name'].str.count(' ') == 1) &
                         (~Coa_Members_Data['member_name'].str.contains('|'.join(exclude_terms), na=False)),
                         'CM UPDATE3'] = 'The Coalition Member Name cannot contain names of individuals. Please enter either the organization name or \'Community member\'.'

    # Concatenate Coalition Members tab updates
    Coa_Members_Data['COALITION MEMBERS TAB UPDATES'] = (Coa_Members_Data['CM UPDATE1'].fillna('') +
                                                         '\n' + Coa_Members_Data['CM UPDATE2'].fillna('') +
                                                         '\n' + Coa_Members_Data['CM UPDATE3'].fillna(''))
    Coa_Members_Data.loc[
        Coa_Members_Data['COALITION MEMBERS TAB UPDATES'].str.isspace(), 'COALITION MEMBERS TAB UPDATES'] = np.nan
    Coa_Members_Data['COALITION MEMBERS TAB UPDATES'] = Coa_Members_Data['COALITION MEMBERS TAB UPDATES'].str.strip()
    Coa_Members_Data['COALITION MEMBERS TAB UPDATES'] = Coa_Members_Data['COALITION MEMBERS TAB UPDATES'].str.replace(
        r'\n+', '', regex=True)

    # Subset records that require updates
    Coa_Corrections = Coa_Members_Data.loc[Coa_Members_Data.filter(like='UPDATE').notnull().any(1)]
    member_updates = ['CM UPDATE2', 'CM UPDATE3']
    Coa_Corrections = drop_child_dupes(Coa_Corrections, member_updates, 'coalition_id', 'member_id')
    # Coa_Corrections is exported in the Corrections Report
    Coa_Corrections_Cols = ['coalition_id',
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
                            'site_id']
    Coa_Corrections2 = Coa_Corrections[Coa_Corrections_Cols].set_index('coalition_id').rename(
        columns={'coalition_unit': 'unit'})
    Coa_Corrections2['site_id'] = pd.to_numeric(Coa_Corrections2['site_id'], downcast='integer')
    Coa_Corrections2 = Coa_Corrections2.fillna('')
    Coa_Corrections2['GENERAL INFORMATION TAB UPDATES'] = Coa_Corrections2['GENERAL INFORMATION TAB UPDATES'].str.replace(
        '\n', ' ', regex=True)
    Coa_Corrections2['COALITION MEMBERS TAB UPDATES'] = Coa_Corrections2['COALITION MEMBERS TAB UPDATES'].str.replace('\n',
                                                                                                                      ' ',
                                                                                                                      regex=True)
    # Coa_Corrections2 is used in the update notification email body

    # Indirect Activities

    # Set Indirect Activity data cleaning flags

    # Convert counties to units for use in update notification email
    IA_Data['unit'] = IA_Data['unit'].str.replace('|'.join([' \(County\)', ' \(District\)', 'Unit ']), '', regex=True)
    IA_Data = pd.merge(IA_Data, unit_counties, how='left', left_on='unit', right_on='County')
    IA_Data.loc[
        (~IA_Data['unit'].isin(unit_counties['Unit #'])) & (IA_Data['unit'].isin(unit_counties['County'])), 'unit'] = \
    IA_Data['Unit #']

    # Filter out test records, select relevant columns
    IA_Data = IA_Data.loc[~IA_Data['title'].str.contains('(?i)TEST', regex=True),
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

    IA_Data['CUSTOM DATA TAB UPDATES'] = np.nan

    IA_Data.loc[IA_Data.duplicated(subset=['reported_by_email', 'type'], keep=False),
                'CUSTOM DATA TAB UPDATES'] = 'Indirect Activity entries with the same Type must be combined into a single entry.'
    IA_Data.loc[IA_Data[
                    'snap_ed_grant_goals'].isnull(), 'CUSTOM DATA TAB UPDATES'] = 'Please complete the Custom Data tab for this entry.'

    # Filter out test records, select relevant columns
    IA_IC = IA_IC.loc[~IA_IC['activity'].str.contains('(?i)TEST', regex=True),
                      ['activity_id',
                       'activity',
                       'channel_id',
                       'channel',
                       'description',
                       'site_id',
                       'site_name',
                       'reach',
                       'newly_reached']]
    IA_IC['INTERVENTION CHANNELS AND REACH TAB UPDATES'] = np.nan

    # Subsequent updates require Intervention Channels data
    IA_IC_Data = pd.merge(IA_Data, IA_IC, how='left', on='activity_id')

    # Flag Intervention Channels that don't contain a date in their description
    IA_IC_Data['IC UPDATE1'] = np.nan
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    IA_IC_Data['description'] = IA_IC_Data['description'].astype(str)
    IA_IC_Data.loc[IA_IC_Data['description'] == 'nan', 'description'] = ''
    IA_IC_Data.loc[~IA_IC_Data['description'].str.contains('|'.join(months)) & ~IA_IC_Data['description'].str.contains(
        r'\d+/|-|.{1}\d{2,4}'),
                   'IC UPDATE1'] = 'Add the date this activity occurred to the Description.'

    IA_IC_Data['IC UPDATE2'] = np.nan
    # IA_IC_Data.loc[(IA_IC_Data['site_name'].isnull()) | (IA_IC_Data['site_name'] == 'abc placeholder'),
    #                'IC UPDATE2'] = 'Select a Site as directed by the cheat sheet.'
    # As of 3/18/22, Indirect Activity site is set to a required field in PEARS

    IA_IC_Data['IC UPDATE3'] = np.nan
    IA_IC_Data.loc[(IA_IC_Data['newly_reached'].notnull()) & (
            IA_IC_Data['newly_reached'] != 0), 'IC UPDATE3'] = 'Newly Reached number must be \'0\'.'

    IA_IC_Data['IC UPDATE4'] = np.nan
    IA_IC_Data.loc[(IA_IC_Data['reach'].isnull()) | (IA_IC_Data['reach'] == 0),
                   'IC UPDATE4'] = 'Please enter a non-zero value for \'Estimated # of unique individuals reached\'.'
    # How are null reach values possible?
    # Reach is null if there are no Intervention Channels for the Indirect Activity

    IA_IC_Data['IC UPDATE5'] = np.nan
    IA_IC_Data.loc[IA_IC_Data.duplicated(subset=['activity_id', 'description', 'site_id'], keep=False),
                   'IC UPDATE5'] = 'Please remove duplicate Intervention Channels or differentiate them by entering date and name in Description according to the cheat sheets.'

    IA_IC_Data['IC UPDATE6'] = np.nan
    IA_IC_Data.loc[IA_IC_Data[
                       'channel'] == "Hard copy materials (e.g. flyers, pamphlets, activity books, posters, banners, postcards, recipe cards, or newsletters for mailings)",
                   'IC UPDATE6'] = 'Please delete Intervention Channels for hard copy materials, which should not be reported for FY22.'

    # Concatenate Intervention Channels and Rearch tab updates
    IA_IC_Data['INTERVENTION CHANNELS AND REACH TAB UPDATES'] = (IA_IC_Data['IC UPDATE1'].fillna('') +
                                                                 '\n' + IA_IC_Data['IC UPDATE2'].fillna('') +
                                                                 '\n' + IA_IC_Data['IC UPDATE3'].fillna('') +
                                                                 '\n' + IA_IC_Data['IC UPDATE4'].fillna('') +
                                                                 '\n' + IA_IC_Data['IC UPDATE5'].fillna('') +
                                                                 '\n' + IA_IC_Data['IC UPDATE6'].fillna(''))
    IA_IC_Data.loc[IA_IC_Data[
                       'INTERVENTION CHANNELS AND REACH TAB UPDATES'].str.isspace(), 'INTERVENTION CHANNELS AND REACH TAB UPDATES'] = np.nan
    IA_IC_Data['INTERVENTION CHANNELS AND REACH TAB UPDATES'] = IA_IC_Data[
        'INTERVENTION CHANNELS AND REACH TAB UPDATES'].str.strip()
    IA_IC_Data['INTERVENTION CHANNELS AND REACH TAB UPDATES'] = IA_IC_Data[
        'INTERVENTION CHANNELS AND REACH TAB UPDATES'].str.replace(r'\n+', '\n', regex=True)

    # Subset records that require updates
    IA_Corrections = IA_IC_Data.loc[IA_IC_Data.filter(like='UPDATE').notnull().any(1)]
    IC_Updates = IA_IC_Data.columns[IA_IC_Data.columns.str.contains('IC')].tolist()
    IA_Corrections = drop_child_dupes(IA_Corrections, IC_Updates, 'activity_id', 'channel_id')
    # IA_Corrections is exported in the Corrections Report
    IA_Corrections_Cols = ['activity_id',
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
                           'reach']
    IA_Corrections2 = IA_Corrections[IA_Corrections_Cols].set_index('activity_id')
    IA_Corrections2['newly_reached'] = pd.to_numeric(IA_Corrections2['newly_reached'], downcast='integer')
    IA_Corrections2['newly_reached'] = pd.to_numeric(IA_Corrections2['channel_id'], downcast='integer')
    IA_Corrections2 = IA_Corrections2.fillna('')
    IA_Corrections2['INTERVENTION CHANNELS AND REACH TAB UPDATES'] = IA_Corrections2[
        'INTERVENTION CHANNELS AND REACH TAB UPDATES'].str.replace('\n', ' ')
    # IA_Corrections2 is used in the update notification email body

    # Partnerships

    # Convert counties to units for use in update notification email
    Part_Data['partnership_unit'] = Part_Data['partnership_unit'].str.replace(
        '|'.join([' \(County\)', ' \(District\)', 'Unit ']), '', regex=True)
    Part_Data = pd.merge(Part_Data, unit_counties, how='left', left_on='partnership_unit', right_on='County')
    Part_Data.loc[(~Part_Data['partnership_unit'].isin(unit_counties['Unit #'])) &
                  (Part_Data['partnership_unit'].isin(unit_counties['County'])), 'partnership_unit'] = Part_Data['Unit #']

    # Filter out test records, select relevant columns
    Part_Data = Part_Data.loc[~Part_Data['partnership_name'].str.contains('(?i)TEST', regex=True),
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

    Part_Data['GENERAL INFORMATION TAB UPDATES'] = np.nan

    Part_Data['GI UPDATE1'] = np.nan
    Part_Data.loc[Part_Data[
                      'action_plan_name'].isnull(), 'GI UPDATE1'] = 'Please enter the Action Plan Name for the current report year.'

    Part_Data['GI UPDATE2'] = np.nan
    Part_Data.loc[(Part_Data['is_direct_education_intervention'] == 0) & (Part_Data['is_pse_intervention'] == 0),
                  'GI UPDATE2'] = 'Please select \'Direct Education\' and/or \'Policy, Systems & Environmental Changes\' for this partner\'s intervention types.'

    Part_Data['GI UPDATE3'] = np.nan
    Part_Data.loc[Part_Data['program_area'] != 'SNAP-Ed', 'GI UPDATE3'] = 'Please select \'SNAP-Ed\' for Program Area.'

    # Concatenate General Information tab updates
    Part_Data['GENERAL INFORMATION TAB UPDATES'] = (Part_Data['GI UPDATE1'].fillna('') +
                                                    '\n' + Part_Data['GI UPDATE2'].fillna('') +
                                                    '\n' + Part_Data['GI UPDATE3'].fillna(''))
    Part_Data.loc[Part_Data['GENERAL INFORMATION TAB UPDATES'].str.isspace(), 'GENERAL INFORMATION TAB UPDATES'] = np.nan
    Part_Data['GENERAL INFORMATION TAB UPDATES'] = Part_Data['GENERAL INFORMATION TAB UPDATES'].str.strip()
    Part_Data['GENERAL INFORMATION TAB UPDATES'] = Part_Data['GENERAL INFORMATION TAB UPDATES'].str.replace(r'\n+', '\n',
                                                                                                            regex=True)

    Part_Data['CUSTOM DATA TAB UPDATES'] = np.nan
    Part_Data.loc[Part_Data[
                      'snap_ed_grant_goals'].isnull(), 'CUSTOM DATA TAB UPDATES'] = 'Please complete the Custom Data tab for this entry.'

    Part_Data['EVALUATION TAB UPDATES'] = np.nan
    Part_Data.loc[Part_Data['relationship_depth'].isnull(), 'EVALUATION TAB UPDATES'] = 'Enter the Relationship Depth.'

    # Subset records that require updates
    Part_Corrections = Part_Data.loc[Part_Data.filter(like='UPDATE').notnull().any(1)]
    # Part_Corrections is exported in the Corrections Report
    Part_Corrections_Cols = ['partnership_id',
                             'partnership_name',
                             'reported_by',
                             'reported_by_email',
                             'partnership_unit',
                             'GENERAL INFORMATION TAB UPDATES',
                             'action_plan_name',
                             'program_area',
                             'CUSTOM DATA TAB UPDATES',
                             'EVALUATION TAB UPDATES',
                             'relationship_depth']
    Part_Corrections2 = Part_Corrections[Part_Corrections_Cols].set_index('partnership_id').rename(
        columns={'partnership_unit': 'unit'}).fillna('')
    Part_Corrections2['GENERAL INFORMATION TAB UPDATES'] = Part_Corrections2['GENERAL INFORMATION TAB UPDATES'].str.replace(
        '\n', ' ')
    # Part_Corrections2 is used in the update notification email body

    # Program Activities

    # Set Program Activities data cleaning flags

    # Select relevant columns
    PA_Sessions = PA_Sessions.loc[:,
                  ['session_id', 'program_id', 'start_date', 'start_date_with_time', 'length', 'num_participants']]

    PA_Sessions['GENERAL INFORMATION TAB UPDATES'] = np.nan

    PA_Sessions['GI UPDATE1'] = np.nan
    PA_Sessions['start_date'] = pd.to_datetime(PA_Sessions['start_date'])
    PA_Sessions.loc[(PA_Sessions['start_date'] < report_year_start) | (PA_Sessions['start_date'] > report_year_end),
                    'GI UPDATE1'] = 'Please enter a session Start Date within 10/01/2021 to 09/30/2022.'

    PA_Sessions['GI UPDATE2'] = np.nan
    PA_Sessions.loc[(PA_Sessions['start_date_with_time'] < ts) & (
        PA_Sessions['num_participants'].isnull()), 'GI UPDATE2'] = 'Please enter the number of participants.'
    PA_Sessions.loc[(PA_Sessions['start_date_with_time'] < ts) & (
            PA_Sessions['num_participants'] == 0), 'GI UPDATE2'] = 'Please delete sessions with 0 participants.'

    PA_Sessions['GI UPDATE3'] = np.nan
    PA_Sessions.loc[PA_Sessions.duplicated(subset=['program_id', 'start_date_with_time'],
                                           keep=False), 'GI UPDATE3'] = 'Please remove duplicate Sessions.'

    # Convert counties to units for use in update notification email
    PA_Data['unit'] = PA_Data['unit'].str.replace('|'.join([' \(County\)', ' \(District\)', 'Unit ']), '', regex=True)
    PA_Data = pd.merge(PA_Data, unit_counties, how='left', left_on='unit', right_on='County')
    PA_Data.loc[
        (~PA_Data['unit'].isin(unit_counties['Unit #'])) & (PA_Data['unit'].isin(unit_counties['County'])), 'unit'] = \
    PA_Data['Unit #']

    # Filter out test records, select relevant columns
    PA_Data = PA_Data.loc[(~PA_Data['name'].str.contains('(?i)TEST', regex=True))
                          & (PA_Data['name'] != 'abc placeholder'),
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
    PA_Sessions_Data = pd.merge(PA_Data, PA_Sessions, how='left', on='program_id', suffixes=('_PA', '_Session'))

    PA_Sessions_Data['GI UPDATE4'] = np.nan
    PA_Sessions_Data.loc[(PA_Sessions_Data['length'] < 20) | (
        PA_Sessions_Data['length'].isnull()), 'GI UPDATE4'] = 'Please enter a Session Length of 20 mins or longer.'

    # Concatenate General Information tab updates
    PA_Sessions_Data['GENERAL INFORMATION TAB UPDATES'] = (PA_Sessions_Data['GI UPDATE1'].fillna('') +
                                                           '\n' + PA_Sessions_Data['GI UPDATE2'].fillna('') +
                                                           '\n' + PA_Sessions_Data['GI UPDATE3'].fillna('') +
                                                           '\n' + PA_Sessions_Data['GI UPDATE4'].fillna(''))
    PA_Sessions_Data.loc[
        PA_Sessions_Data['GENERAL INFORMATION TAB UPDATES'].str.isspace(), 'GENERAL INFORMATION TAB UPDATES'] = np.nan
    PA_Sessions_Data['GENERAL INFORMATION TAB UPDATES'] = PA_Sessions_Data['GENERAL INFORMATION TAB UPDATES'].str.strip()
    PA_Sessions_Data['GENERAL INFORMATION TAB UPDATES'] = PA_Sessions_Data['GENERAL INFORMATION TAB UPDATES'].str.replace(
        r'\n+', '\n', regex=True)

    PA_Sessions_Data['CUSTOM DATA TAB UPDATES'] = np.nan
    PA_Sessions_Data.loc[PA_Sessions_Data[
                             'snap_ed_grant_goals'].isnull(), 'CUSTOM DATA TAB UPDATES'] = 'Please complete the Custom Data tab for this entry.'
    PA_Sessions_Data.loc[(PA_Sessions_Data['snap_ed_special_projects'].str.contains('None'))
                         & (PA_Sessions_Data[
                                'snap_ed_special_projects'] != 'None'), 'CUSTOM DATA TAB UPDATES'] = 'Please remove \'None\' from Special Projects.'

    PA_Sessions_Data['SNAP-ED CUSTOM DATA TAB UPDATES'] = np.nan

    PA_Sessions_Data['SCD UPDATE1'] = np.nan
    PA_Sessions_Data.loc[PA_Sessions_Data[
                             'intervention'].isnull(), 'SCD UPDATE1'] = 'Please complete the SNAP-Ed Custom Data tab for this entry.'
    PA_Sessions_Data.loc[PA_Sessions_Data['intervention'] != 'SNAP-Ed Community Network',
                         'SCD UPDATE1'] = 'Please select \'SNAP-Ed Community Network\' for the Intervention Name.'

    PA_Sessions_Data['SCD UPDATE2'] = np.nan
    settings_other = ['Other places people', 'Other settings people']
    PA_Sessions_Data.loc[PA_Sessions_Data['setting'].str.contains('|'.join(settings_other),
                                                                  na=False), 'SCD UPDATE2'] = 'Please select a Setting besides \'Other\'.'

    PA_Sessions_Data['SNAP-ED CUSTOM DATA TAB UPDATES'] = PA_Sessions_Data['SCD UPDATE1'].fillna('') + '\n' + \
                                                          PA_Sessions_Data['SCD UPDATE2'].fillna('')
    PA_Sessions_Data.loc[
        PA_Sessions_Data['SNAP-ED CUSTOM DATA TAB UPDATES'].str.isspace(), 'SNAP-ED CUSTOM DATA TAB UPDATES'] = np.nan
    PA_Sessions_Data['SNAP-ED CUSTOM DATA TAB UPDATES'] = PA_Sessions_Data['SNAP-ED CUSTOM DATA TAB UPDATES'].str.strip()

    # Flag Program Activities where the unique participants is equal to the sum of session participants
    PA_Sessions_Data['DEMOGRAPHICS TAB UPDATES'] = np.nan
    PA_Sessions_Metrics = PA_Sessions.groupby('program_id').agg({'session_id': 'count',
                                                                 'num_participants': 'sum'}).reset_index().rename(
        columns={'session_id': '# of Sessions',
                 'num_participants': 'Total Session Participants'})
    PA_Sessions_Data = pd.merge(PA_Sessions_Data, PA_Sessions_Metrics, how='left', on='program_id')
    PA_Sessions_Data.loc[(PA_Sessions_Data['# of Sessions'] > 1) &
                         (PA_Sessions_Data['Total Session Participants'] == PA_Sessions_Data['participants_total']),
                         'DEMOGRAPHICS TAB UPDATES'] = 'Total number of unique participants should not equal total session participants. Please enter total unique participants according to the Cheat Sheet. If each session had brand new people, please enter these as individual program activity entries.'
    # End of year: For entries with only 1 session, the total # of session participants should = total # of unique participants.

    # Data clean FCS Program Activities
    PA_Data_FCS['CUSTOM DATA TAB UPDATES'] = np.nan
    PA_Data_FCS.loc[(PA_Data_FCS['fcs_program_team'].str.contains('SNAP-Ed'))
                    & (PA_Data_FCS[
                           'fcs_grant_goals'].isnull()), 'CUSTOM DATA TAB UPDATES'] = 'Please enter the IL SNAP-Ed Grant Goals for this entry.'

    # Append FCS Program Activities to SNAP-Ed Program Activities
    add_cols = PA_Sessions_Data.columns[~PA_Sessions_Data.columns.isin(PA_Data_FCS.columns)].tolist()
    PA_Data_FCS = pd.concat([PA_Data_FCS, pd.DataFrame(columns=add_cols)])  # turns program_id into float
    PA_Data_FCS['program_id'] = PA_Data_FCS['program_id'].astype(int)
    sub_cols = PA_Sessions_Data.columns[PA_Sessions_Data.columns.isin(PA_Data_FCS.columns)].tolist()
    PA_Sessions_Data = PA_Sessions_Data.append(PA_Data_FCS[sub_cols], ignore_index=True)
    # possible dupes added?

    # Subset records that require updates
    PA_Corrections = PA_Sessions_Data.loc[PA_Sessions_Data.filter(like='UPDATE').notnull().any(1)]
    session_updates = ['GI UPDATE1', 'GI UPDATE2', 'GI UPDATE4']
    PA_Corrections = drop_child_dupes(PA_Corrections, session_updates, 'program_id', 'session_id')
    # PA_Corrections is exported in the Corrections Report
    PA_Corrections_Cols = ['program_id',
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
                           'primary_curriculum']
    PA_Corrections2 = PA_Corrections[PA_Corrections_Cols].set_index('program_id')
    PA_Corrections2['num_participants'] = pd.to_numeric(PA_Corrections2['num_participants'], downcast='integer')
    PA_Corrections2 = PA_Corrections2.fillna('')
    PA_Corrections2['GENERAL INFORMATION TAB UPDATES'] = PA_Corrections2['GENERAL INFORMATION TAB UPDATES'].str.replace(
        '\n', ' ')
    PA_Corrections2['SNAP-ED CUSTOM DATA TAB UPDATES'] = PA_Corrections2['SNAP-ED CUSTOM DATA TAB UPDATES'].str.replace(
        '\n', ' ')
    PA_Corrections2['start_date_with_time'] = pd.to_datetime(PA_Corrections2['start_date_with_time']).dt.strftime(
        '%m-%d-%Y %r').fillna('')
    # PA_Corrections2 is used in the update notification email body

    # PSE Site Activities

    # Convert counties to units for use in update notification email
    PSE_Data['pse_unit'] = PSE_Data['pse_unit'].str.replace('|'.join([' \(County\)', ' \(District\)', 'Unit ']), '',
                                                            regex=True)
    PSE_Data = pd.merge(PSE_Data, unit_counties, how='left', left_on='pse_unit', right_on='County')
    PSE_Data.loc[(~PSE_Data['pse_unit'].isin(unit_counties['Unit #'])) & (
        PSE_Data['pse_unit'].isin(unit_counties['County'])), 'pse_unit'] = PSE_Data['Unit #']

    # Filter out test records, select relevant columns
    PSE_Data['name'] = PSE_Data['name'].astype(str)
    PSE_Data = PSE_Data.loc[
        (~PSE_Data['name'].str.contains('(?i)TEST', regex=True, na=False)) & (PSE_Data['site_name'] != 'abc placeholder'),
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

    PSE_Data['GENERAL INFORMATION TAB UPDATES'] = np.nan

    PSE_Data['GI UPDATE1'] = np.nan
    PSE_Data.loc[(PSE_Data['start_fiscal_year'] != 2022) & (
            PSE_Data['planning_stage_sites_contacted_and_agreed_to_participate'] == 1),
                 'GI UPDATE1'] = 'Only sites who agreed to partner for the first time in FY22 should have \'Sites contacted and agreed to participate\' selected.'

    PSE_Data['GI UPDATE2'] = np.nan
    PSE_Data.loc[PSE_Data['program_area'] == 'Family Consumer Science',
                 'GI UPDATE2'] = 'Please create a new PSE Site Activity for this record with Program Area set to "SNAP-Ed" and delete this entry.'

    PSE_Data['GI UPDATE3'] = np.nan
    PSE_Data.loc[(PSE_Data['site_id'].duplicated(keep=False))
                 & (PSE_Data[
                        'pse_unit'] != 'CPHP'), 'GI UPDATE3'] = 'Please remove duplicate PSE Site Activity entries for the same site.'

    PSE_Data['GI UPDATE4'] = np.nan
    PSE_Data.loc[PSE_Data[
                     'intervention'] != 'SNAP-Ed Community Network', 'GI UPDATE4'] = 'Please select \'SNAP-Ed Community Network\' for the Intervention Name.'

    # Concatenate General Information tab updates
    PSE_Data['GENERAL INFORMATION TAB UPDATES'] = (PSE_Data['GI UPDATE1'].fillna('') +
                                                   '\n' + PSE_Data['GI UPDATE2'].fillna('') +
                                                   '\n' + PSE_Data['GI UPDATE3'].fillna('') +
                                                   '\n' + PSE_Data['GI UPDATE4'].fillna(''))
    PSE_Data.loc[PSE_Data['GENERAL INFORMATION TAB UPDATES'].str.isspace(), 'GENERAL INFORMATION TAB UPDATES'] = np.nan
    PSE_Data['GENERAL INFORMATION TAB UPDATES'] = PSE_Data['GENERAL INFORMATION TAB UPDATES'].str.strip()
    PSE_Data['GENERAL INFORMATION TAB UPDATES'] = PSE_Data['GENERAL INFORMATION TAB UPDATES'].str.replace(r'\n+', '\n',
                                                                                                          regex=True)

    PSE_Data['CUSTOM DATA TAB UPDATES'] = np.nan
    PSE_Data.loc[PSE_Data[
                     'snap_ed_grant_goals'].isnull(), 'CUSTOM DATA TAB UPDATES'] = 'Please complete the Custom Data tab for this entry.'

    # Select relevant Needs, Readiness, Effectiveness columns
    PSE_NRE = PSE_NRE.loc[:,
              ['pse_id', 'assessment_id', 'assessment_type', 'assessment_tool', 'baseline_score', 'baseline_date',
               'follow_up_date', 'follow_up_score']]

    # Subsequent updates require Needs, Readiness, Effectiveness data
    PSE_NRE_Data = pd.merge(PSE_Data, PSE_NRE, how='left', on='pse_id')

    PSE_NRE_Data['NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'] = np.nan

    PSE_NRE_Data['NRE UPDATE1'] = np.nan
    PSE_NRE_Data.loc[(PSE_NRE_Data['assessment_type'] == 'Needs assessment/environmental scan') &
                     (PSE_NRE_Data['baseline_score'].isnull()) &
                     (~PSE_NRE_Data['assessment_tool'].str.contains('SLAQ',
                                                                    na=False)), 'NRE UPDATE1'] = 'Please enter the Assessment Score according to the Cheat Sheet.'

    PSE_NRE_Data['NRE UPDATE2'] = np.nan
    PSE_NRE_Data['baseline_date'] = pd.to_datetime(PSE_NRE_Data['baseline_date'])
    PSE_NRE_Data.loc[(PSE_NRE_Data['assessment_type'] == 'Needs assessment/environmental scan') &
                     (PSE_NRE_Data[
                          'baseline_date'].isnull()), 'NRE UPDATE2'] = 'Please enter a Baseline Date even if pre-assessment was conducted in a previous reporting year.'

    PSE_NRE_Data['NRE UPDATE3'] = np.nan
    PSE_NRE_Data['follow_up_date'] = pd.to_datetime(PSE_NRE_Data['follow_up_date'])
    PSE_NRE_Data.loc[(PSE_NRE_Data['assessment_type'] == 'Needs assessment/environmental scan') &
                     (PSE_NRE_Data['follow_up_date'].notnull()) &
                     (PSE_NRE_Data[
                          'follow_up_score'].isnull()), 'NRE UPDATE3'] = 'Please enter the Follow Up Assessment Score according to the Cheat Sheet.'

    PSE_NRE_Data['NRE UPDATE4'] = np.nan
    PSE_NRE_Data.loc[(PSE_NRE_Data['assessment_type'] == 'Needs assessment/environmental scan') &
                     (PSE_NRE_Data['follow_up_date'].isnull()) &
                     (PSE_NRE_Data[
                          'follow_up_score'].notnull()), 'NRE UPDATE4'] = 'Please enter the Follow Up Assessment Date.'

    # Concatenate Needs, Readiness, Effectiveness tab updates
    PSE_NRE_Data['NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'] = (PSE_NRE_Data['NRE UPDATE1'].fillna('') +
                                                                    '\n' + PSE_NRE_Data['NRE UPDATE2'].fillna('') +
                                                                    '\n' + PSE_NRE_Data['NRE UPDATE3'].fillna('') +
                                                                    '\n' + PSE_NRE_Data['NRE UPDATE4'].fillna(''))
    PSE_NRE_Data.loc[PSE_NRE_Data[
                         'NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'].str.isspace(), 'NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'] = np.nan
    PSE_NRE_Data['NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'] = PSE_NRE_Data[
        'NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'].str.strip()
    PSE_NRE_Data['NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'] = PSE_NRE_Data[
        'NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'].str.replace(r'\n+', '\n', regex=True)

    # Subsequent updates require Changes Adopted data
    PSE_NRE_Changes_Data = pd.merge(PSE_NRE_Data, PSE_Changes[['pse_id', 'change_id']], how='left',
                                    on='pse_id').drop_duplicates(subset=['pse_id', 'assessment_id'])

    PSE_NRE_Changes_Data['CHANGES ADOPTED TAB UPDATES'] = np.nan
    PSE_NRE_Changes_Data.loc[(PSE_NRE_Changes_Data['change_id'].notnull()) & (PSE_NRE_Changes_Data['total_reach'].isnull()),
                             'CHANGES ADOPTED TAB UPDATES'] = 'Please enter Total Reach according to the Cheat Sheet.'

    # Subset records that require updates
    PSE_Corrections = PSE_NRE_Changes_Data.loc[PSE_NRE_Changes_Data.filter(like='UPDATE').notnull().any(1)]
    # PSE_Corrections is exported in the Corrections Report
    PSE_Corrections_Cols = ['pse_id',
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
                            'total_reach']
    PSE_Corrections2 = PSE_Corrections[PSE_Corrections_Cols].set_index('pse_id').rename(
        columns={'pse_unit': 'unit'}).fillna('')
    PSE_Corrections2['GENERAL INFORMATION TAB UPDATES'] = PSE_Corrections2['GENERAL INFORMATION TAB UPDATES'].str.replace(
        '\n', ' ')
    PSE_Corrections2['NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'] = PSE_Corrections2[
        'NEEDS, READINESS & EFFECTIVENESS TAB UPDATES'].str.replace('\n', ' ')
    # When run from local Python env, these dates seem to be converted into objects when new dfs are made
    PSE_Corrections2['baseline_date'] = pd.to_datetime(PSE_Corrections2['baseline_date']).dt.strftime('%m-%d-%Y').fillna('')
    PSE_Corrections2['follow_up_date'] = pd.to_datetime(PSE_Corrections2['follow_up_date']).dt.strftime('%m-%d-%Y').fillna(
        '')


    # PSE_Corrections2 is used in the update notification email body

    # Corrections Report

    # Summarize and concatenate module corrections
    Coa_Sum = corrections_sum(Coa_Corrections, 'Coalitions')
    IA_Sum = corrections_sum(IA_Corrections, 'Indirect Activities')
    Part_Sum = corrections_sum(Part_Corrections, 'Partnerships')
    PA_Sum = corrections_sum(PA_Corrections, 'Program Activities')
    PSE_Sum = corrections_sum(PSE_Corrections, 'PSE Site Activities')
    Module_Sums = [Coa_Sum, IA_Sum, Part_Sum, PA_Sum, PSE_Sum]

    Corrections_Sum = pd.concat(Module_Sums, ignore_index=True)
    Corrections_Sum.insert(0, 'Module', Corrections_Sum.pop('Module'))
    Corrections_Sum = pd.merge(Corrections_Sum, Update_Notes, how='left', on=['Module', 'Update'])

    # Calculate the month for this report
    prev_month = (ts - pd.DateOffset(months=1)).to_period('M')

    # Export the Corrections Report as an Excel file

    filename1 = 'Monthly PEARS Corrections ' + prev_month.strftime('%Y-%m') + '.xlsx'
    file_path1 = output_dir + '/' + filename1

    dfs = {'Corrections Summary': Corrections_Sum,
           'Coalitions': Coa_Corrections,
           'Indirect Activities': IA_Corrections,
           'Partnerships': Part_Corrections,
           'Program Activities': PA_Corrections,
           'PSE': PSE_Corrections}

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
        Module_Corrections2 = [Coa_Corrections2, IA_Corrections2, Part_Corrections2, PA_Corrections2, PSE_Corrections2]
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

            Coa_df = staff_corrections(Coa_Corrections2, former=False, staff_email=x[1])
            IA_df = staff_corrections(IA_Corrections2, former=False, staff_email=x[1])
            Part_df = staff_corrections(Part_Corrections2, former=False, staff_email=x[1])
            PA_df = staff_corrections(PA_Corrections2, former=False, staff_email=x[1])
            PSE_df = staff_corrections(PSE_Corrections2, former=False, staff_email=x[1])

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

        Coa_df = staff_corrections(Coa_Corrections2)
        IA_df = staff_corrections(IA_Corrections2)
        Part_df = staff_corrections(Part_Corrections2)
        PA_df = staff_corrections(PA_Corrections2)
        PSE_df = staff_corrections(PSE_Corrections2)

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
