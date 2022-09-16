import pandas as pd
from datetime import datetime
from functools import reduce
import py_pears.utils as utils


# Assign and explode records by quarter
# df: dataframe of module export sheet records
# qtr_bounds: list of date strings for the lower/upper bounds of each quarter
# date_field: df column name of the date field to base quarters on (default: 'created')
def explode_quarters(df, qtr_bounds, date_field='created'):
    df[date_field] = pd.to_datetime(df[date_field])  # SettingWithCopyWarning
    df['report_quarter'] = ''  # SettingWithCopyWarning
    dt_list = [datetime.strptime(s, "%m/%d/%Y") for s in qtr_bounds]
    # Upper bound is exclusive
    df.loc[(df[date_field] >= dt_list[0]) & (df[date_field] < dt_list[1]), 'report_quarter'] = '1, 2, 3, 4'
    df.loc[(df[date_field] >= dt_list[1]) & (df[date_field] < dt_list[2]), 'report_quarter'] = '2, 3, 4'
    df.loc[(df[date_field] >= dt_list[2]) & (df[date_field] < dt_list[3]), 'report_quarter'] = '3, 4'
    df.loc[(df[date_field] >= dt_list[3]) & (df[date_field] < dt_list[4]), 'report_quarter'] = '4'
    df['report_quarter'] = df['report_quarter'].str.split(', ').tolist()  # SettingWithCopyWarning
    df = df.explode('report_quarter')
    df['report_quarter'] = pd.to_numeric(df['report_quarter'])
    return df


# Explode snap_ed_grant_goals field
# df: dataframe of module export sheet records
def explode_goals(df):
    df['goal'] = df['snap_ed_grant_goals'].str.split(',').tolist()
    df = df.explode('goal')
    return df


# Function to calculate the quarterly value of a given field
# df: dataframe of PEARS module data with 'report_quarter' column
# field: column used to calculate the quarterly value
# metric: 'sum' or 'count'
# label: string for the column label of the quarterly value
# goals: boolean for whether the metric should be grouped by 'goals' (default: False)
def quarterly_value(df, field, metric, label, goals=False):
    if goals:
        return df.groupby([
            'report_quarter', 'goal'])[field].agg(metric).reset_index(name=label).rename(columns={'goal': 'Goal'})
    else:
        return df.groupby('report_quarter')[field].agg(metric).reset_index(name=label)


# Class that bundles the input arguments of quarterly_value()
class QuarterlyValueInputs:
    def __init__(self, df, field, metric, label, goals=False):
        self.df = df
        self.field = field
        self.metric = metric
        self.label = label
        self.goals = goals


# Function to assign a percent column to a dataframe
# df: dataframe used to calculate the percent column
# num: column label to use as the percent numerator
# denom: column label to use as the percent denominator
# label: label for the resulting percent column
def percent(df, num, denom, label):
    df_copy = df.copy()
    df_copy[label] = 100 * df_copy[num] / df_copy[denom]
    return df_copy


# Function to filter a list of dataframes up to the specified fiscal quarter
# dfs: list of dataframes to be filters
# fq_int: integer value for the fiscal quarter
def filter_fq(dfs, fq):
    filtered_dfs = []
    for df in dfs:
        filtered_dfs.append(
            df.loc[df['report_quarter'] <= fq].rename(columns={'report_quarter': 'Report Quarter (YTD)'}))
    return filtered_dfs


# Run the Sites Report
# creds: dict of credentials loaded from credentials.json
# export_dir: directory where PEARS exports are downloaded to
# output_dir: directory where report outputs are saved
def main(creds, export_dir, output_dir):
    # Download required PEARS exports from S3
    utils.download_s3_exports(profile=creds['aws_profile'],
                              org=creds['s3_organization'],
                              dst=export_dir,
                              modules=['Program_Activities',
                                       'Indirect_Activity',
                                       'Coalition',
                                       'Partnership',
                                       'PSE_Site_Activity'])

    # Custom fields that require reformatting
    # Only needed for multi-select dropdowns
    custom_field_labels = ['fcs_program_team', 'snap_ed_grant_goals', 'fcs_grant_goals', 'fcs_special_projects',
                           'snap_ed_special_projects']

    # Import Indirect Activity data and Intervention Channels
    indirect_activities_export = pd.ExcelFile(export_dir + "Indirect_Activity_Export.xlsx")
    ia_export = pd.read_excel(indirect_activities_export, 'Indirect Activity Data')
    # Only report on records for SNAP-Ed
    ia_data = ia_export.loc[
        (ia_export['program_area'] == 'SNAP-Ed') & (~ia_export['title'].str.contains('(?i)TEST', regex=True))]
    ia_ic_export = pd.read_excel(indirect_activities_export, 'Intervention Channels')

    # Import Coalitions data and Coalition Members
    coalitions_export = pd.ExcelFile(export_dir + "Coalition_Export.xlsx")
    coa_export = pd.read_excel(coalitions_export, 'Coalition Data')
    # Only report on records for SNAP-Ed
    coa_data = coa_export.loc[
        (coa_export['program_area'] == 'SNAP-Ed') & (
            ~coa_export['coalition_name'].str.contains('(?i)TEST', regex=True))]
    coa_members_export = pd.read_excel(coalitions_export, 'Members')

    # Import Program Activity data and Sessions
    program_activities_export = pd.ExcelFile(export_dir + "program_activities_export.xlsx")
    pa_export = pd.read_excel(program_activities_export, 'Program Activity Data')
    # PA is only module to have cross-program_area collaboration
    pa_data = pa_export.loc[
        (pa_export['program_areas'].str.contains('SNAP-Ed')) & (
            ~pa_export['name'].str.contains('(?i)TEST', regex=True))]
    pa_sessions_export = pd.read_excel(program_activities_export, 'Sessions')

    # Import Partnerships data
    partnerships_export = pd.ExcelFile(export_dir + "Partnership_Export.xlsx")
    part_export = pd.read_excel(partnerships_export, 'Partnership Data')
    # Only report on records for SNAP-Ed
    part_data = part_export.loc[(part_export['program_area'] == 'SNAP-Ed') & (
        ~part_export['partnership_name'].str.contains('(?i)TEST', regex=True))]

    # Import PSE Site Activity data, Needs, Readiness, Effectiveness, and Changes
    pse_site_activities_export = pd.ExcelFile(export_dir + "PSE_Site_Activity_Export.xlsx")
    pse_export = pd.read_excel(pse_site_activities_export, 'PSE Data')
    pse_data = pse_export.loc[~pse_export['name'].str.contains('(?i)TEST', regex=True, na=False)]
    pse_changes_export = pd.read_excel(pse_site_activities_export, 'Changes')
    pse_nre_export = pd.read_excel(pse_site_activities_export, 'Needs, Readiness, Effectiveness')

    # Assign Quarters

    fy_22_qtr_bounds = ['10/01/2021', '01/11/2022', '04/11/2022', '07/11/2022', '10/18/2022']

    # Prep Coalitions data
    coa_data = utils.reformat(coa_data, custom_field_labels)
    coa_data = explode_quarters(coa_data, fy_22_qtr_bounds)
    coa_members_data = pd.merge(coa_members_export, coa_data[['coalition_id', 'program_area', 'report_quarter']],
                                how='left', on='coalition_id')
    coa_members_data = coa_members_data.loc[coa_members_data['program_area'] == 'SNAP-Ed']

    # Prep Indirect Activities data
    ia_data = utils.reformat(ia_data, custom_field_labels)
    ia_data = explode_quarters(ia_data, fy_22_qtr_bounds)
    ia_ic_data = pd.merge(ia_ic_export, ia_data[['activity_id', 'program_area', 'report_quarter']], how='left',
                          on='activity_id')
    ia_ic_data = ia_ic_data.loc[ia_ic_data['program_area'] == 'SNAP-Ed']
    # Use IC created field when export updated

    # Prep Program Activities data
    pa_data = utils.reformat(pa_data, custom_field_labels)
    pa_data = explode_quarters(pa_data, fy_22_qtr_bounds)
    pa_sessions_data = pd.merge(pa_sessions_export, pa_data[['program_id', 'program_areas']].drop_duplicates(),
                                how='left',
                                on='program_id')
    pa_sessions_data = pa_sessions_data.loc[pa_sessions_data['program_areas'].str.contains('SNAP-Ed', na=False)]
    pa_sessions_data = explode_quarters(pa_sessions_data, fy_22_qtr_bounds, date_field='start_date')
    pa_sessions_data = pa_sessions_data.loc[pa_sessions_data['report_quarter'] != '']
    # EARS – Program Activity Sessions:
    # Only program activities that have either more than one session or one
    # session greater than or equal to 20 minutes in length are counted.

    # Prep Partnerships data
    part_data = utils.reformat(part_data, custom_field_labels)
    part_data = explode_quarters(part_data, fy_22_qtr_bounds)

    # Prep PSE Site Activities data
    pse_data = utils.reformat(pse_data, custom_field_labels)
    pse_data = explode_quarters(pse_data, fy_22_qtr_bounds)
    pse_nre_data = explode_quarters(pse_nre_export, fy_22_qtr_bounds, date_field='baseline_date')
    pse_nre_data = pse_nre_data.loc[pse_nre_data['baseline_date'] >= fy_22_qtr_bounds[0]]

    # Calculate DHS report metrics

    # # of unique programming sites (direct ed & PSE)

    unique_sites = pa_data[['report_quarter', 'snap_ed_grant_goals', 'site_id']].append(
        pse_data[['report_quarter', 'snap_ed_grant_goals', 'site_id']], ignore_index=True).drop_duplicates()
    unique_sites = explode_goals(unique_sites)
    unique_sites = quarterly_value(df=unique_sites,
                                   field='site_id',
                                   metric='count',
                                   label='# of unique programming sites (direct ed & PSE)',
                                   goals=True)
    # Remove (direct ed & PSE) from column, add to snap_ed_grant_goals?

    unique_coalitions = quarterly_value(coa_data[['report_quarter', 'coalition_id']].drop_duplicates(),
                                        'coalition_id',
                                        'count',
                                        '# of unique programming sites (direct ed & PSE)')
    unique_coalitions[
        'Goal'] = 'Create community collaborations (reported as # coalitions and # organizational members)'
    unique_sites = unique_sites.append(unique_coalitions, ignore_index=True)

    # Total Unique Reach
    # Create package function for total_unique_reach(), pending FY23 guidance

    pa_sites_reach = pa_data[['report_quarter', 'snap_ed_grant_goals', 'site_id', 'participants_total']]
    pa_sites_reach = explode_goals(pa_sites_reach)
    pa_sites_reach = pa_sites_reach.rename(columns={'participants_total': 'PA_participants_total'}).groupby(
        ['report_quarter', 'goal', 'site_id'])['PA_participants_total'].agg('sum').reset_index(
        name='PA_participants_sum')
    pse_sites_reach = explode_goals(pse_data.loc[
                                        pse_data['total_reach'].notnull(), ['report_quarter', 'snap_ed_grant_goals',
                                                                            'site_id', 'total_reach']])
    pse_sites_reach = pse_sites_reach.sort_values(['report_quarter', 'site_id', 'total_reach']).rename(
        columns={'total_reach': 'PSE_total_reach'}).drop_duplicates(subset=['report_quarter', 'site_id'], keep='last')
    site_reach = pd.merge(pa_sites_reach, pse_sites_reach, how='outer', on=['report_quarter', 'goal', 'site_id'])
    site_reach['Site Reach'] = site_reach[['PA_participants_sum', 'PSE_total_reach']].max(axis=1)
    reach = quarterly_value(df=site_reach,
                            field='Site Reach',
                            metric='sum',
                            label='Total Reach',
                            goals=True)

    # coa_reach = coa_data[
    #     ['report_quarter',
    #      'number_of_members']].groupby('report_quarter')['number_of_members'].agg('sum').reset_index(name='Total Reach')
    # Reach via coa_data['number_of_members'] != reach via coa_members_data['member_id']
    coa_reach = quarterly_value(df=coa_members_data[['report_quarter', 'member_id']],
                                field='member_id',
                                metric='count',
                                label='Total Reach')
    coa_reach['Goal'] = 'Create community collaborations (reported as # coalitions and # organizational members)'
    reach = reach.append(coa_reach, ignore_index=True)

    goals_sites_reach = pd.merge(unique_sites, reach, how='outer', on=['Goal', 'report_quarter'])

    # DE participants reached

    pa_demo_dfs = []
    demo_subsets = {'participants_total': 'Total',
                    'participants_race_amerind': 'American Indian or Alaska Native',
                    'participants_race_asian': 'Asian',
                    'participants_race_black': 'Black or African American',
                    'participants_race_hawpac': 'Native Hawaiian/Other Pacific Islander',
                    'participants_race_white': 'White',
                    'participants_ethnicity_hispanic': 'Hispanic/Latinx',
                    'participants_ethnicity_non_hispanic': 'Non-Hispanic/Non-Latinx'
                    }

    for demo_field, demo_label in demo_subsets.items():
        pa_demo_dfs.append(quarterly_value(df=pa_data, field=demo_field, metric='sum', label=demo_label))

    # Create function for merging lists of dfs on report_quarter
    pa_demo = reduce(lambda left, right: pd.merge(left, right, how='outer', on='report_quarter'), pa_demo_dfs)

    for demo_field in demo_subsets.values():
        if demo_field == 'Total':
            continue
        pa_demo = percent(pa_demo, num=demo_field, denom='Total', label='% ' + demo_field)

    pa_demo = pa_demo.round(0).drop(columns=['Total'])
    # Pivot demo columns into values of Demographic Group column
    pa_demo = pa_demo.set_index('report_quarter').stack().reset_index().rename(columns={'level_1': 'Demographic Group'})
    pa_demo_a = pa_demo.loc[~pa_demo['Demographic Group'].str.contains('%')]
    pa_demo_b = pa_demo.loc[pa_demo['Demographic Group'].str.contains('%')]
    pa_demo_b['Demographic Group'] = pa_demo_b['Demographic Group'].str.replace('% ', '')

    pa_demo = pd.merge(pa_demo_a, pa_demo_b, how='left', on=['report_quarter', 'Demographic Group']).rename(
        columns={'0_x': 'Total (YTD)', '0_y': '%'})

    # RE-AIM Measures of Success

    # Reach

    reach_inputs = [
        QuarterlyValueInputs(
            df=pa_sites_reach,
            field='PA_participants_sum',
            metric='sum',
            label='# of unique participants attending direct education'),
        QuarterlyValueInputs(
            df=pa_sessions_data,
            field='num_participants',
            metric='sum',
            label='# of educational contacts via direct education'),
        QuarterlyValueInputs(
            df=pa_sessions_data,
            field='report_quarter',
            metric='count',
            label='# of lessons attended'),
        QuarterlyValueInputs(
            df=ia_ic_data,
            field='reach',
            metric='sum',
            label='# of indirect education contacts'),
        QuarterlyValueInputs(
            df=pse_sites_reach,
            field='PSE_total_reach',
            metric='sum',
            label='PSE total estimated reach')
    ]

    reach_dfs = []

    for inputs in reach_inputs:
        reach_dfs.append(quarterly_value(inputs.df, inputs.field, inputs.metric, inputs.label))

    re_aim_reach = reduce(lambda left, right: pd.merge(left, right, how='outer', on='report_quarter'), reach_dfs)

    # Adoption

    part_orgs_reach = quarterly_value(
        df=part_data,
        field='partnership_id',
        metric='count',
        label='# of partnering organizations reached'
    )

    programming_zips_cols = ['report_quarter', 'site_zip']
    programming_zips = pd.concat(
        [coa_members_data.loc[coa_members_data['site_zip'].notnull(), programming_zips_cols],
         ia_ic_data.loc[ia_ic_data['site_zip'].notnull(), programming_zips_cols],
         pa_data[programming_zips_cols],
         part_data[programming_zips_cols],
         pse_data[programming_zips_cols]
         ], ignore_index=True).drop_duplicates()

    adoption_zips = quarterly_value(
        df=programming_zips,
        field='site_zip',
        metric='count',
        label='# of unique zip codes reached'
    )

    re_aim_adoption = pd.merge(part_orgs_reach, adoption_zips, how='outer', on='report_quarter')

    # Implementation

    pse_changes = pd.merge(pse_changes_export, pse_data[['pse_id', 'report_quarter']], how='left', on='pse_id')

    pse_change_sites = quarterly_value(
        df=pse_changes.drop_duplicates(subset=['report_quarter', 'site_id']),
        field='site_id',
        metric='count',
        label='# of sites implementing a PSE change'
    )

    pse_nre_sites = quarterly_value(
        df=pse_nre_data.drop_duplicates(subset=['report_quarter', 'site_id']),
        field='site_id',
        metric='count',
        label='# of unique sites where an organizational readiness or environmental assessment was conducted'
    )

    re_aim_implementation = pd.merge(pse_change_sites, pse_nre_sites, how='outer', on='report_quarter')

    # Count of unique counties/cities with at least 1 partnership entry

    part_cities = quarterly_value(
        df=part_data.drop_duplicates(subset=['report_quarter', 'site_city']),
        field='site_city',
        metric='count',
        label='# of unique cities with at least 1 partnership entry'
    )  # Add this metric to report output

    part_counties = quarterly_value(
        df=part_data.drop_duplicates(subset=['report_quarter', 'site_county']),
        field='site_county',
        metric='count',
        label='# of unique counties with at least 1 partnership entry'
    )  # Add this metric to report output

    # Count and percent of PSE site activities will have a change adopted
    # related to food access, diet quality, or physical activity

    pse_changes_count = quarterly_value(
        df=pse_changes.drop_duplicates(subset=['report_quarter', 'pse_id']),
        field='pse_id',
        metric='count',
        label='# of PSE sites with a plan to implement PSE changes had at least one change adopted'
    )

    pse_count = quarterly_value(
        df=pse_data,
        field='pse_id',
        metric='count',
        label='# of PSE sites with a plan to implement PSE changes'
    )

    pse_changes_count = pd.merge(pse_changes_count, pse_count, how='left', on='report_quarter')

    # Add this metric to report output
    pse_changes_count = percent(df=pse_changes_count,
                                num='# of PSE sites with a plan to implement PSE changes had at least one change adopted',
                                denom='# of PSE sites with a plan to implement PSE changes',
                                label='% of PSE sites with a plan to implement PSE changes had at least one change adopted')

    # Final Report

    report_dfs = [goals_sites_reach, pa_demo, re_aim_reach, re_aim_adoption, re_aim_implementation]

    # Check if previous month is the last in the quarter
    prev_month = (pd.to_datetime("today") - pd.DateOffset(months=1)).strftime('%m')
    fq_lookup = pd.DataFrame({'fq': [1, 2, 3, 4], 'month': ['12', '03', '06', '09']})
    if prev_month in fq_lookup['month']:
        current_fq = fq_lookup.loc[fq_lookup['month'] == prev_month, 'fq'].item()
    # If not, report all four quarters
    else:
        current_fq = 4

    utils.write_report(file=output_dir + 'DHS Report FY2022 Q' + str(current_fq) + '.xlsx',
                       sheet_names=['Unique Sites and Reach by Goal',
                                    'Direct Education Demographics',
                                    'RE-AIM Reach',
                                    'RE-AIM Adoption',
                                    'RE-AIM Implementation'],
                       dfs=filter_fq(report_dfs, current_fq))
