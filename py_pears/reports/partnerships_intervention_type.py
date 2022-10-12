import pandas as pd
import py_pears.utils as utils


# UPDATE Extraneous report, refactor into notification flag in Monthly Data Cleaning

# Run the Partnerships Intervention Type report
# creds: dict of credentials loaded from org_settings.json
# export_dir: directory where PEARS exports are downloaded to
# output_dir: directory where report outputs are saved
# staff_list: path to the staff list Excel workbook
def main(creds, export_dir, output_dir, staff_list):
    # Download required PEARS exports from S3
    utils.download_s3_exports(profile=creds['aws_profile'],
                              org=creds['s3_organization'],
                              modules=['User',
                                       'Site',
                                       'Program_Activities',
                                       'Indirect_Activity',
                                       'Partnership',
                                       'PSE_Site_Activity'])

    fy22_inep_staff = pd.ExcelFile(staff_list)
    # Adjust header argument in following lines for actual staff list
    snap_ed_staff = pd.read_excel(fy22_inep_staff, sheet_name='SNAP-Ed Staff List', header=1)
    # Import list of former staff
    former_snap_ed_staff = pd.read_excel(fy22_inep_staff, sheet_name='Former Staff')
    former_snap_ed_staff['email'] = former_snap_ed_staff['E-MAIL/NETID'].map(str) + '@illinois.edu'

    # Import Indirect Activity data and Intervention Channels
    indirect_activities_export = pd.ExcelFile(export_dir + "Indirect_Activity_Export.xlsx")
    ia_data = pd.read_excel(indirect_activities_export, 'Indirect Activity Data')
    # Only data clean records for SNAP-Ed
    ia_data = ia_data.loc[ia_data['program_area'] == 'SNAP-Ed']
    ia_ic = pd.read_excel(indirect_activities_export, 'Intervention Channels')
    ia_ic_data = pd.merge(ia_data, ia_ic, how='left', on='activity_id')[['activity_id', 'site_id']]

    # Import Partnerships data
    partnerships_export = pd.ExcelFile(export_dir + "Partnership_Export.xlsx")
    part_data = pd.read_excel(partnerships_export, 'Partnership Data')
    # Only data clean records for SNAP-Ed
    # SNAP-Ed staff occasionally select the wrong program_area for Partnerships
    part_data = part_data.loc[(part_data['program_area'] == 'SNAP-Ed') |
                              (part_data['reported_by_email'].isin(snap_ed_staff['E-MAIL'])) |
                              (part_data['reported_by_email'].isin(former_snap_ed_staff['email'])),
                              ['partnership_id',
                               'partnership_name',
                               'is_direct_education_intervention',
                               'is_pse_intervention',
                               'site_id']]
    # Filtering for former staff will include transfers

    # Import Program Activity data
    program_activities_export = pd.ExcelFile(export_dir + "program_activities_export.xlsx")
    pa_data = pd.read_excel(program_activities_export, 'Program Activity Data')
    # Subset Program Activities for Family Consumer Science
    pa_data_fcs = pa_data.loc[pa_data['program_areas'].str.contains('Family Consumer Science')]
    # Subset Program Activities for SNAP-Ed
    pa_data = pa_data.loc[pa_data['program_areas'].str.contains('SNAP-Ed'), ['program_id', 'site_id']]

    # Import PSE Site Activity data, Needs, Readiness, Effectiveness, and Changes
    pse_site_activities_export = pd.ExcelFile(export_dir + "PSE_Site_Activity_Export.xlsx")
    pse_data = pd.read_excel(pse_site_activities_export, 'PSE Data')[['pse_id', 'site_id']]

    # Create a class, list of objects for these lists
    related_records = [ia_ic_data, pa_data, pse_data]
    related_ids = ['activity_id', 'program_id', 'pse_id']
    count_labels = ['related_indirect_activities', 'related_program_activities', 'related_pse_site_activities']

    for index, related_df in enumerate(related_records):
        part_data = utils.count_related_records(primary_records=part_data,
                                                primary_id='partnership_id',
                                                related_records=related_df,
                                                merge_on='site_id',
                                                related_id=related_ids[index],
                                                count_label=count_labels[index],
                                                binary=True)

    # Partnerships that require updates to intervention type fields
    part_int = part_data.loc[((part_data['is_direct_education_intervention'] == 0)
                              & ((part_data['related_indirect_activities'] == 1)
                                 | (part_data['is_direct_education_intervention'] == 1)))
                             | ((part_data['is_pse_intervention'] == 0)
                                & (part_data['related_pse_site_activities'] == 1))]

# Just excel workbook or send notifications to users?
    utils.write_report(file=output_dir + 'Update Partnerships Intervention Type.xlsx',
                       sheet_names=['Partnerships'],
                       dfs=[part_int])


# Run report directly from command line
# Parse inputs with argparse
# if __name__ == '__main__':
#     main()
