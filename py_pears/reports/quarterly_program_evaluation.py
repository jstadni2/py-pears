import pandas as pd
from datetime import datetime
from functools import reduce


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
