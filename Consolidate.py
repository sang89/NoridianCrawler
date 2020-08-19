from openpyxl import load_workbook
import xlsxwriter
from constants import *
import pandas as pd
import numpy as np
from ExcelUtils import *
from utils import *
from datetime import datetime
from enum import Enum


def color_rows(s, length):
    return ['background-color: yellow' if s['count'] == -1 else 'background-color: white' for _ in range(length)]


def coloring_invalid(s, length):
    return ['color: red' if (s['Invalid'] and pd.isna(s['count'])) or (s['Result'] == 1)
            else 'color: black' for _ in range(length)]


def change_date_format(s):
    if isinstance(s, str):
        try:
            stripped_str = s.strip()
            result = convert_to_datetime_obj(stripped_str).strftime('%m/%d/%Y')
            return result
        except:
            print('Convert invalid date to empty string:', s)
            return ''
    else:
        return s.strftime('%m/%d/%Y')


def consolidate():
    main_df = pd.read_excel(MAIN_EXCEL_FILE_NAME, sheet_name=CONSOLIDATE_MAIN_SHEET_NAME)
    main_df.rename( columns={'Unnamed: 0':'count'}, inplace=True )
    main_df.insert(0, 'Sent out', ['' for _ in range(len(main_df))])

    source_df = get_query_info()
    invalid_source_df = source_df[source_df['Invalid'] == True]
    valid_source_df = source_df[source_df['Invalid'] == False]

    # make sure From DOS and To DOS is a datetime object for sorting
    main_df['From DOS'] = pd.to_datetime(main_df['From DOS'], format='%m/%d/%Y', errors='ignore', dayfirst=True).dt.date
    main_df['To DOS'] = pd.to_datetime(main_df['To DOS'], format='%m/%d/%Y', errors='ignore', dayfirst=True).dt.date
    #main_df = main_df.sort_values(by=['MEDICARE #', 'From DOS'])

    valid_source_df['count'] = [-1 for _ in range(len(valid_source_df))]
    merge_df = pd.concat([main_df, valid_source_df], axis=0, ignore_index=True)

    # Sort the table
    #merge_df = merge_df.sort_values(by=['MEDICARE #', 'count'])
    merge_df = merge_df.sort_values(by=['LAST', 'FIRST', 'count'])

    merge_df.loc[merge_df['count'] >= 0, 'MEDICARE #'] = ''
    merge_df.loc[merge_df['count'] >= 0, 'DOB'] = ''
    #merge_df.loc[merge_df['count'] > 0, 'FIRST IN NOR'] = merge_df.loc[merge_df['count'] > 0, 'FIRST IN NOR'].apply(lambda elem: '')
    #merge_df.loc[merge_df['count'] > 0, 'LAST IN NOR'] = merge_df.loc[merge_df['count'] > 0, 'LAST IN NOR'].apply(
    #    lambda elem: '')
    merge_df.loc[merge_df['count'] > 0, 'count'] = merge_df.loc[merge_df['count'] > 0, 'count'].apply(
        lambda elem: 0)

    # Sort the table
    #merge_df = merge_df.sort_values(by=['MEDICARE #', 'count', 'From DOS', 'To DOS'])
    merge_df = merge_df.sort_values(by=['LAST', 'FIRST', 'count', 'From DOS', 'To DOS'])

    #.loc[merge_df['count'] >= 0, 'MEDICARE #'] = merge_df.loc[merge_df['count'] > 0, 'MEDICARE #'].apply(
    #    lambda elem: '')
    #merge_df.loc[merge_df['count'] >= 0, 'FIRST'] = merge_df.loc[merge_df['count'] > 0, 'FIRST'].apply(lambda elem: '')
    #merge_df.loc[merge_df['count'] >= 0, 'LAST'] = merge_df.loc[merge_df['count'] > 0, 'LAST'].apply(lambda elem: '')

    # Convert string back to our format
    merge_df.loc[merge_df['count'] == 0, 'From DOS'] = merge_df['From DOS'].apply(lambda dos: change_date_format(dos))
    merge_df.loc[merge_df['count'] == 0, 'To DOS'] = merge_df['To DOS'].apply(lambda dos: change_date_format(dos))

    # Clean the invalid data frame
    invalid_source_df = invalid_source_df.applymap(lambda el: '' if pd.isna(el) or el=='nan' else el)
    invalid_source_df['From DOS'] = invalid_source_df['From DOS'].apply(
        lambda dos: '' if pd.isna(dos) or dos=='' else change_date_format(dos))
    invalid_source_df['To DOS'] = invalid_source_df['To DOS'].apply(
        lambda dos: '' if pd.isna(dos) or dos=='' else change_date_format(dos))
    invalid_source_df['DOB'] = invalid_source_df['DOB'].apply(
        lambda dob: '' if pd.isna(dob) or dob=='' else change_date_format(dob))

    # Concat the invalid at the end
    merge_df = pd.concat([merge_df, invalid_source_df], axis=0, ignore_index=True)

    # post processing
    merge_df = postprocessing(merge_df)

    # Apply the color
    num_of_cols = len(merge_df.columns)
    merge_df = merge_df.style.apply(color_rows, axis=1, length = num_of_cols).apply(
        coloring_invalid, axis=1, length = num_of_cols)

    book = load_workbook(CONSOLIDATE_FILE_NAME)
    with pd.ExcelWriter(CONSOLIDATE_FILE_NAME, engine='openpyxl') as writer:
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        merge_df.to_excel(writer, sheet_name=CONSOLIDATE_SHEET_NAME)
        writer.save()

def enum(*sequential, **named):
    enums = dict(zip(sequential, range(len(sequential))), **named)
    return type('Enum', (), enums)

DataSource = enum('ORG', 'CRAWL_WITH_ORG', 'CRAWL_WITH_NAME_FIX', 'NOT_FOUND')
Result = enum('OK', 'Mismatch')

def postprocessing(consolidated_df):
    new_df = consolidated_df

    # insert 2 new postprocessed columns
    new_df.insert(0, 'Data Source', [DataSource.CRAWL_WITH_ORG for _ in range(len(new_df))])
    new_df.insert(1, 'Result', ['' for _ in range(len(new_df))])
    new_df.insert(12, 'Missing Dates', ['' for _ in range(len(new_df))])
    new_df.insert(13, 'Extra Dates', ['' for _ in range(len(new_df))])

    mask = (new_df['FIRST'] != new_df['FIRST IN NOR']) | (new_df['LAST'] != new_df['LAST IN NOR']) \
        & (new_df['Data Source'] == DataSource.NOT_FOUND)
    new_df.loc[mask, 'Data Source'] = DataSource.CRAWL_WITH_NAME_FIX
    new_df.loc[pd.isna(new_df['Units']), 'Data Source'] = DataSource.NOT_FOUND
    new_df.loc[new_df['count'] == -1, 'Data Source'] = DataSource.ORG
    new_df.loc[new_df['Invalid'] == True, 'Data Source'] = DataSource.NOT_FOUND

    new_df.loc[new_df['count'] >= 0, 'FIRST'] = ''
    new_df.loc[new_df['count'] >= 0, 'LAST'] = ''

    start_date = ''
    end_date = ''
    service_dates = []
    negative_cnt_index = 0
    invalid_indices = []
    missing_dates = []
    extra_dates = []
    for row_cntr in range(len(new_df)):
        cur_count = new_df.at[row_cntr, 'count']
        if cur_count == -1:
            if len(service_dates) > 0:
                invalid_indices, missing_dates, extra_dates = \
                    check_if_service_dates_in_range(service_dates, start_date, end_date)
            start_date = convert_to_datetime_obj(new_df.at[row_cntr, 'From DOS'])
            end_date = convert_to_datetime_obj(new_df.at[row_cntr, 'To DOS'])
            service_dates = []
            for i in range(len(invalid_indices)):
                new_df.at[negative_cnt_index + 1 + invalid_indices[i], 'Result'] = 1

            if len(missing_dates) > 0:
                new_df.at[negative_cnt_index, 'Missing Dates'] = missing_dates
            if len(extra_dates) > 0:
                new_df.at[negative_cnt_index, 'Extra Dates'] = extra_dates

            invalid_indices = []
            missing_dates = []
            extra_dates = []
            negative_cnt_index = row_cntr
        elif cur_count == 0:
            service_date_start = convert_to_datetime_obj(new_df.at[row_cntr, 'From DOS'])
            service_dates_end = convert_to_datetime_obj(new_df.at[row_cntr, 'To DOS'])
            cur_date_range = [service_date_start, service_dates_end]
            service_dates.append(cur_date_range)

    new_df.loc[new_df['count'] == -1, 'Result'] = ''

    return new_df

if __name__ == '__main__':
    consolidate()

