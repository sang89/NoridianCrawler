from openpyxl import load_workbook
import xlsxwriter
from constants import *
import pandas as pd
import numpy as np
from ExcelUtils import *
from utils import *
from datetime import datetime

def color_rows(s):
    return ['background-color: yellow' if s['count'] == -1 else 'background-color: white' for _ in range(20)]

def coloring_invalid(s):
    return ['color: red' if s['Invalid'] and pd.isna(s['count']) else 'color: black' for _ in range(20)]

def change_date_format(s):
    if isinstance(s, str):
        return convert_to_datetime_obj(s).strftime('%m/%d/%Y')
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

    merge_df.loc[merge_df['count'] >= 0, 'FIRST'] = merge_df.loc[merge_df['count'] > 0, 'FIRST'].apply(lambda elem : '')
    merge_df.loc[merge_df['count'] >= 0, 'LAST'] = merge_df.loc[merge_df['count'] > 0, 'LAST'].apply(lambda elem : '')
    #merge_df.loc[merge_df['count'] > 0, 'MEDICARE #'] = merge_df.loc[merge_df['count'] > 0, 'MEDICARE #'].apply(lambda elem: '')
    merge_df.loc[merge_df['count'] >= 0, 'DOB'] = merge_df.loc[merge_df['count'] > 0, 'DOB'].apply(lambda elem: '')
    merge_df.loc[merge_df['count'] > 0, 'FIRST IN NOR'] = merge_df.loc[merge_df['count'] > 0, 'FIRST IN NOR'].apply(lambda elem: '')
    merge_df.loc[merge_df['count'] > 0, 'LAST IN NOR'] = merge_df.loc[merge_df['count'] > 0, 'LAST IN NOR'].apply(
        lambda elem: '')
    merge_df.loc[merge_df['count'] > 0, 'count'] = merge_df.loc[merge_df['count'] > 0, 'count'].apply(
        lambda elem: 0)

    # Sort the table
    merge_df = merge_df.sort_values(by=['MEDICARE #', 'count', 'From DOS', 'To DOS'])
    merge_df.loc[merge_df['count'] >= 0, 'MEDICARE #'] = merge_df.loc[merge_df['count'] > 0, 'MEDICARE #'].apply(
        lambda elem: '')

    # Convert string back to our format
    merge_df.loc[merge_df['count'] == 0, 'From DOS'] = merge_df['From DOS'].apply(lambda dos: change_date_format(dos))
    merge_df.loc[merge_df['count'] == 0, 'To DOS'] = merge_df['To DOS'].apply(lambda dos: change_date_format(dos))

    # Clean the invalid data frame
    invalid_source_df = invalid_source_df.applymap(lambda el: '' if pd.isna(el) or el=='nan' else el)
    invalid_source_df['From DOS'] = invalid_source_df['From DOS'].apply(
        lambda dos: '' if pd.isna(dos) or dos=='' else dos.strftime('%m/%d/%Y'))
    invalid_source_df['To DOS'] = invalid_source_df['To DOS'].apply(
        lambda dos: '' if pd.isna(dos) or dos=='' else dos.strftime('%m/%d/%Y'))
    invalid_source_df['DOB'] = invalid_source_df['DOB'].apply(
        lambda dob: '' if pd.isna(dob) or dob=='' else dob.strftime('%m/%d/%Y'))

    # Concat the invalid at the end
    merge_df = pd.concat([merge_df, invalid_source_df], axis=0, ignore_index=True)

    # Apply the color
    merge_df = merge_df.style.apply(color_rows, axis=1).apply(coloring_invalid, axis=1)

    book = load_workbook(CONSOLIDATE_FILE_NAME)
    with pd.ExcelWriter(CONSOLIDATE_FILE_NAME, engine='openpyxl') as writer:
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        merge_df.to_excel(writer, sheet_name=CONSOLIDATE_SHEET_NAME)
        writer.save()

if __name__ == '__main__':
    consolidate()
