from openpyxl import load_workbook
import pandas as pd
from PatientInfo import *
from constants import *

def get_query_info():
    df = pd.read_excel(SOURCE_FILE_PATH, sheet_name=SOURCE_SHEET_NAME)
    df = df.reindex()
    df = df[['FIRST', 'LAST', 'MEDICARE #', 'DOB', 'From DOS', 'To DOS']]
    df['Invalid'] = df.isna().any(axis=1)
    df['DOB'] = pd.to_datetime(df['DOB'] , dayfirst=True)
    df.loc[df['Invalid'] == False, 'DOB'] = df.loc[df['Invalid'] == False, 'DOB'].apply(
        lambda dob: dob.strftime('%m/%d/%Y'))
    df.loc[df['Invalid'] == False, 'From DOS'] = df.loc[df['Invalid'] == False, 'From DOS'].apply(
        lambda dos: dos.strftime('%m/%d/%Y'))
    df.loc[df['Invalid'] == False, 'To DOS'] = df.loc[df['Invalid'] == False, 'To DOS'].apply(
        lambda dos: dos.strftime('%m/%d/%Y'))
    df['MEDICARE #'] = df['MEDICARE #'].apply(lambda num: str(num).strip())
    return df


def get_list_of_queries():
    result = []
    df = get_query_info()
    num_of_entries = len(df)
    for i in range(num_of_entries):
        query_info = QueryInfo(df['FIRST'][i], df['LAST'][i], df['MEDICARE #'][i], df['DOB'][i], df['From DOS'][i],
                               df['To DOS'][i], df['Invalid'][i])
        if not query_info.invalid:
            result.append(query_info)
    return result