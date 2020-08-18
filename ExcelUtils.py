import pandas as pd
from PatientInfo import *
from constants import *
from utils import *

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

def get_query_info():
    df = pd.read_excel(SOURCE_FILE_PATH, sheet_name=SOURCE_SHEET_NAME)
    df = df.reindex()
    df = df[['FIRST', 'LAST', 'MEDICARE #', 'DOB', 'From DOS', 'To DOS', 'Facesheet', 'Sent out']]
    df['Invalid'] = df[['FIRST', 'LAST', 'MEDICARE #', 'DOB', 'From DOS', 'To DOS']].isna().any(axis=1)
    #df['DOB'] = pd.to_datetime(df['DOB'] , dayfirst=True)
    df.loc[df['Invalid'] == False, 'DOB'] = df.loc[df['Invalid'] == False, 'DOB'].apply(
        lambda dob: change_date_format(dob))
    df.loc[df['Invalid'] == False, 'From DOS'] = df.loc[df['Invalid'] == False, 'From DOS'].apply(
        lambda dos: change_date_format(dos))
    df.loc[df['Invalid'] == False, 'To DOS'] = df.loc[df['Invalid'] == False, 'To DOS'].apply(
        lambda dos: change_date_format(dos))
    df['MEDICARE #'] = df['MEDICARE #'].apply(lambda num: str(num).strip())

    df.loc[df['DOB'] == '', 'Invalid'] = df.loc[df['DOB'] == '', 'Invalid'].apply(lambda invalid: True)
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
