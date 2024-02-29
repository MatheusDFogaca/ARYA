import snowflake.connector as sf
import datetime as dt
import sys

'''
Basic Structure
To connect into FED use:
Open cmd
write: setx EMAIL "your email"
Open and reopen cmd
write: echo %EMAIL%
Check if the email appearing is correct
'''
class SnowSQL():
        def __init__(self,user):
           try:
                self.con = sf.connect(
                    user=user,
                    account='xom-fin',  # account identifier
                    authenticator='externalbrowser',
                    warehouse='BI_FED_WH'  # warehouse can be changed according to database requiremets
                )
           except:
                print('FED connection error...')
                sys.exit()


        def execute_query(self,query):
            now = dt.datetime.now()
            print("Querying data...")
            try:
                cursor = self.con.execute_string(query)[0]
                print("Query successfull")
                print(f'Elapsed time :{dt.datetime.now()-now}')
                return cursor
            except:
                print("There was an error querying the database")
                return []

import pandas as pd
import numpy as np
import os


import warnings
warnings.simplefilter(action='ignore')

__COMMENTARY_THRESHOLD = int(1*10**6)
__SCALE = 10**3

def analysis_comment(i,j, analysis_dict, quote_flag):
    """
    i := amount of change
    j := is main account
    asset := True or False (True if asset, False is liability)
    analysis := analysis result (default is 'COMMENT REQUIRED', if different dictionary analysis_dict wil show)
    quote_flag := if forex depreciation or apreciation (True if depreciation, False if apreciation)
    flag := delta, is the change on from one period to another.
    """

    analysis = 'COMMENT REQUIRED'
    flag = is_positive_flag(i)
    asset = is_asset_or_liability(j)

    if quote_flag:  # se depreciou
        if flag and asset:
            analysis = analysis_dict.get(1)

        if not flag and not asset:
            analysis = analysis_dict.get(2)

    else:  # se apreciou
        if flag and not asset:
            analysis = analysis_dict.get(3)

        if not flag and asset:
            analysis = analysis_dict.get(4)

    return analysis


# analysis_dict = {1: 'Forex gain - asset increase, as expected',
#                 2: 'Forex loss - liability increase, as expected',
#                 3: 'Forex gain - liability decrease, as expected',
#                 4: 'Forex loss - asset decrease, as expected'}

def is_positive_flag(i):
    flag = True
    if i > 0:
        flag = False
    return flag


def is_asset_or_liability(j):
    flag = True

    if int(j[:3]) > 500:
        flag = False


    return flag

#### Data Download from SAP
def get_data_from_fed_and_analyze(df_from_fed,year, period, period_comp,RU,variant, quote_flag, currency):
    file_name = f'Forex_ANALYSIS_{RU}_{year}_{period}.xlsx'

    #### Analysis
    df = df_from_fed
    df['Delta'] = df[df.columns[-2]]-df[df.columns[-1]]
    df['Account Number'].astype(str)[1][:2]

    x = 0
    df['++'] = 0
    for i in df['Account Number']:
        df['++'][x] = str(i)[:2]
        x = x + 1



    x=0
    df['Concept']='0'


    for i in df[df.columns[2]]:

        if i is np.nan:
            df['Concept'][x] = i

        else:
            df['Concept'][x] = i.split(',')[3]
        x+=1


    df=df.dropna()
    df=df[df['Account Number']!=96959799]
    df=df[df['Account Number']!=95959799]


    pivot = pd.pivot_table(df, values=df.columns[-5:-2], index='++', aggfunc=np.sum)
    pivot = np.round(pivot[[pivot.columns[1], pivot.columns[2], pivot.columns[0]]]/__SCALE,0)


    #x = 0
    #df['Main'] = '0'
    #for i in df['Concept']:
    #    df["Main"][x] = i[:3]
    #    x += 1

    Realized_Forex = df[df['++'] == '95']
    Realized_Forex = np.round(Realized_Forex.groupby('Concept').sum()[Realized_Forex.columns[-5:-2]]/__SCALE,0)

    Unrealized_Forex = df[df['++'] == '96' ]
    Unrealized_Forex = np.round(Unrealized_Forex.groupby('Concept').sum()[Unrealized_Forex.columns[-5:-2]]/__SCALE,0)


    Commentary = df.groupby(['++','Concept']).sum()[df.columns[-5:-2]]
    Commentary = np.round(Commentary[abs(Commentary['Delta']) > __COMMENTARY_THRESHOLD]/__SCALE,0)
    Commentary['Analysis'] = '-'

    x = 0
    analysis_dict = {1: 'Forex gain - asset increase, as expected',
                     2: 'Forex loss - liability increase, as expected',
                     3: 'Forex gain - liability decrease, as expected',
                     4: 'Forex loss - asset decrease, as expected'}

    for i , j in zip(Commentary['Delta'],Commentary.reset_index()['Concept']):
        analysis = analysis_comment(i,j, analysis_dict, quote_flag)
        Commentary['Analysis'][x] = analysis
        x += 1



    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(os.environ["USERPROFILE"] + f"\Desktop\\{file_name}", engine="xlsxwriter")

    Commentary.to_excel(writer, sheet_name="Commentary")

    pivot.to_excel(writer, sheet_name="Totals")
    Realized_Forex.to_excel(writer, sheet_name="Realized Forex")
    Unrealized_Forex.to_excel(writer, sheet_name="Unrealized Forex")

    quote = pd.read_csv(f"{os.environ['USERPROFILE']}\\quotes.csv")
    quote.to_excel(writer,sheet_name="Quotes")

    writer.close()