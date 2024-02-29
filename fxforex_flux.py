"""
Created on Mon Jul 31 16:45:35 2023
@author: MoysesRibeiro & SabrineSousa

Tool to provide data and analysis for Forex fluctuations

"""

import pandas as pd
import numpy as np
import os
from sap_gui import SapGui

import warnings

warnings.simplefilter(action='ignore')

__COMMENTARY_THRESHOLD = int((1 * 10 ** 6))
__SCALE = 10 ** 6


def analysis_comment(i, j, analysis_dict, quote_flag):
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


    if quote_flag: # se depreciou
        if flag and asset:
            analysis = analysis_dict.get(1)

        if not flag and not asset:
            analysis = analysis_dict.get(2)

    else: #se apreciou
        if flag and not asset:
            analysis = analysis_dict.get(3)

        if not flag and asset:
            analysis = analysis_dict.get(4)

    return analysis

   #analysis_dict = {1: 'Forex gain - asset increase, as expected',
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

    if int(j) > 500:
        flag = False

    return flag


#### Data Download from SAP
def get_data_from_sap_and_analyze(sap_instance, year, period, period_comp, RU, variant, quote_flag, currency):
    file_name = f'Forex_ANALYSIS_{RU}_{year}_{period}.xlsx'

    SAP = SapGui(sap_instance)
    session = SAP.session

    os.system('taskkill -f -im excel.exe')

    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nf.01"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = "0095000000"
    session.findById("wnd[0]/usr/ctxtSD_SAKNR-HIGH").text = "0096999999"
    session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = RU
    session.findById("wnd[0]/usr/ctxtSD_CURTP").text = "30"
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILBJAHR").text = year
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-HIGH").text = period
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILVJAHR").text = year
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").text = period_comp
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").setFocus()
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").caretPosition = 2
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").text = variant
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").caretPosition = 4
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2").select()

    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").text = "3"
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").setFocus()
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").caretPosition = 1
    session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1").select()
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/radBILAGRID").select()
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAGVAR").text = "/MOYSES"
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAGVAR").setFocus()
    session.findById(
        "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAGVAR").caretPosition = 7
    if currency.upper() == "USD":
        session.findById("wnd[0]/usr/ctxtSD_CURTP").text = "10"
    session.findById("wnd[0]").sendVKey(8)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").pressF4()
    session.findById("wnd[1]").close()
    session.findById("wnd[0]/tbar[1]/btn[29]").press()
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press()
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = RU
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 4
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.environ["USERPROFILE"]
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "forex_tool.xlsx"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    os.system('taskkill -f -im excel.exe')

    #### Analysis
    df = pd.read_excel(os.environ["USERPROFILE"] + r"\forex_tool.xlsx")
    df['Delta'] = df[df.columns[-2]] - df[df.columns[-1]]
    df['Account Number'].astype(str)[1][:2]

    x = 0
    df['++'] = 0
    for i in df['Account Number']:
        df['++'][x] = str(i)[:2]
        x = x + 1

    x = 0
    df['Concept'] = '0'

    for i in df[df.columns[2]]:

        if i is np.nan:
            df['Concept'][x] = i

        else:
            df['Concept'][x] = i.split(',')[3]
        x += 1

    df = df.dropna()
    df = df[df['Account Number'] != 96959799]
    df = df[df['Account Number'] != 95959799]

    pivot = pd.pivot_table(df, values=df.columns[-5:-2], index='++', aggfunc=np.sum)
    pivot = np.round(pivot[[pivot.columns[1], pivot.columns[2], pivot.columns[0]]] / __SCALE, 0)

    # x = 0
    # df['Main'] = '0'
    # for i in df['Concept']:
    #    df["Main"][x] = i[:3]
    #    x += 1

    Realized_Forex = df[df['++'] == '95']
    Realized_Forex = np.round(Realized_Forex.groupby('Concept').sum()[Realized_Forex.columns[-5:-2]] / __SCALE, 0)

    Unrealized_Forex = df[df['++'] == '96']
    Unrealized_Forex = np.round(Unrealized_Forex.groupby('Concept').sum()[Unrealized_Forex.columns[-5:-2]] / __SCALE, 0)

    Commentary = df.groupby(['++', 'Concept']).sum()[df.columns[-5:-2]]
    Commentary = np.round(Commentary[abs(Commentary['Delta']) > __COMMENTARY_THRESHOLD] / __SCALE, 0)
    Commentary['Analysis'] = '-'

    x = 0
    analysis_dict = {1: 'Forex gain - asset increase, as expected',
                     2: 'Forex loss - liability increase, as expected',
                     3: 'Forex gain - liability decrease, as expected',
                     4: 'Forex loss - asset decrease, as expected'}

    for i, j in zip(Commentary['Delta'], Commentary.reset_index()['Concept']):
        analysis = analysis_comment(i, j, analysis_dict, quote_flag)
        Commentary['Analysis'][x] = analysis
        x += 1

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(os.environ["USERPROFILE"] + f"\Desktop\\{file_name}", engine="xlsxwriter")

    Commentary.to_excel(writer, sheet_name="Commentary")

    pivot.to_excel(writer, sheet_name="Totals")
    Realized_Forex.to_excel(writer, sheet_name="Realized Forex")
    Unrealized_Forex.to_excel(writer, sheet_name="Unrealized Forex")

    quote = pd.read_csv(f"{os.environ['USERPROFILE']}\\quotes.csv")
    quote.to_excel(writer, sheet_name="Quotes")

    writer.close()


#### Data Download from SAP
def analyze_data(df, year, period, RU, quote_flag):
    file_name = f'Forex_ANALYSIS_{RU}_{year}_{period}.xlsx'

    #### Analysis
    df = df

    x = 0
    df['++'] = 0
    for i in df['G/L Acct']:
        df['++'][x] = str(i)[:4]
        x = x + 1

    # df.to_excel('testing.xlsx')

    pivot = None
    try:
        pivot = pd.pivot_table(df, index='++', aggfunc=np.sum)[['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']]
        pivot = np.round(pivot / __SCALE, 3)
    except:
        print('Error on creating pivot summary')

    Realized_Forex = df[df['++'] == '0095']
    Realized_Forex = Realized_Forex.groupby(["++", "Main", 'Description']).sum()[
        ['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']]
    Realized_Forex[['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']] = np.round(
        Realized_Forex[['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']] / __SCALE, 3)

    Unrealized_Forex = df[df['++'] == '0096']
    Unrealized_Forex = Unrealized_Forex.groupby(["++", "Main", 'Description']).sum()[
        ['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']]
    Unrealized_Forex[['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']] = np.round(
        Unrealized_Forex[['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']] / __SCALE, 3)

    Commentary = df.groupby(["++", "Main", 'Description']).sum()[['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']]
    Commentary = Commentary[abs(Commentary['CURRENT_MONTH']) > __COMMENTARY_THRESHOLD]
    Commentary[['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']] = np.round(
        Commentary[['CURRENT_PERIOD', 'PREVIOUS_PERIOD', 'CURRENT_MONTH']] / __SCALE, 3)
    Commentary['% Change'] = np.round((Commentary["CURRENT_MONTH"] / Commentary["PREVIOUS_PERIOD"]) * 100, 2)
    Commentary['Analysis'] = '-'

    x = 0
    analysis_dict = {1: 'Forex gain - asset increase, as expected',
                     2: 'Forex loss - liability increase, as expected',
                     3: 'Forex gain - liability decrease, as expected',
                     4: 'Forex loss - asset decrease, as expected'}

    for i, j in zip(Commentary['CURRENT_MONTH'], Commentary.reset_index()['Main']):
        analysis = analysis_comment(i, int(j), analysis_dict, quote_flag)
        Commentary['Analysis'][x] = analysis
        x += 1

    Commentary.rename(columns={'CURRENT_MONTH': "Change in $M", "PREVIOUS_PERIOD": "Previous Period in $M",
                               'CURRENT_PERIOD': 'Current Period in $M'}, inplace=True)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(os.environ["USERPROFILE"] + f"\Desktop\\ForexAnalysisOutput\\{file_name}",
                            engine="xlsxwriter")

    Commentary.to_excel(writer, sheet_name="Commentary")

    pivot.to_excel(writer, sheet_name="Totals")
    Realized_Forex.to_excel(writer, sheet_name="Realized Forex")
    Unrealized_Forex.to_excel(writer, sheet_name="Unrealized Forex")

    quote = pd.read_csv(f"{os.environ['USERPROFILE']}\\quotes.csv")
    quote.to_excel(writer, sheet_name="Quotes")

    writer.close()
