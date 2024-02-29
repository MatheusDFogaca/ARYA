"""
File related to the program GUI
using mainly tkinter facilities to
display the user interface

to build the .exe file use code to follow:python -m PyInstaller --additional-hooks-dir . --noconfirm --onefile main.py


"""

import tkinter as tk
from tkinter import filedialog
from datetime import timedelta
from sap_gui import SapGui
from fed_extraction import SnowSQL
import numpy as np
import guistyles as styles
from rvworkbook import *
from fxexcel_formating import *
from fxlook_up_table import look_up_table as look_up_table_from_database
from fxforex import get_dataframe, plot_png
from fxforex_flux import analyze_data, get_data_from_sap_and_analyze

AMP = "|SSO| AMP [NA-Stripes Prod FI/TAX]"
GEMS = "|SSO| G9P [GEMS Prod Financials]"
Deskpath = os.environ['userprofile'] + "\\Desktop\\"

x = dt.date.today()

y = x.isoformat()[:-2] + '01'
z = dt.date.fromisoformat(y)
year = z - dt.timedelta(days=1)
month = year.month
year = year.year


class MainGUI(tk.Tk):
    """
    Class to encapsulate all properties and functions to the main window of the program
    """
#set x email
    def __init__(self):
        """
        Initializate the program
        """
        super().__init__()
        self.__FED = SnowSQL(os.environ["EMAIL"])
        os.system('cls')
        self.rvcreate_folder()
        self.create_folder()
        print("Welcome to the ARYA tool:")

        self.look_up_table = look_up_table_from_database
        self.geometry("900x220")
        self.title('ARYA TOOL')
        self.upperFrame = tk.Frame(self, bg='purple')
        self.upperFrame.pack(fill="both", side=tk.TOP)
        self.Label1 = tk.Label(self.upperFrame, text="Automation for Revaluation and Analysis Activities", bg='purple', fg='white',
                               font=styles.header)
        self.Label1.pack(side='top')

        """
        ******************* F R A M E S ***********************
        """
        self.second_frame1 = tk.Frame(self)
        self.second_frame1.pack(fill='both', side=tk.TOP)

        self.ramdom_frame1 = tk.Frame(self)
        self.ramdom_frame1.pack(fill='both', side=tk.TOP)

        """
        Text variables Sections     
        """
        self.RU = tk.StringVar()
        self.RU.set('2042')
        self.year = tk.StringVar()
        self.year.set(str(year))
        self.period = tk.StringVar()
        self.period.set(str(month))
        self.system = tk.StringVar()
        self.system.set('G9P')
        self.currency = tk.StringVar()
        self.currency.set('MXN')

        l2 = tk.Label(self, text="Insert RU:", font=styles.header, fg='black')
        l2.pack(in_=self.second_frame1, side=tk.LEFT)
        tk.Entry(self, width=9, font=styles.body, textvariable=self.RU, fg='red').pack(in_=self.second_frame1,
                                                                                       side=tk.LEFT)

        l3 = tk.Label(self, text="Year:", font=styles.header, fg='black')
        l3.pack(in_=self.second_frame1, side=tk.LEFT)
        tk.Entry(self, width=10, font=styles.body, textvariable=self.year, fg='red').pack(in_=self.second_frame1,
                                                                                          side=tk.LEFT)

        l4 = tk.Label(self, text="Period:", font=styles.header, fg='black')
        l4.pack(in_=self.second_frame1, side=tk.LEFT)
        tk.Entry(self, width=10, font=styles.body, textvariable=self.period, fg='red').pack(in_=self.second_frame1,
                                                                                            side=tk.LEFT)

        l5 = tk.Label(self, text="System:", font=styles.header, fg='black')
        l5.pack(in_=self.second_frame1, side=tk.LEFT)
        tk.Entry(self, width=10, font=styles.body, textvariable=self.system, fg='red').pack(in_=self.second_frame1,
                                                                                            side=tk.LEFT)
        l6 = tk.Label(self, text="Currency", font=styles.header, fg='black')
        l6.pack(in_=self.second_frame1, side=tk.LEFT)
        tk.Entry(self, width=4, font=styles.body, textvariable=self.currency, fg='red').pack(in_=self.second_frame1,
                                                                                             side=tk.LEFT)

        tk.Label(text="Revaluation and Distribution Check",
                 font=styles.body).place(x=112, y=70)
        tk.Label(text="|",
                 font=styles.body).place(x=450, y=70)
        tk.Label(text="|",
                 font=styles.body).place(x=450, y=90)
        tk.Label(text="|",
                 font=styles.body).place(x=450, y=110)
        tk.Label(text="|",
                 font=styles.body).place(x=450, y=130)
        tk.Label(text="|",
                 font=styles.body).place(x=450, y=150)
        tk.Label(text="|",
                 font=styles.body).place(x=450, y=170)

        tk.Label(text="Flux Analysis",
                 font=styles.body).place(x=635, y=70)

        tk.Button(text="INDIVIDUAL RUN", font=styles.body, command=self.rv_run_individual_ru).place(x=100, y=100)
        tk.Button(text="MASSIVE", width=10, font=styles.body, command=self.rv_get_massive_reports).place(x=260, y=100)

        tk.Button(text="INDIVIDUAL RUN", font=styles.body, command=self.fx_run_individual_ru).place(x=530, y=100)
        tk.Button(text="MASSIVE", width=10, font=styles.body, command=self.fx_get_massive_reports).place(x=690, y=100)

    def rv_get_massive_reports(self):

        # Massive run for revaluation

        now = dt.datetime.now()

        filepath = self.get_filenames()
        df5 = pd.read_excel(filepath[0], dtype=str)
        error_list = []
        for i in df5.index:
            print(i)
            temp = df5.iloc[i]
            RU = temp[0]
            year = temp[1]
            period = temp[2]
            currency = temp[3]
            system = temp[4]
            lista = [RU, year, period, system]

            try:
                print("Querying data from:", lista)
                df_rv = self.get_revaluation_data_from_fed(RU, year, period, system, currency)
                self.Rates(self, RU, year, period, system, currency)
                df_db = self.get_distribution_data_from_fed(RU, year, period, system, currency)
                self.rev_get_data(df_rv, df_db, RU, year, period, system, currency)

            except:
                error_list.append(lista)

        print("MASSIVE JOB FINISHED")
        if len(error_list) != 0:
            print("Errors on the RUs to follow:")
            for i in error_list:
                print(i)

        later = dt.datetime.now()
        print(now, "+++", later)
        print(f"Elapsed time {later - now}")

    @staticmethod
    def get_filenames():
        """Function to call a filedialog to select excel files from DFX Spreadsheets folder
        so the program can start up a stream of excels to read an write"""
        filenames = filedialog.askopenfilenames(
            initialdir=os.environ["userprofile"], title="Select a File",
            filetypes=(("Excel", "*.xl*"), ("all files", "*.*")))

        return filenames


    def rv_run_individual_ru(self):

        # Individual run for revaluation

        # sap_instance = self.style.get()
        RU = self.RU.get()
        year = self.year.get()
        period = f"{'0' + str(self.period.get())}" if int(self.period.get()) < 10 else self.period.get()
        system = self.system.get().upper()
        currency = self.currency.get().upper()

        # self.Rates(RU, year, period, system, currency)
        df_rv = self.get_revaluation_data_from_fed(RU, year, period, system, currency)
        self.Rates(self, RU, year, period, system, currency)
        df_db = self.get_distribution_data_from_fed(RU, year, period, system, currency)
        self.rev_get_data(df_rv, df_db, RU, year, period, system, currency)
        # create_jv(RU, year, period, system, currency

    def get_revaluation_data_from_fed(self, RU, year, period, system, currency):

        # Getting data from accounts

        query = f'''
                          SELECT "Company Code" AS "RU", REPLACE(LTRIM(REPLACE("G/L Acct", '0', ' ')),' ', '0') AS "Account", "TCurr Key" AS "DCurr", "LCurr Key" AS "LCurr", "GCurr Key" AS "GCurr", SUM("TC Amount") AS "DC Amount", SUM("LC Amount") AS "LC Amount", SUM("GC Amount") AS "GC Amount"
            FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_JET_WWSL_TOTAL_ACTUAL_ALL 
            WHERE "Company Code" = '{RU}' 
                AND "Fiscal Yr" = '{year}'
                AND "Period" <= '{period}'
                AND LENGTH("Account") = '9'
                AND "Account" BETWEEN '10000000' and '79999999'
                AND "Record Type" = '0'
                AND "Source Table" ='YWS01T'
                AND "Version" = '1'
                AND NOT SUBSTR("Account", -6, 2) = '98'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '198001' AND '198008'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '198010' AND '198012'
                AND NOT SUBSTR("Account", -9, 3) BETWEEN '300' AND '400'
                AND NOT SUBSTR("Account", -9, 3) BETWEEN '190' AND '194'
                AND NOT SUBSTR("Account", -9, 3) BETWEEN '200' AND '249'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '478011' AND '478032'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '478039' AND '478042'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '478044' AND '478050'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '748102' AND '748105'
                AND NOT SUBSTR("Account", -9, 3) IN ('260', '269', '478', '645', '760')
                AND NOT SUBSTR("Account", -3, 3) IN ('757', '758', '783', '790', '798')
                AND NOT SUBSTR("Account", -9, 6) IN ('198977', '198200', '678050', '478060', '478200', '478033', '478034', '650001', '650003', '650005', '650977', '650200', '650210', 
        '748107', '748200', '748210')
                OR ("Company Code" = '{RU}'
                AND "Fiscal Yr" = '{year}'
                AND "Period" <= '{period}'
                AND LENGTH("Account") = '9'
                AND "Account" BETWEEN '10000000' and '79999999'
                AND "Record Type" = '2'
                AND "Source Table" ='YWS01T'
                AND "Version" = '1'
                AND NOT SUBSTR("Account", -6, 2) = '98'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '198001' AND '198008'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '198010' AND '198012'
                AND NOT SUBSTR("Account", -9, 3) BETWEEN '300' AND '400'
                AND NOT SUBSTR("Account", -9, 3) BETWEEN '190' AND '194'
                AND NOT SUBSTR("Account", -9, 3) BETWEEN '200' AND '249'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '478011' AND '478032'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '478039' AND '478042'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '478044' AND '478050'
                AND NOT SUBSTR("Account", -9, 6) BETWEEN '748102' AND '748105'
                AND NOT SUBSTR("Account", -9, 3) IN ('260', '269', '478', '645', '760')
                AND NOT SUBSTR("Account", -3, 3) IN ('757', '758', '783', '790', '798')
                AND NOT SUBSTR("Account", -9, 6) IN ('198977', '198200', '678050', '478060', '478200', '478033', '478034', '650001', '650003', '650005', '650977', '650200', '650210', 
        '748107', '748200', '748210'))
            GROUP BY "TCurr Key", "Company Code", "LCurr Key", "GCurr Key", "Account"

            '''
        df_rv = self.__FED.execute_query(query).fetch_pandas_all()
        df_rv["DC Amount"] = df_rv["DC Amount"].astype(float)
        df_rv["LC Amount"] = df_rv["LC Amount"].astype(float)
        df_rv["GC Amount"] = df_rv["GC Amount"].astype(float)

        return df_rv

    def get_distribution_data_from_fed(self, RU, year, period, system, currency):

        # Getting data from distribution

        Y0 = ['APP', 'EUP', 'G3P', "AMP"]
        Y1 = ['G9P']
        U1 = ['S8P']

        if system in Y0:
            query = f'''
            SELECT "Company Code" AS "RU", REPLACE(LTRIM(REPLACE("G/L Acct", '0', ' ')),' ', '0') AS "Account", SUM("TC Amount") AS "DC Amount", SUM("LC Amount") AS "LC Amount", SUM("GC Amount") AS "GC Amount"
                FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_JET_WWSL_TOTAL_ACTUAL_ALL 
                WHERE "Company Code" = '{RU}' 
                    AND "Fiscal Yr" = '{year}'
                    AND "Period" <= '{period}'
                    AND LENGTH("Account") = '8'
                    AND "Ledger" = 'Y0'
                    AND "Version" = '1'
                    AND ("Account" between '95000000' AND '96999999' OR "Account" = '21959000' OR "Account" = '21959799' OR "Account" = '71959000' OR "Account" = '71959799')
                    GROUP BY ROW("Company Code", "G/L Acct")
                '''
        elif system in Y1:
            query = f'''
            SELECT "Company Code" AS "RU", REPLACE(LTRIM(REPLACE("G/L Acct", '0', ' ')),' ', '0') AS "Account", SUM("TC Amount") AS "DC Amount", SUM("LC Amount") AS "LC Amount", SUM("GC Amount") AS "GC Amount"
                FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_JET_WWSL_TOTAL_ACTUAL_ALL 
                WHERE "Company Code" = '{RU}' 
                    AND "Fiscal Yr" = '{year}'
                    AND "Period" <= '{period}'
                    AND LENGTH("Account") = '8'
                    AND "Ledger" = 'Y1'
                    AND "Version" = '1'
                    AND ("Account" between '95000000' AND '96999999' OR "Account" = '21959000' OR "Account" = '21959799' OR "Account" = '71959000' OR "Account" = '71959799')
                    GROUP BY ROW("Company Code", "G/L Acct")
                '''
        elif system in U1:
            query = f'''
            SELECT "Company Code" AS "RU", REPLACE(LTRIM(REPLACE("G/L Acct", '0', ' ')),' ', '0') AS "Account", SUM("TC Amount") AS "DC Amount", SUM("LC Amount") AS "LC Amount", SUM("GC Amount") AS "GC Amount"
                FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_JET_WWSL_TOTAL_ACTUAL_ALL 
                WHERE "Company Code" = '{RU}' 
                    AND "Fiscal Yr" = '{year}'
                    AND "Period" <= '{period}'
                    AND LENGTH("Account") = '8'
                    AND "Ledger" = 'U1'
                    AND "Version" = '1'
                    AND ("Account" between '95000000' AND '96999999' OR "Account" = '21959000' OR "Account" = '71959799')
                    GROUP BY ROW("Company Code", "G/L Acct")
                '''

        df_db = self.__FED.execute_query(query).fetch_pandas_all()
        df_db['Base Account'] = df_db.Account.apply(lambda x:x[:2])
        df_db = df_db.drop('Account', axis=1)
        df_db = pd.pivot_table(df_db, index=['RU', 'Base Account'], values=['DC Amount', 'LC Amount', 'GC Amount'], aggfunc=np.sum, margins=True, margins_name='Total')

        return df_db

    @staticmethod
    def rvcreate_folder():

        # Create a folder for revaluation if necessary

        p = f'{os.environ["USERPROFILE"]}\\Desktop\\ForexRevaluationOutput'
        if not os.path.isdir(p):
            os.system(f'mkdir {p}')

    @staticmethod
    def rev_get_data(df_rv, df_db, RU, year, period, system, currency):

        # Here, all the main functions will be executed

        create_workbook(df_rv, df_db, RU, year, period, system, currency)
        xl, wb = open_rv_file(
            f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput\Forex_Revaluation_{RU}_{year}_{period}.xlsx')
        xl.DisplayAlerts = False
        fed_data(wb, currency)
        wb.SaveAs(f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput\Forex_Revaluation_{RU}_{year}_{period}.xlsx', ConflictResolution=2)
        wb.Close()
        xl, wb = open_rv_file(
            f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput\Forex_Revaluation_{RU}_{year}_{period}.xlsx')
        xl.DisplayAlerts = False
        power_pivot(RU, year, period, system, currency, xl, wb)
        distribution_check(wb)
        rv_excel_formating(RU, year, period, wb, xl)
        wb.SaveAs(
            f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput\Forex_Revaluation_{RU}_{year}_{period}.xlsx', ConflictResolution=2)
        wb.Close()
        xl.Quit()

        print(
            f'File successfully saved at: {os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput\Forex_Revaluation_{RU}_{year}_{period}.xlsx')

    def Rates(self, df_rv, RU, year, period, system, currency):

        # Currency extraction from SAP

        if int(self.period.get()) == 12:
            x = dt.date(int(year) + 1, int(period) - 11, 1)
            print(x)
        else:
            x = dt.date(int(year), int(period) + 1, 1)

        fd = pd.offsets.MonthBegin().rollback(x)
        first_wd = fd.strftime('%m/%d/%Y')
        ad = pd.offsets.MonthBegin().rollback(x) - timedelta(days=1)
        antepenultimate_wd = ad.strftime('%m/%d/%Y')

        df_rv = self.get_revaluation_data_from_fed(RU, year, period, system, currency)

        df = pd.concat([df_rv['DCurr'], df_rv['LCurr'], df_rv['GCurr']], axis=0)
        df = df.drop_duplicates()
        df = df.dropna()
        df.to_string(f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput\Rates.txt', index=False)

        if system != 'G9P' and currency == 'USD':
            return

        else:

            # Variables

            session = SapGui(GEMS).get_session
            ipes_currency_type = {"7248": "1MOZ", "4242": "1MOZ", "4241": "1MOZ", '3322': "1GUY", "3321": "1GUY",
                                  "1575": "1RUS", "1958": "1BRA", "2204": "1AOB", "3359": "1AOB", "4107": "1MOZ",
                                  "4108": "1MOZ", "4109": "1MOZ", "2116": "1MOZ", "2216": "1AOB", "2123": "1AOB",
                                  "2864": "1DZD", "4113": "1EGP", "4114": "1EGP", "4239": "EURX", "4240": "EURX",
                                  "4932": "EURX", "4933": "EURX"}

            if type(ipes_currency_type.get(RU)) == type(None):
                KURST = "001B"

            else:
                KURST = ipes_currency_type.get(RU)

            ExportName = "Currencies.txt"

            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "SE16"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "YFIA_FOREX_DSTDT"
            session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 16
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtI1-LOW").text = antepenultimate_wd
            session.findById("wnd[0]/usr/ctxtI1-HIGH").text = first_wd
            session.findById("wnd[0]/usr/ctxtI10-LOW").text = KURST
            session.findById("wnd[0]/usr/ctxtI10-LOW").setFocus()
            session.findById("wnd[0]/usr/ctxtI10-LOW").caretPosition = 4
            session.findById("wnd[0]/usr/btn%_I8_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[23]").press()
            session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
            session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
            session.findById("wnd[2]").sendVKey(4)
            session.findById("wnd[3]/usr/ctxtDY_PATH").text = f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput'
            session.findById("wnd[3]/usr/ctxtDY_FILENAME").text = "Rates.txt"
            session.findById("wnd[3]/usr/ctxtDY_FILENAME").caretPosition = 9
            session.findById("wnd[3]/tbar[0]/btn[0]").press()
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/usr/btn%_I9_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[23]").press()
            session.findById("wnd[2]/usr/ctxtDY_PATH").text = f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput'
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "Rates.txt"
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 9
            session.findById("wnd[2]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]").select()
            session.findById(
                "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            session.findById(
                "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput'
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ExportName
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
            session.findById("wnd[1]/tbar[0]/btn[11]").press()

    @staticmethod
    def create_folder():
        p = f'{os.environ["USERPROFILE"]}\\Desktop\\ForexAnalysisOutput'
        if not os.path.isdir(p):
            os.system(f'mkdir {p}')


    def fx_get_massive_reports(self):

        now = dt.datetime.now()

        filepath = self.get_filenames()
        df = pd.read_excel(filepath[0], dtype=str)
        error_list = []
        for i in df.index:
            temp = df.iloc[i]
            RU = temp[0]
            year = temp[1]
            period = temp[2]
            currency = temp[3]
            lista = [RU, year, period, currency]

            try:
                print("Querying data from:", lista)
                df_snowflake = self.fx_get_data_from_fed(RU, year, period)
                self.fx_get_data(df_snowflake, RU, year, period, currency)

            except:
                error_list.append(lista)


        print("MASSIVE JOB FINISHED")
        if len(error_list) != 0:
            print("Errors on the RUs to follow:")
            for i in error_list:
                print(i)

        later = dt.datetime.now()
        print(now,"+++",later)
        print(f"Elapsed time {later-now}")

    def fx_run_individual_ru(self):
        #sap_instance = self.style.get()
        RU = self.RU.get()
        year =  self.year.get()
        period = f"{'0'+str(self.period.get())}" if int(self.period.get()) < 10 else self.period.get()
        #period_comp = (int(self.period.get()) - 1) if (int(self.period.get()) - 1) != 0 else 1
        currency = self.currency.get()

        print(RU,year,period)

        df = self.fx_get_data_from_fed(RU,year,period)

        self.fx_get_data(df, RU, year, period, currency)
        #self.get_data_from_fed(RU,year,period,currency)

    def fx_get_data_from_fed(self,RU,year,period):


        query = f'''
            SELECT "G/L Acct","Period",SUM("GC Amount") AS Amount
            FROM FIN_CORP_RESTR_PRD_ANALYTICS.JET_CONSUMPTION.VW_JET_WWSL_TOTAL_ACTUAL_ALL 
            WHERE LEFT("G/L Acct",4) = '0095'
                AND "Company Code" = '{RU}' 
                AND "Fiscal Yr" = '{year}'
                AND "Period" <= '{period}'
            OR (LEFT("G/L Acct",4) ='0096'
                AND "Company Code" = '{RU}' 
                AND "Fiscal Yr" = '{year}'
                AND "Period" <= '{period}')                
            GROUP BY "G/L Acct","Period"
            ORDER BY "G/L Acct"

            '''

        df = self.__FED.execute_query(query).fetch_pandas_all()
        df = df[df['G/L Acct'] != '0096959799']
        df = df[df['G/L Acct'] != '0095959799']
        df["AMOUNT"]=df["AMOUNT"].astype(float)
        df["CURRENT_PERIOD"] = df['AMOUNT']
        df["PREVIOUS_PERIOD"] = df[df['Period'].astype(int) < int(period)]['AMOUNT']
        df['CURRENT_MONTH'] = df[df['Period'] == period]['AMOUNT']

        df.groupby('G/L Acct').sum(['CURRENT_PERIOD','PREVIOUS_PERIOD','CURRENT_MONTH'])
        df = df.merge(self.look_up_table,on='G/L Acct',how='left')
        df['Main'].fillna(method='ffill', axis=0, inplace=True)


        return df

    @staticmethod
    def fx_get_data(df_snowflake, RU, year, period, currency):
        if currency.upper() == "USD":

            quote_flag = True

        else:
            try:
                print('Downloading plots')
                month = int(period)
                date1 = dt.datetime(int(year) if month < 12 else int(year) + 1, month + 1 if month < 12 else 1,
                                    1) - dt.timedelta(days=1)
                x = date1.date()

                print(x)
                df = get_dataframe(currency, start=f'{str(int(year) - 1)}-12-31', end=str(x))
                plot_png(df, ticker=currency)
                quote_flag = True if df['Close'][-1] / df['Close'][-21] < 1 else False

                # print('DF ClOSE -1 =', df['Close'][-1])
                # print('DF ClOSE -22 =', df['Close'][-21])
                # print('Division', df['Close'][-1] / df['Close'][-21])
                # print('Quote Flag =', quote_flag)

            except:
                print('Error on downloading plots')


        analyze_data(df=df_snowflake, year=year, period=period, RU=RU,  quote_flag = quote_flag)



        xl, wb, sheet = open_excel_file(f'{os.environ["USERPROFILE"]}\Desktop\ForexAnalysisOutput\Forex_ANALYSIS_{RU}_{year}_{period}.xlsx')
        xl.DisplayAlerts = False
        init_structure(xl, wb, sheet, RU, year, period, currency)

        format_cells(xl,wb)

        wb.SaveAs(f'{os.environ["USERPROFILE"]}\Desktop\ForexAnalysisOutput\Forex_ANALYSIS_{RU}_{year}_{period}.xlsx', ConflictResolution=2)
        wb.Close()
        print(
            f"File successfully saved at {os.environ['USERPROFILE']}\Desktop\ForexAnalysisOutput\Forex_ANALYSIS_{RU}_{year}_{period}.xlsx")


    @staticmethod
    def fx_SAP_get_data(sap_instance, year, period, period_comp,RU,variant, quote_flag, currency):
        if currency.upper() == "USD":

            quote_flag = True

        else:
            try:
                print('Downloading plots')
                month = int(period)
                date1 = dt.datetime(int(year) if month < 12 else int(year) + 1, month + 1 if month < 12 else 1,
                                    1) - dt.timedelta(days=1)
                x = date1.date()

                print(x)
                df = get_dataframe(currency, start=f'{str(int(year) - 1)}-12-31', end=str(x))
                plot_png(df, ticker=currency)
                quote_flag = True if df['Close'][-1] / df['Close'][-22] < 1 else False

            except:
                print('Error on downloading plots')

        get_data_from_sap_and_analyze(sap_instance, year, period, period_comp, RU, variant, quote_flag, currency)

        xl, wb, sheet = open_excel_file(f'{os.environ["USERPROFILE"]}\Desktop\Forex_ANALYSIS_{RU}_{year}_{period}.xlsx')
        xl.DisplayAlerts = False
        init_structure(xl, wb, sheet, RU, year, period)

        wb.SaveAs(f'{os.environ["USERPROFILE"]}\Desktop\Forex_ANALYSIS_{RU}_{year}_{period}.xlsx', ConflictResolution=2)
        wb.Close()
        print(
            f"File successfully saved at {os.environ['USERPROFILE']}\Desktop\Forex_ANALYSIS_{RU}_{year}_{period}.xlsx")










