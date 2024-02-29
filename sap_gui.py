import win32com.client
import subprocess

import time
import sys
import os


class SapGui(object):
    """
    Class to instanciate a SAP session
    It can return the session and then the user can access SAP scripting
    """

    def __doc__(self):
        pass

    def __init__(self, sap_instance):
        sap_instance = sap_instance
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        print('Building object {}'.format(self))

        time.sleep(3)
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = self.SapGuiAuto.GetScriptingEngine

        if application.Connections.Count > 4:
            print("There are too many SAP sessions.Impossible to continue")

        else:
            if application.Connections.Count > 0:
                flag = False
                for i in range((application.Connections.Count)):
                    conn = application.Children(int(i))
                    desc = conn.Description
                    if desc == sap_instance:
                        flag = True
                        self.connection = conn
                        break
                if flag == False:
                    self.connection = application.OpenConnection(sap_instance, True, True)

            else:
                self.connection = application.OpenConnection(sap_instance, True, True)

            # NOME DA CONEXAO DO SAP#

            # time.sleep(1)
            self.session = self.connection.Children(0)
            print('Session ID:{}'.format(self.session.Id))
            print('Session Parent:{}'.format(self.session.Parent))
            self.session.findById('wnd[0]').maximize()

            self.user = str(os.environ['USERPROFILE'] + r'\Desktop')

    def sapLogin(self):
        try:
            self.session.findById('wnd[0]').sendVKey(0)
        except:
            print(sys.exc_info()[0])

    @property
    def get_session(self):
        "returns the sap session"
        return self.session

    @staticmethod
    def format_negative_numbers(df):
        df.replace(',', '', inplace=True)
        for i in range(len(df.columns)):
            for j in range(len(df.index)):
                if ',' in df.iloc[j, i]:
                    chars = []
                    for c in df.iloc[j, i]:
                        if c == ',':
                            pass
                        else:
                            chars.append(c)
                    df.iloc[j, i] = ''.join(chars)

                if df.iloc[j, i][-1] == '-':
                    df.iloc[j, i] = float(str(df.iloc[j, i][:-1]).strip()) * -1
                else:

                    df.iloc[j, i] = float(str(df.iloc[j, i]).strip())
        return df


