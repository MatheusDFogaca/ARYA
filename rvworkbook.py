import os
import win32com.client as win32
from pandas.tseries.offsets import BMonthEnd
import pandas as pd
from datetime import datetime
import datetime as dt

def open_rv_file(rv_file_name):

        # Openning new excel application

        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        wb = xl.Workbooks.Open(rv_file_name)
        xl.ActiveWindow.DisplayGridlines = False
        return xl, wb

def create_workbook(df_rv, df_db, RU, year, period, system, currency):

    # Creating workbook and manipulating about rates (DSTDT) and account values (gd13)

    rv_file_name = f'Forex_Revaluation_{RU}_{year}_{period}.xlsx'

    writer = pd.ExcelWriter(os.environ["USERPROFILE"] + f"\Desktop\ForexRevaluationOutput\\{rv_file_name}", engine="xlsxwriter")

    if system != 'G9P' and currency == 'USD':
        df_rv.to_excel(writer, sheet_name='FED data')

        df_generic_for_usd = pd.DataFrame(
            {'Client': [100], 'Target Sys': ['Any'], 'Ref. Date': ['any'], 'From Curr.': ['USD'],
             'To Curr.': ['USD'], 'Tgt RateTy': ['usdxusd'], 'Exchg rate': [1], 'Numerator': [1],
             'Denomintr.': [1], 'Valid. Dt.': ['everyday']})
        df_dv = pd.DataFrame([""])
        df_generic_for_usd.to_excel(writer, sheet_name='Currencies', index=False)
        df_rv.to_excel(writer, sheet_name='FED data')
        df_dv.to_excel(writer, sheet_name='Data Visualization', index=False)
        df_db.to_excel(writer, sheet_name='Distribution Check')

        writer.close()

    else:
        df = pd.read_csv(f'{os.environ["USERPROFILE"]}\Desktop\ForexRevaluationOutput\Currencies.txt', encoding="utf-8", usecols=lambda c: not c.startswith('Unnamed:'),
                          skiprows=3, sep='\t')

        # Dropping different systems values
        if system == 'AMP':
            df.drop(df[df['Target Sys'] != 'AMX'].index, inplace=True)
        elif system == 'G9P':
            df.drop(df[df['Target Sys'] != 'G9X'].index, inplace=True)
        elif system == 'G3P':
            df.drop(df[df['Target Sys'] != '3GX'].index, inplace=True)
        elif system == 'S8B':
            df.drop(df[df['Target Sys'] != 'S8X'].index, inplace=True)
        elif system == 'EUP':
            df.drop(df[df['Target Sys'] != 'ESX'].index, inplace=True)
        elif system == 'APP':
            df.drop(df[df['Target Sys'] != 'PAX'].index, inplace=True)
        elif system == 'F2P':
            df.drop(df[df['Target Sys'] != 'F2X'].index, inplace=True)

        # Deleting other columns

        df = df.drop(
            ['Dist. Date', 'Dist. Grp.', 'Item num', 'Ref. Item', 'Rate Ind.', '      Outb. IDOC', 'Indir Quot',
             '       Inb. IDOC', 'Load Ind.', 'Created by', 'Date', 'Time'], axis=1)

        df_dv = pd.DataFrame([""])
        df_rv.to_excel(writer, sheet_name='FED data')
        df.to_excel(writer, sheet_name='Currencies', index=False)
        df_dv.to_excel(writer, sheet_name='Data Visualization', index=False)
        df_db.to_excel(writer, sheet_name='Distribution Check')

        writer.close()

def fed_data(wb, currency):

    # Calculating currencies deviation in FED Data tab

    sheet = wb.Sheets('FED Data')
    sheet.Columns(1).Delete()
    sheet.Cells(1, 9).Value = 'DC Exchange Rate'
    sheet.Cells(1, 10).Value = 'LC Exchange Rate'
    sheet.Cells(1, 11).Value = 'DeltaLC-GC'
    sheet.Cells(1, 12).Value = 'DeltaGC-LC'
    sheet.Cells(1, 13).Value = 'DeltaDC-GC'
    sheet.Cells(1, 14).Value = 'DeltaGC-DC'
    sheet.Cells(1, 15).Value = 'Base Account'
    sheet.Cells(1, 16).Value = 'Closing Rate - DC'
    sheet.Cells(1, 17).Value = 'Closing Rate - LC'
    maxr = sheet.UsedRange.Rows.Count

    for i in range(2, maxr + 1):
        sheet.Cells(i, 9).Value = '=IFERROR(ABS(INDIRECT("RC[-3]",FALSE)/INDIRECT("RC[-1]",FALSE)),)'
        sheet.Cells(i, 10).Value = '=IFERROR(ABS(INDIRECT("RC[-3]",FALSE)/INDIRECT("RC[-2]",FALSE)),)'
        sheet.Cells(i, 11).Value = '=IFERROR(IF(INDIRECT("RC[-3]",FALSE) = "", "", (INDIRECT("RC[-3]",FALSE)*(INDIRECT("RC[6]",FALSE)))-INDIRECT("RC[-4]",FALSE)), "")'
        sheet.Cells(i, 12).Value = '=IFERROR(IF(INDIRECT("RC[-5]",FALSE) = "", "", (INDIRECT("RC[-5]",FALSE)/(INDIRECT("RC[5]",FALSE)))-INDIRECT("RC[-4]",FALSE)), "")'
        sheet.Cells(i, 13).Value = '=IFERROR((INDIRECT("RC[-5]",FALSE)*(INDIRECT("RC[3]",FALSE)))-INDIRECT("RC[-7]",FALSE), 0)'
        sheet.Cells(i, 14).Value = '=IFERROR((INDIRECT("RC[-8]",FALSE)/(INDIRECT("RC[2]",FALSE)))-INDIRECT("RC[-6]",FALSE), 0)'
        sheet.Cells(i, 15).Value = '=LEFT((INDIRECT("RC[-13]",FALSE)),6)'
        sheet.Cells(i, 16).Value = '=IFERROR(IFERROR(IFERROR(IF(OR((INDIRECT("RC[-13]",FALSE))="USD",AND((INDIRECT("RC[-13]",FALSE))="",(INDIRECT("RC[-12]",FALSE))="USD")),1,VLOOKUP((INDIRECT("RC[-13]",FALSE)),Currencies!D:I,4,0)*VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!D:I,5,0)/VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!D:I,6,0)),VLOOKUP((INDIRECT("RC[-12]",FALSE)),Currencies!D:I,4,0)*VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!D:I,5,0)/VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!D:I,6,0)),VLOOKUP((INDIRECT("RC[-13]",FALSE)),Currencies!E:I,3,0)*VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!E:I,4,0)/VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!E:I,5,0)), VLOOKUP((INDIRECT("RC[-12]",FALSE)),Currencies!E:I,3,0)*VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!E:I,4,0)/VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!E:I,5,0))'
        if currency in ["EGP", "DZD"]:
            sheet.Cells(i,
                        17).Value = '=IFERROR(IFERROR(IFERROR(IF(OR((INDIRECT("RC[-13]",FALSE))="USD",(INDIRECT("RC[-13]",FALSE))=""),1,VLOOKUP((INDIRECT("RC[-13]",FALSE)),Currencies!E:I,3,0)*VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!D:I,5,0)/VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!D:I,6,0)),VLOOKUP((INDIRECT("RC[-12]",FALSE)),Currencies!D:I,4,0)*VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!D:I,5,0)/VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!D:I,6,0)),VLOOKUP((INDIRECT("RC[-13]",FALSE)),Currencies!E:I,3,0)*VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!E:I,4,0)/VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!E:I,5,0)), VLOOKUP((INDIRECT("RC[-12]",FALSE)),Currencies!E:I,3,0)*VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!E:I,4,0)/VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!E:I,5,0))'
        else:
            sheet.Cells(i,
                        17).Value = '=IFERROR(IFERROR(IFERROR(IF(OR((INDIRECT("RC[-13]",FALSE))="USD",(INDIRECT("RC[-13]",FALSE))=""),1,VLOOKUP((INDIRECT("RC[-13]",FALSE)),Currencies!D:I,4,0)*VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!D:I,5,0)/VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!D:I,6,0)),VLOOKUP((INDIRECT("RC[-12]",FALSE)),Currencies!D:I,4,0)*VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!D:I,5,0)/VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!D:I,6,0)),VLOOKUP((INDIRECT("RC[-13]",FALSE)),Currencies!E:I,3,0)*VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!E:I,4,0)/VLOOKUP(INDIRECT("RC[-13]",FALSE),Currencies!E:I,5,0)), VLOOKUP((INDIRECT("RC[-12]",FALSE)),Currencies!E:I,3,0)*VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!E:I,4,0)/VLOOKUP(INDIRECT("RC[-12]",FALSE),Currencies!E:I,5,0))'
    sheet.Columns("F:H").NumberFormat = "#,##0.00"
    sheet.Columns("I:J").NumberFormat = "#,##0.00000"
    sheet.Columns("K:N").NumberFormat = "#,##0.00"
    sheet.Columns.AutoFit()

def power_pivot(RU, year, period, system, currency, xl, wb):

    # Creating a power pivot with FED Data tab into Data Visualization

    def insert_pt_field_DC(pt):
        field_rows = {}
        field_rows['RU'] = pt.PivotFields('RU')
        field_rows['Base Account'] = pt.PivotFields('Base Account')
        field_rows['DCurr'] = pt.PivotFields('DCurr')
        field_rows['Account'] = pt.PivotFields('Account')

        field_values = {}
        field_values['DC Amount'] = pt.PivotFields('DC Amount')
        field_values['LC Amount'] = pt.PivotFields('LC Amount')
        field_values['GC Amount'] = pt.PivotFields('GC Amount')
        field_values['DeltaDC-GC'] = pt.PivotFields('DeltaDC-GC')
        field_values['DeltaGC-DC'] = pt.PivotFields('DeltaGC-DC')
        field_values['DeltaLC-GC'] = pt.PivotFields('DeltaLC-GC')
        field_values['DeltaGC-LC'] = pt.PivotFields('DeltaGC-LC')

        field_rows['Account'].Orientation = 1
        field_rows['Account'].Position = 1

        field_rows['DCurr'].Orientation = 1
        field_rows['DCurr'].Position = 1

        field_rows['Base Account'].Orientation = 1
        field_rows['Base Account'].Position = 1

        field_rows['RU'].Orientation = 1
        field_rows['RU'].Position = 1

        field_values['DC Amount'].Orientation = 4
        field_values['DC Amount'].Function = -4157
        field_values['DC Amount'].NumberFormat = "#,##0.00"

        field_values['LC Amount'].Orientation = 4
        field_values['LC Amount'].Function = -4157
        field_values['LC Amount'].NumberFormat = "#,##0.00"

        field_values['GC Amount'].Orientation = 4
        field_values['GC Amount'].Function = -4157
        field_values['GC Amount'].NumberFormat = "#,##0.00"

        field_values['DeltaDC-GC'].Orientation = 4
        field_values['DeltaDC-GC'].Function = -4157
        field_values['DeltaDC-GC'].NumberFormat = "#,##0.00"

        field_values['DeltaGC-DC'].Orientation = 4
        field_values['DeltaGC-DC'].Function = -4157
        field_values['DeltaGC-DC'].NumberFormat = "#,##0.00"

        field_values['DeltaLC-GC'].Orientation = 4
        field_values['DeltaLC-GC'].Function = -4157
        field_values['DeltaLC-GC'].NumberFormat = "#,##0.00"

        field_values['DeltaGC-LC'].Orientation = 4
        field_values['DeltaGC-LC'].Function = -4157
        field_values['DeltaGC-LC'].NumberFormat = "#,##0.00"

    def insert_pt_field_LC(pt):
        field_rows = {}
        field_rows['RU'] = pt.PivotFields('RU')
        field_rows['Base Account'] = pt.PivotFields('Base Account')
        field_rows['LCurr'] = pt.PivotFields('LCurr')
        field_rows['Account'] = pt.PivotFields('Account')

        field_values = {}
        field_values['DC Amount'] = pt.PivotFields('DC Amount')
        field_values['LC Amount'] = pt.PivotFields('LC Amount')
        field_values['GC Amount'] = pt.PivotFields('GC Amount')
        field_values['DeltaDC-GC'] = pt.PivotFields('DeltaDC-GC')
        field_values['DeltaGC-DC'] = pt.PivotFields('DeltaGC-DC')
        field_values['DeltaLC-GC'] = pt.PivotFields('DeltaLC-GC')
        field_values['DeltaGC-LC'] = pt.PivotFields('DeltaGC-LC')

        field_rows['Account'].Orientation = 1
        field_rows['Account'].Position = 1

        field_rows['LCurr'].Orientation = 1
        field_rows['LCurr'].Position = 1

        field_rows['Base Account'].Orientation = 1
        field_rows['Base Account'].Position = 1

        field_rows['RU'].Orientation = 1
        field_rows['RU'].Position = 1

        field_values['DC Amount'].Orientation = 4
        field_values['DC Amount'].Function = -4157
        field_values['DC Amount'].NumberFormat = "#,##0.00"

        field_values['LC Amount'].Orientation = 4
        field_values['LC Amount'].Function = -4157
        field_values['LC Amount'].NumberFormat = "#,##0.00"

        field_values['GC Amount'].Orientation = 4
        field_values['GC Amount'].Function = -4157
        field_values['GC Amount'].NumberFormat = "#,##0.00"

        field_values['DeltaDC-GC'].Orientation = 4
        field_values['DeltaDC-GC'].Function = -4157
        field_values['DeltaDC-GC'].NumberFormat = "#,##0.00"

        field_values['DeltaGC-DC'].Orientation = 4
        field_values['DeltaGC-DC'].Function = -4157
        field_values['DeltaGC-DC'].NumberFormat = "#,##0.00"

        field_values['DeltaLC-GC'].Orientation = 4
        field_values['DeltaLC-GC'].Function = -4157
        field_values['DeltaLC-GC'].NumberFormat = "#,##0.00"

        field_values['DeltaGC-LC'].Orientation = 4
        field_values['DeltaGC-LC'].Function = -4157
        field_values['DeltaGC-LC'].NumberFormat = "#,##0.00"

    ws_data = wb.Worksheets('FED Data')
    sheet = wb.Worksheets('Data Visualization')
    pt_cache = wb.PivotCaches().Create(1, SourceData=ws_data.Range("A:P"))

    # create pivot table
    pt = pt_cache.CreatePivotTable(sheet.Range("A3"), "MyReport")

    #toggle grand totals
    pt.ColumnGrand = False
    pt.RowGrand = False

    #pivot table styles
    pt.TableStyle2 = "PivotStyleMedium9"

    DC_systems = ['G9P', 'S8P']

    if system in DC_systems:
        insert_pt_field_DC(pt)
    else:
        insert_pt_field_LC(pt)

    last = sheet.UsedRange.rows.Count
    sheet.Cells(1, 5).Value = 'Number of Required Corrections:'
    sheet.Cells(1, 6).Value = '=COUNTIF(J:J, "Yes")'
    sheet.Cells(1, 7).Value = 'Closing Rate'
    sheet.Cells(1, 8).Value = 'Required?'
    sheet.Cells(1, 3).Value = f'Arya for RU {RU}'
    sheet.Cells(1, 7).Value = 'File Data'
    sheet.Cells(1, 8).Value = datetime.today().strftime('%m/%d/%Y')
    sheet.Cells(2, 7).Value = 'File Time'
    sheet.Cells(2, 8).Value = datetime.today().strftime('%H:%M:%S')
    for i in range (6, last):
        sheet.Cells(i, 9).Value = """=IFERROR(IF(AND((INDIRECT("RC[-8]",FALSE))="USD",(INDIRECT("RC[-7]",FALSE))=(INDIRECT("RC[-6]",FALSE))),1, IF(AND(LEN(INDIRECT("RC[-8]",FALSE))=3, INDIRECT("RC[-8]",FALSE)<>"USD"), INDIRECT("RC[-7]",FALSE)/INDIRECT("RC[-5]",FALSE), IF(AND((INDIRECT("RC[-8]",FALSE))="USD",(INDIRECT("RC[-7]",FALSE))<>(INDIRECT("RC[-6]",FALSE))), INDIRECT("RC[-6]",FALSE)/INDIRECT("RC[-5]",FALSE), ""))),"")"""
        sheet.Cells(i, 10).Value = '=IF(LEN(INDIRECT("RC[-9]",FALSE))=3,IF(AND(INDIRECT("RC[-8]",FALSE)=0, INDIRECT("RC[-7]",FALSE)=0), "No", IF(AND((INDIRECT("RC[-8]",FALSE))="USD",(INDIRECT("RC[-8]",FALSE))<>0, (INDIRECT("RC[-7]",FALSE))=(INDIRECT("RC[-6]",FALSE))),IF(OR(INDIRECT("RC[-4]",FALSE)>=1000000,INDIRECT("RC[-4]",FALSE)<=-1000000),"Yes","No"),IF(OR(INDIRECT("RC[-2]",FALSE)>=1000000,INDIRECT("RC[-2]",FALSE)<=-1000000),"Yes","No"))),"")'

    sheet.Columns.AutoFit()
    sheet.Select()
    xl.ActiveWindow.DisplayGridlines = False

def distribution_check(wb):

    # Special check to see if the distribution happened as expected

    sheet = wb.Sheets("Distribution Check")
    wb.Worksheets('Distribution Check').Activate()

    maxr = sheet.UsedRange.Rows.Count

    for i in range(2, maxr + 1):
        sheet.Cells(i, 3).NumberFormat = "#,##0.00"
        sheet.Cells(i, 4).NumberFormat = "#,##0.00"
        sheet.Cells(i, 5).NumberFormat = "#,##0.00"


def rv_excel_formating(RU, year, period, wb, xl):

    # Last formatting part

    sheet = wb.Worksheets("Data Visualization")
    offset = BMonthEnd()
    print(period)
    x = dt.date(int(year), int(period), 1)
    # ld = offset.rollback(x)

    sheet.Cells(1, 3).Value = f'Arya for RU {RU}'
    sheet.Cells(1, 7).Value = 'File Data'
    sheet.Cells(1, 8).Value = datetime.today().strftime('%m/%d/%Y')
    sheet.Cells(2, 7).Value = 'File Time'
    sheet.Cells(2, 8).Value = datetime.today().strftime('%H:%M:%S')


    def rgb_to_hex(rgb):
        '''
        sheet.Cells(1, i).Interior.color uses bgr in hex

        '''
        bgr = (rgb[2], rgb[1], rgb[0])
        strValue = '%02x%02x%02x' % bgr
        # print(strValue)
        iValue = int(strValue, 16)
        return iValue

    last = sheet.UsedRange.rows.Count

    sheet.Cells(1, 6).Interior.Color = rgb_to_hex((230, 184, 183))
    for i in range(6, last):
        if sheet.Cells(i, 10).Value == "Yes":
            sheet.Cells(i, 10).Interior.Color = rgb_to_hex((230, 184, 183))
    sheet.Cells(3, 7).Interior.Color = rgb_to_hex((79, 129, 189))
    sheet.Cells(3, 8).Interior.Color = rgb_to_hex((79, 129, 189))
    sheet.Cells(4, 7).Interior.Color = rgb_to_hex((79, 129, 189))
    sheet.Cells(4, 8).Interior.Color = rgb_to_hex((79, 129, 189))
    sheet.Cells(5, 7).Interior.Color = rgb_to_hex((220, 230, 241))
    sheet.Cells(5, 8).Interior.Color = rgb_to_hex((220, 230, 241))

    sheet.Cells(4, 7).Font.Color = rgb_to_hex((255, 255, 255))
    sheet.Cells(4, 8).Font.Color = rgb_to_hex((255, 255, 255))

    sheet.Cells(1, 9).Font.Bold = True
    sheet.Cells(1, 10).Font.Bold = True
    sheet.Cells(1, 11).Font.Bold = True
    sheet.Cells(1, 12).Font.Bold = True
    sheet.Cells(1, 13).Font.Bold = True
    sheet.Cells(1, 14).Font.Bold = True
    sheet.Cells(1, 15).Font.Bold = True
    sheet.Cells(1, 16).Font.Bold = True

    sheet.Cells(4, 7).Font.Bold = True
    sheet.Cells(4, 8).Font.Bold = True
    sheet.Cells(1, 3).Font.Bold = True
    sheet.Cells(2, 3).Font.Bold = True
    sheet.Cells(2, 4).Font.Bold = True
    sheet.Cells(1, 5).Font.Bold = True
    sheet.Cells(1, 6).Font.Bold = True
    sheet.Cells(1, 7).Font.Bold = True
    sheet.Cells(2, 7).Font.Bold = True
    sheet.Cells(1, 8).Font.Bold = True
    sheet.Cells(2, 8).Font.Bold = True

    sheet = wb.Worksheets("Currencies")
    sheet.Columns.AutoFit()
    sheet.Select()
    xl.ActiveWindow.DisplayGridlines = False
    sheet = wb.Worksheets("Distribution Check")
    sheet.Columns.AutoFit()
    sheet.Select()
    xl.ActiveWindow.DisplayGridlines = False



