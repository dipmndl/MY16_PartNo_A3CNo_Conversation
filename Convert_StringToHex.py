import binascii
import pandas as pd
import datetime as dt
import sys
import pandas as pd
import Formating_Report as fr

pd.set_option('display.width', 1400)
pd.set_option('display.max_columns', 1000)
pd.set_option('display.max_rows', 1000)

#ecn_data = pd.read_excel("D:\\2.0 HONDA\\1.0 Project_Docs\\MY21\\1. HomeWork\\Result Review\\ECN 10434277_T43A.xlsx",index_col="Sr No.")
#ecn_data = pd.read_excel("D:\\2.0 HONDA\\Automation_Dev\\MY16\ECM Files\\1. ECM 10451940""\\ECM 10456526 3A0A MP.xlsx", index_col="Sr No.")
#ecn_data.sort_index(inplace=True)
#formation_date = str(dt.datetime.today().strftime('%d-%m-%Y'))

try:
    display_obj = fr.FormatData.get_display_doors(3000, 1000)
    df_A3C_No = fr.read_excel_data()
    df_A3C_No.rename({"Unnamed: 0": "a"}, axis="columns", inplace=True)
    df_A3C_No.drop(["a"], axis=1, inplace=True)
    df_A3C_No.fillna(0, inplace=True)
    new_header = df_A3C_No.iloc[9]  # grab the first row for the header(iloc by index position --> SAPPDM excel row no = 11 --> 'High / SSP / Low'). If Actual excel file format change, Change this index position accordingly
    df_A3C_No = df_A3C_No[10:]  # take the data less the header row(iloc by index position --> SAPPDM excel row no = 12 -->  BCM Variant)
    df_A3C_No.columns = new_header
    df_A3C_No = df_A3C_No.loc[df_A3C_No['High / SSP / Low'].isin(['BCM Variant', 'Customer Part Number', 'Material-No.'])]
    df_A3C_No = df_A3C_No.loc[:, (df_A3C_No != 0).any(axis=0)]  # remove columns having 0 value.
    df_A3C_No.drop('Smart', axis=1, inplace=True)
    df_A3C_No.columns.name = None  # remove index from header column
    df_A3C_No = df_A3C_No.set_index('High / SSP / Low')
    df_A3C_No_transposed = df_A3C_No.T

except Exception as e:
    print("Oops!", sys.exc_info()[0], "occurred.")
    print("Oops!", e.__class__, "occurred in A3C No file.")
    print(e)