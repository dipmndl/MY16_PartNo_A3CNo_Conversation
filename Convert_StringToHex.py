import binascii
import pandas as pd
import datetime as dt
import sys
import Formating_Report as fr

pd.set_option('display.width', 1400)
pd.set_option('display.max_columns', 1000)
pd.set_option('display.max_rows', 1000)

#ecn_data = pd.read_excel("D:\\2.0 HONDA\\1.0 Project_Docs\\MY21\\1. HomeWork\\Result Review\\ECN 10434277_T43A.xlsx",index_col="Sr No.")
#ecn_data = pd.read_excel("D:\\2.0 HONDA\\Automation_Dev\\MY16\ECM Files\\1. ECM 10451940""\\ECM 10456526 3A0A MP.xlsx", index_col="Sr No.")
#ecn_data.sort_index(inplace=True)
#formation_date = str(dt.datetime.today().strftime('%d-%m-%Y'))

try:
    '''
    display_obj = fr.FormatData.get_display_doors(3000, 1000)
    df = fr.read_excel_data()
    df.rename({"Unnamed: 0": "a"}, axis="columns", inplace=True)
    df.drop(["a"], axis=1, inplace=True)
    df.fillna(0, inplace=True)
    options = ['Customer Part Number'] # if require add 'Material-No.'
    # rslt_df = df[df['SAP Customer Project Name: HONMY16IBC'].isin(options)]

    new_header = df.iloc[5] # select 'High / SSP / Low'
    # rslt_df = df[11:13] # select 'Customer Part Number' by index positions
    rslt_df_name  = df.loc[df['SAP Customer Project Name: HONMY16IBC'].isin(options)]  # select 'Customer Part Number' by name
    rslt_df_name.columns = new_header
    # rslt_df_name = rslt_df_name.set_index('High / SSP / Low')
    rslt_df_name = rslt_df_name.loc[:, (rslt_df_name != 0).any(axis=0)]  # remove columns having 0 value.
    rslt_df_name_transpose = rslt_df_name.T
    # Factory_df = rslt_df_name_transpose.loc[rslt_df_name_transpose['High / SSP / Low'] == 'High']
    # Factory_df = rslt_df_name['High']
    # Factory_df_transpose = rslt_df_name.T
    # Spare_df  = rslt_df_name['SSP High']

    excel_file = pd.ExcelWriter("D:\\MY16\ECM Files\\ECM_result.xlsx")
    rslt_df_name_transpose.to_excel(excel_file, sheet_name="Result", index=True)
    excel_file.save()
     df_new = pd.read_excel("D:\\MY16\ECM Files\\ECM_result.xlsx", sheet_name="Result")
    new_header = df_new.iloc[0]  # grab the first row for the header
    df_new = df_new[1:]  # take the data less the header row
    df_new.columns = new_header  # set the header row as the df header
    # Factory
    Factory_df = df_new.loc[df_new['High / SSP / Low'] == 'High']
    Factory_df = Factory_df.loc[2:, ['Customer Part Number']]
    Factory_df.reset_index(drop=True, inplace=True)
    excel_file = pd.ExcelWriter("D:\\MY16\ECM Files\\Factory_ECM.xlsx")
    Factory_df.to_excel(excel_file, sheet_name="Result", index=True)
    excel_file.save()
    # Spare
    Spare_df = df_new.loc[df_new['High / SSP / Low'].isin(['SSP High'])]
    Spare_df = Spare_df.loc[2:, ['Customer Part Number']]
    Spare_df.reset_index(drop=True, inplace=True)
    excel_file = pd.ExcelWriter("D:\\MY16\ECM Files\\Spare_ECM.xlsx")
    Spare_df.to_excel(excel_file, sheet_name="Result", index=True)
    excel_file.save()
    print(df_new)
    '''
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

