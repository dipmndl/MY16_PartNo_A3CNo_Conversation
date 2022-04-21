import sys

import pandas as pd
import Formatting_Report_old as fr
import Formatting_Report_old as fr
''' Collect variant wise A3C no from SAPPDM file  
'''
factory_A3C_No = ''
spare_A3C_No = ''
def create_dict_A3C_No(ind_var, df_type):

    try:
        display_obj = fr.FormatData.get_display_doors(3000, 1000)
        df_A3C_No = fr.read_excel_data()
        df_A3C_No.rename({"Unnamed: 0": "a"}, axis="columns", inplace=True)
        df_A3C_No.drop(["a"], axis=1, inplace=True)
        df_A3C_No.fillna(0, inplace=True)
        new_header = df_A3C_No.iloc[9] # grab the first row for the header(iloc by index position --> SAPPDM excel row no = 11 --> 'High / SSP / Low'). If Actual excel file format change, Change this index position accordingly
        df_A3C_No = df_A3C_No[10:]              # take the data less the header row(iloc by index position --> SAPPDM excel row no = 12 -->  BCM Variant)
        df_A3C_No.columns = new_header
        df_A3C_No = df_A3C_No.loc[df_A3C_No['High / SSP / Low'].isin(['BCM Variant', 'Customer Part Number', 'Material-No.'])]
        df_A3C_No = df_A3C_No.loc[:, (df_A3C_No != 0).any(axis=0)]      # remove columns having 0 value.
        df_A3C_No.drop('Smart', axis=1, inplace=True)
        df_A3C_No.columns.name = None           # remove index from header column
        df_A3C_No = df_A3C_No.set_index('High / SSP / Low')
        df_A3C_No_transposed = df_A3C_No.T
        # df_A3C_No_dict =  df_A3C_No_transposed.to_dict('list')
        result = {}
        for lst in df_A3C_No_transposed.values:
            leaf = result
            for path in lst[:-2]:
                leaf = leaf.setdefault(path, {})
            leaf.setdefault(lst[-2], list()).append(lst[-1])
        spare_A3C_No = get_A3C_No(dict_result= result, variant= ind_var, type= df_type)
        return spare_A3C_No

    except Exception as e:
        print("Oops!", sys.exc_info()[0], "occurred.")
        print("Oops!", e.__class__, "occurred in A3C No file.")
        print(e)

def get_A3C_No(dict_result, variant, type):
    try:
        # Get A3C No and assign to both Factory and Spare variable
        spare_A3C_No  = ''
        for k, v in dict_result.items():
            if variant in k:
                # print(k,v)
                for i, j in v.items():
                    if type in i:
                        spare_A3C_No = j
                        return spare_A3C_No
                    # else:
                    #     factory_A3C_No = 'NA'
                    #     return  factory_A3C_No

    except Exception as e:
        print("Oops!", sys.exc_info()[0], "occurred.")
        print("Oops!", e.__class__, "occurred in A3C No file.")
        print(e)