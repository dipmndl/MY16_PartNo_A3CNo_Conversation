import sys
import pandas as pd

class FormatData:

    # A class method is a method that is bound to a class rather than its object.
    # It doesn't require creation of a class instance, much like staticmethod.
    # This is called a decorator that convert below method as class method
    @classmethod
    def get_display_doors(cls, width, max_col):
        pd.set_option('display.width', width)
        pd.set_option('display.max_columns', max_col)


def read_excel_data():
    try:
        df = pd.read_excel(
            "D:\\2.0 HONDA\\Automation_Dev\\MY16\ECM Files\\1. ECM 10451940\\SAPPDM LIST_BCM_TGJT.xlsm")
        return df

    except Exception as e:
        print("Oops!", sys.exc_info()[0], "occurred.")
        print("Oops!", e.__class__, "occurred in read_excel_data method.")


# Collect column header name based on individual variant
def get_col_name(df, variant):
    try:
        col_name = []
        for i in df:
            if '_' in i:
                s = str(i)
                var = s.split('_')
                if variant in var[2]:
                    col_name.append(i)
            elif 'Object Identifier' in i:
                col_name.append(i)
        return col_name
    except Exception as e:
        print("Oops!", e.__class__, "occurred in get_col_name method.")


# Selecting Type(Factory Or Spare)
def get_df_fact_type(ind_df, fact_type):
    try:
        f_type = 0
        f_type_val = 0
        for f in fact_type:
            if ind_df['ECN_New'].str.contains(f).all():
                f_type = 1
                f_type_val = f
        return f_type, f_type_val
    except Exception as e:
        print("Oops!", e.__class__, "occurred in get_df_fact_type method.")


def get_df_spare_type(ind_df, spare_type):
    try:
        s_type = 0
        s_type_val = 0
        for s in spare_type:
            if ind_df['ECN_New'].str.contains(s).all():
                s_type = 1
                s_type_val = s
        return s_type, s_type_val

    except Exception as e:
        print("Oops!", e.__class__, "occurred in get_df_spare_type method.")


# Selecting variants as per individual dataframe
def list_all_variants(all_names):
    try:
        var = []
        for vr in all_names:
            s = vr.split('-')
            s_lst = list(s)
            s_lst = s_lst[2][:2]
            var.append(s_lst)
        return var
    except Exception as e:
        print("Oops!", e.__class__, "occurred in list_all_variants method.")


def get_variant_name(ind_df, variants):
    try:
        ind_variant = ''
        for v in variants:
            # Coommented as if 'M1' variant is available then code was not working
            # if ind_df['ECN_New'].str.contains(v).all():
            #     ind_variant = v
            s = ind_df['ECN_New'].unique()
            s_split = s[0].split('-')
            if v in s_split[2]:
                # print(s_split[2])
                ind_variant = v
        return ind_variant
    except Exception as e:
        print("Oops!", e.__class__, "occurred in get_variant_name method.")


# check column in dataframe
def validate_column(df, col_name_lst):
    try:
        for col in col_name_lst:
            if col in df.columns:
                df.drop([col], axis=1, inplace=True)
        return df
    except Exception as e:
        print("Oops!", e.__class__, "occurred in validate_column method.")


# check column in dataframe
def find_difference(new_val, old_val, flag):
    try:
        if flag == 'N':                # New value difference check
            p1_lst = list(new_val)
            p2_lst = list(old_val)
            res = [idx for idx, elem in enumerate(p2_lst)
                   if elem != p1_lst[idx]]
            val = ''
            for i in res:
                val += p1_lst[i]
            str1 = ','.join(val)
        elif flag == 'O':                # Old value difference check
            p1_lst = list(new_val)
            p2_lst = list(old_val)
            res = [idx for idx, elem in enumerate(p1_lst)
                   if elem != p2_lst[idx]]
            val = ''
            for i in res:
                val += p2_lst[i]
            str1 = ','.join(val)
        else:
            str1 = ''
        return str1

    except Exception as e:
        print("Oops!", e.__class__, "occurred in validate_column method.")