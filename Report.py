import Formatting_Report_old as Fr
import pandas as pd
import A3C_No as An
import StringToHex as sh


def duplicate_rows(df, countcol):
    try:
        for _, row in df.iterrows():
            for i in range(int(row[countcol]) - 1):
                # Append this row at the end of the DataFrame
                df = df.append(row)
        # optional: sort it by index
        df.sort_index(inplace=True)
        return df
    except Exception as e:
        print("Oops!", e.__class__, "occurred in Report file duplicate_rows method.")
        print(e)


try:
    display_obj = Fr.FormatData.get_display_doors(3000, 1000)
    # df = pd.read_excel("D:\\2.0 HONDA\\Automation_Dev\\PartNo_A3CNo_Conversation"
    #                    "\\Converted_Data\\DOORS_data_new_12-03-2021.xlsx", sheet_name="DOORS_NEW_VALUE")
    df = pd.read_excel("D:\\2.0 HONDA\Automation_Dev\\PartNo_A3CNo_Conversation\\Converted_Data\\57. ECM 10456526"
                       "\\3A0A MP data_new_13-04-2022.xlsx", sheet_name="DOORS_NEW_VALUE")
    # df_old_doors = pd.read_excel("D:\\2.0 HONDA\\Automation_Dev\\PartNo_A3CNo_Conversation\\Converted_Data\\ECM 10446975\\2. 3A0A"
    #                              "\\CONF_30AA_Rel_7.0_MP_BeforeUpdate.xlsx")

    df_old_doors = pd.read_excel("D:\\2.0 HONDA\Automation_Dev\\PartNo_A3CNo_Conversation\\Converted_Data\\57. ECM 10456526"
                                 "\\CONF_3A0A_Rel_5.5_MP_AfterUpdate.xlsx")

    df.drop('ECN_NEW', axis=1, inplace=True)
    df.rename(columns={'Sr No.': 'SrNo.', 'ECN': 'ECN_New'}, inplace=True)
    # insert columns based on position
    df.insert(2, column='A3C_No', value='')
    df.insert(3, column='ID', value='')
    df.insert(4, column='Variants', value='')
    df.insert(7, column='Count', value='4')
    df_dup = duplicate_rows(df, 'Count')
    ind_df = pd.DataFrame()
    compl_df = pd.DataFrame()
    ind_variant = ''
    # TODO:- Get all the variants dynamically
    all_names = df_dup['ECN_New'].unique()
    variants = Fr.list_all_variants(all_names)
    # variants = ['H0', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7']
    # TODO:- Create a loop and assign variant wise data
    # if 'c_Release' in df_old_doors.columns:
    #     df_old_doors.drop(
    #         ['c_Release'], axis=1, inplace=True)
    col_name = ['c_Release', 'System specification for Module Configuration', 'p_Customer_Definition', 'p_Range', 'p_Description']
    df_old_doors = Fr.validate_column(df_old_doors, col_name)

    # df_old_doors.drop(['c_Release', 'System specification for Module Configuration', 'p_Customer_Definition', 'p_Range', 'p_Description'], axis=1, inplace=True)
    # print(df_old_doors)
    df_columns = df_old_doors.columns.get_values()
    # TODO:- Create a loop and pass one variant at a time while creation DOORS Old data values.
    id_range = df_dup['SrNo.'].unique()  # Contain all the variants in a dataframe
    # id_range = [1, 2, 3]
    fact_type = ['38800']
    spare_type = ['38809']

    for r in id_range:
        df_type = ''
        ind_df = df_dup.loc[df_dup['SrNo.'] == r]
        type_fact, fact_val = Fr.get_df_fact_type(ind_df, fact_type)
        type_spare, spare_val = Fr.get_df_spare_type(ind_df, spare_type)
        ind_variant = Fr.get_variant_name(ind_df, variants)
        # Select type
        if type_fact != 0:
            df_type = fact_val
        else:
            df_type = spare_val

        col_name_lst = Fr.get_col_name(df_columns, ind_variant)
        var_name = col_name_lst[1]

        # TODO:- New Data Start
        # Declare a list that is to be converted into a column
        Temp_Id = [1, 2, 3, 4]
        # ind_df['Temp_Id'] = Temp_Id                 # To Avoid SettingWithCopyWarning error add below code
        ind_df.insert(8, column='Temp_Id', value=Temp_Id)
        i = 0
        # j = 1
        # Creation Of New Data
        for i in range(len(ind_df['Temp_Id'])):
            if ind_df['Temp_Id'].values[i] == 1:  # 1st index position
                ind_df['ID'].values[i] = 'SYS_CONFIG_582'
                A3C_No_val = An.create_dict_A3C_No(ind_variant, df_type)
                ind_df['A3C_No'].values[i] = A3C_No_val[0]  # Assign: A3C No dynamically
                var_A3CNo = ind_df['A3C_No'].values[i]
                var_A3CNo = var_A3CNo[0:11]  # A3C no always 11 digit
                ind_df['DOORS_NEW_DATA'].values[i] = var_A3CNo
                ind_df['DOORS_NEW_DATA_877'].values[i] = ''
            elif ind_df['Temp_Id'].values[i] == 2:  # 2nd index position
                ind_df['ID'].values[i] = 'SYS_CONFIG_608'
                ind_df['DOORS_NEW_DATA_877'].values[i] = ''
            elif ind_df['Temp_Id'].values[i] == 3:  # 3rd index position
                ind_df['ID'].values[i] = 'SYS_CONFIG_636'
                ind_df['DOORS_NEW_DATA_877'].values[i] = ''
            else:  # 4th index position
                ind_df['ID'].values[i] = 'SYS_CONFIG_877'
                ind_df['DOORS_NEW_DATA'].values[i] = ind_df['DOORS_NEW_DATA_877'].values[i]
                ind_df['DOORS_NEW_DATA_877'].values[i] = ''

        # To Avoid SettingWithCopyWarning error for dropping column change below commented code
        # ind_df.drop('Temp_Id', axis=1, inplace=True)
        # ind_df.insert(9, column='DOORS_OLD', value='')
        ind_df = ind_df.copy()
        ind_df.rename(columns={'Temp_Id': 'DOORS_OLD'}, inplace=True)
        ind_df['DOORS_OLD'] = ''  # To Avoid value error while assigning value empty the column.
        # ind_df.drop(['Temp_Id'], axis=1, inplace=True)

        # TODO:- New Data End
        ind_old_doors = df_old_doors[[col_name_lst[0], col_name_lst[1]]]
        dict_value = ind_old_doors.set_index('Object Identifier').T.to_dict('list')
        for j in range(len(ind_df['ID'])):
            for k, v in dict_value.items():
                if ind_df['ID'].values[j] == k:
                    ind_df['DOORS_OLD'].values[j] = v[0]

        compl_df = compl_df.append(ind_df)

    compl_df.drop('DOORS_NEW_DATA_877', axis=1, inplace=True)
    compl_df.drop('Count', axis=1, inplace=True)
    compl_df.drop('Variants', axis=1, inplace=True)
    compl_df.insert(6, column='Judgement ', value="")
    compl_df.insert(7, column='Description ', value="")
    # compl_df.set_index(['SrNo.'], inplace = True)
    # print(compl_df)
    by = ['SrNo.', 'ECN_New']  # groupby 'by' argument
    grp = compl_df.groupby(by).apply(lambda a: a[:])
    result = grp.iloc[:, 2:]
    excel_file = pd.ExcelWriter(
        "D:\\2.0 HONDA\\Automation_Dev\\PartNo_A3CNo_Conversation\\Result\\CONF_3A0A_Rel_5.5_MP_AfterUpdate_" + sh.formation_date + ".xlsx")
    result.to_excel(excel_file, sheet_name="Result", index=True)
    excel_file.save()
    ''' 
    Convert row values as a key and variant as a value into a dict:-
    The to_dict() method sets the column names as dictionary keys so you'll need to reshape your DataFrame slightly.
    Setting the 'Object Identifier' column as the index and then transposing the DataFrame is one way to achieve this.
    '''
except Exception as e:
    print("Oops!", e.__class__, "occurred in Report file.")
    print(e)
