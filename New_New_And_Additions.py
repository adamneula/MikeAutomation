import pandas as pd
import numpy as np
import os
from Rep_Objects import rep_lookup
from Utils import get_unique_filename, clean_numeric_columns
import Rep_Objects

def New_And_Addition(thisMonth: str, thisMonthSheet: str, lastMonth: str, lastMonthSheet: str, models: str | list[str], sheet_name: str) -> str:
    '''
    Generates augmented version of Account-Rep details with information about the account's balance last month
    information about the advisor, and whether the account is an Open or an Addition
    
    :arg1:
    Path in which to find account-rep details for current month
    :arg2:
    Name of account rep details sheet on this month's file
    :arg3:
    Path in which to find account-rep details for last month
    :arg4:
    Name of account rep details sheet on last month's file
    :arg5:
    List of models to do this for (string will work if it's just one)
    '''
    
    if isinstance(models, str):
        models = [models]
        
    df_raw = pd.read_excel(thisMonth, sheet_name=thisMonthSheet)
    df_raw.columns = df_raw.columns.str.strip()     #Clean Headers
    df_last_raw = pd.read_excel(lastMonth, sheet_name=lastMonthSheet)
    df_last_raw.columns = df_last_raw.columns.str.strip()
    
    cols_to_fix = ['ModelName', 'IBD/Sponsor Name', 'accountid']
    df_raw[cols_to_fix] = df_raw[cols_to_fix].astype(str).apply(lambda x: x.str.strip())
    df_last_raw[cols_to_fix] = df_last_raw[cols_to_fix].astype(str).apply(lambda x: x.str.strip())

    
    target_institution = "Primerica Brokerage Services"
    #Filter down to Target Institution and Models
    mask = (df_raw['ModelName'].isin(models)) & (df_raw['IBD/Sponsor Name'] == target_institution) 
    df = df_raw[mask].copy()    #Account-Rep Details with only the targeted models from Primerica
    mask = (df_last_raw['ModelName'].isin(models)) & (df_last_raw['IBD/Sponsor Name'] == target_institution) 
    df_last = df_last_raw[mask].copy()
    
    df = clean_numeric_columns(df, ['Total Assets'])
    df_last = clean_numeric_columns(df_last, ['Total Assets'])
    last_month_lookup = df_last[['accountid', 'ModelName', 'Total Assets']]
    
    df_final = pd.merge(
        df,
        last_month_lookup,
        on=['accountid', 'ModelName'],
        how='left',
        suffixes=('', '_prev')
    )
    df_final['Total Assets_prev'] = df_final['Total Assets_prev'].fillna(0)
    
    df_final['% Change'] = np.where(
    df_final['Total Assets_prev'] > 0, 
    (df_final['Total Assets'] - df_final['Total Assets_prev']) / df_final['Total Assets_prev'], 
    0)
    
    #Calculate Mode Growth to figure out growth from market conditions
    df_final['Mode.Sngl'] = df_final.groupby('ModelName')['% Change'].transform(
    lambda x: x[(x != 0)].round(4).mode().iloc[0] if not x.mode().empty else 0)
    df_final['Flow (MODE)'] = df_final['Total Assets'] - (df_final['Total Assets_prev'] * (1 + df_final['Mode.Sngl']))
    
    #Classify Open, Addition, or Reclassification
    threshold = 1000
    is_new_to_firm = ~df_final['accountid'].isin(df_last['accountid'])
    conditions = [
        # A: Truly Brand New (Not in last month's file at all)
        (df_final['Flow (MODE)'] >= threshold) & (df_final['Total Assets_prev'] <= 0) & is_new_to_firm,

        # B: Existing Client, but this is a New Model for them (Reallocation or Expansion)
        (df_final['Flow (MODE)'] >= threshold) & (df_final['Total Assets_prev'] <= 0) & (~is_new_to_firm),
        
        # C: Existing Client, Existing Model (Simple Addition)
        (df_final['Flow (MODE)'] >= threshold) & (df_final['Total Assets_prev'] > 0)
    ]
    
    choices = ['Open', 'Reclassification', 'Addition']
    df_final['Type'] = np.select(conditions, choices, default='')
    
    #Look up Rep information now
    df_final['Rep_Obj'] = df_final['Rep Name'].apply(rep_lookup)
    
    #Extract information from Rep_Obj into separate columns
    rep_id_values = df_final['Rep_Obj'].apply(lambda x: getattr(x, 'Primary_Rep_ID', 'Not Found'))
    df_final.insert(22, 'Rep ID', rep_id_values)
    df_final['AE'] = df_final['Rep_Obj'].apply(lambda x: getattr(x, 'AE', 'Not Found'))
    df_final['Territory'] = df_final['Rep_Obj'].apply(lambda x: getattr(x, 'Territory', 'Not Found'))
    df_final['True State'] = df_final['Rep_Obj'].apply(lambda x: getattr(x, 'True_State', 'Not Found'))
    df_final['Rep Email'] = df_final['Rep_Obj'].apply(lambda x: getattr(x, 'Email', 'Not Found'))
    
    df_final = df_final.drop(columns=['Rep_Obj'])
    
    #Handle output and format it right
    base_dir = os.path.dirname(os.path.abspath(__file__))
    clean_name = os.path.splitext(os.path.basename(thisMonth))[0]
    output_path = get_unique_filename(os.path.join(base_dir, f'{clean_name} - New and Additions ({sheet_name}).xlsx'))
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Styles
        base_style = {'font_name': 'Aptos Narrow', 'font_size': 11, 'border': 0}
        header_fmt = workbook.add_format({**base_style, 'bold': True, 'bottom': 1, 'bg_color': '#D9D9D9', 'align': 'left'})
        default_fmt = workbook.add_format(base_style)
        yellow_bg = workbook.add_format({**base_style, 'bg_color': '#FFFF00'})
        money_fmt = workbook.add_format({**base_style, 'num_format': '$#,##0.00'})
        money_yellow = workbook.add_format({**base_style, 'num_format': '$#,##0.00', 'bg_color': '#FFFF00'})
        percent_fmt = workbook.add_format({**base_style, 'num_format': '0.00%'})
        green_fmt = workbook.add_format({**base_style, 'bg_color': '#C6EFCE'})
        purple_fmt = workbook.add_format({**base_style, 'bg_color': '#E1D5E7'})
        orange_fmt = workbook.add_format({**base_style, 'bg_color': "#FFA000"})
        red_fmt = workbook.add_format({**base_style, 'bg_color': '#FF0000'})
        last_row = len(df_final)

        # Write headers
        for col_num, value in enumerate(df_final.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Apply Column Formatting
        yellow_targets = ['ModelCode', 'accountid', 'Total Assets', 'Rep Name', 'Rep City', 'Rep State', 'Flow (MODE)', 'AE', 'Territory']
        
        for i, col in enumerate(df_final.columns):
            max_len = min(max(df_final[col].astype(str).str.len().max(), len(str(col))) + 2, 40)
            
            if col == "Rep ID":
                fmt = orange_fmt
                worksheet.conditional_format(1, i, last_row, i, {'type': 'cell', 'criteria': 'equal to', 'value': '"Not Found"', 'format': red_fmt})
            elif col in yellow_targets:
                fmt = money_yellow if any(x in col for x in ['Assets', 'Flow']) else yellow_bg
            elif any(x in col for x in ['%', 'Mode']):
                fmt = percent_fmt
            elif any(x in col for x in ['Change', 'Assets', 'Flow']):
                fmt = money_fmt
        
            else:
                fmt = default_fmt 

            worksheet.set_column(i, i, max_len, fmt)

        # Conditional Formatting (Regional/Status)
        Type_idx = df_final.columns.get_loc('Type')
        terr_idx = df_final.columns.get_loc('Territory')
        ae_idx = df_final.columns.get_loc('AE')
        
        def get_col_letter(n):
            res = ""
            while n >= 0:
                res = chr(n % 26 + 65) + res
                n = n // 26 - 1
            return res
        
        t_letter = get_col_letter(terr_idx)

        worksheet.conditional_format(1, Type_idx, last_row, Type_idx, {'type': 'cell', 'criteria': 'equal to', 'value': '"Open"', 'format': green_fmt})
        worksheet.conditional_format(1, Type_idx, last_row, Type_idx, {'type': 'cell', 'criteria': 'equal to', 'value': '"Addition"', 'format': purple_fmt})
        worksheet.conditional_format(1, Type_idx, last_row, Type_idx, {'type': 'cell', 'criteria': 'equal to', 'value': '"Reclassification"', 'format': orange_fmt})

        for idx in [ae_idx, terr_idx]:
            worksheet.conditional_format(1, idx, last_row, idx, {'type': 'formula', 'criteria': f'=${t_letter}2="West"', 'format': purple_fmt})
            worksheet.conditional_format(1, idx, last_row, idx, {'type': 'formula', 'criteria': f'=${t_letter}2="East"', 'format': green_fmt})

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, last_row, len(df_final.columns) - 1)
    
    print(f"SUCCESS: Report saved to {os.path.abspath(output_path)}")
    return os.path.abspath(output_path)