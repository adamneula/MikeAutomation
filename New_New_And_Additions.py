import pandas as pd
import numpy as np
import os
from Rep_Objects import rep_lookup
from Utils import get_unique_filename, clean_numeric_columns

def New_And_Addition(thisMonth: str, thisMonthSheet: str, lastMonth: str, lastMonthSheet: str, models: str | list[str]):
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
    #Classify Open or Addition
    threshold = 1000
    conditions = [
        (df_final['Flow (MODE)'] >= threshold) & (df_final['Total Assets_prev'] == 0), # Brand new account
        (df_final['Flow (MODE)'] >= threshold) & (df_final['Total Assets_prev'] > 0)  # Existing account adding money
    ]
    choices = ['Open', 'Addition']
    df_final['Type'] = np.select(conditions, choices, default='')
    
    output_path = get_unique_filename("New_and_Additions_Check.xlsx")
    df_final.to_excel(output_path, index=False)


New_And_Addition('H:\_INSTITUTIONAL DIVISION\INTERN FOLDER\Adam Neulander\MikeAutomation\ModelProvider_AUM_RNC_FEB2026_Pivot.xlsx',
                 'Account-Rep Details',
                 'H:\_INSTITUTIONAL DIVISION\INTERN FOLDER\Adam Neulander\MikeAutomation\ModelProvider_AUM_RNC_JAN2026_Pivot.xlsx',
                 'Account-Rep Details',
                 ['Genter Capital Balanced Growth with GENM',
                  'Genter Capital Balanced Growth with GENT',
                  'Genter Capital Balanced Income with GENM',
                  'Genter Capital Balanced Income with GENT',
                  'Genter Capital Balanced with GENM',
                  'Genter Capital Balanced with GENT'])

