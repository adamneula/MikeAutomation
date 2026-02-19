import pandas as pd
import numpy as np
import os
from datetime import datetime
from Rep_Objects import rep_lookup
from Utils import *

def Primerica_Div_Model_New_And_Addition(thisMonth, thisMonthSheet, lastMonth, lastMonthSheet):
    # --- DIAGNOSTIC 1: Initial Load ---
    df_raw = pd.read_excel(thisMonth, sheet_name=thisMonthSheet)
    df_raw.columns = df_raw.columns.str.strip()
    print(f"DEBUG: Total rows in raw file: {len(df_raw)}")

    # Clean the ModelName data itself (not just the header)
    df_raw['ModelName'] = df_raw['ModelName'].astype(str).str.strip()
    
    # --- DIAGNOSTIC 2: The Filter ---
    target_model = 'Genter Capital Dividend Income Model'
    target_institution = 'Primerica Brokerage Services'

    # Filter for BOTH at once. 
    # Use .str.contains if the names might have extra spaces
    df = df_raw[(df_raw['ModelName'] == target_model) & (df_raw['IBD/Sponsor Name'] == target_institution)].copy()
    
    if len(df) == 0:
        print("DEBUG: Available models in file were:", df_raw['ModelName'].unique()[:10])

    # --- DIAGNOSTIC 3: The ID Match ---
    df_prev_raw = pd.read_excel(lastMonth, sheet_name=lastMonthSheet)
    df_prev_raw.columns = df_prev_raw.columns.str.strip()
    
    # Cast to string to prevent Int vs String mismatch
    df['accountid'] = df['accountid'].astype(str).str.strip()
    df_prev_raw['accountid'] = df_prev_raw['accountid'].astype(str).str.strip()
    
    prev_assets_map = dict(zip(df_prev_raw['accountid'], df_prev_raw['Total Assets']))
    
    # Count how many current accounts exist in the previous month's map
    match_count = df['accountid'].isin(prev_assets_map.keys()).sum()
    print(f"DEBUG: Out of {len(df)} accounts, {match_count} were found in last month's data.")
    print(f"DEBUG: {len(df) - match_count} accounts are being treated as 'New Open'.")

    # ... proceed with the rest of your logic ...
    
    rep_name_idx = df.columns.get_loc('Rep Name')
    #V
    df.insert(rep_name_idx + 1, 'Rep ID', df['Rep Name'].apply(lambda x: rep_lookup(x).Primary_Rep_ID if rep_lookup(x) else 'Not Found'))
    #AD
    df['Rep Email'] = df['Rep Name'].apply(lambda x: rep_lookup(x).Email if rep_lookup(x) else 'Not Found')
    #AF
    df['Prev Month Assets'] = df['accountid'].map(prev_assets_map).fillna(0)
    #AG
    df['Total Assets'] = pd.to_numeric(df['Total Assets'], errors='coerce').fillna(0)
    df['Prev Month Assets'] = pd.to_numeric(df['Prev Month Assets'], errors='coerce').fillna(0)
    df['$ Change'] = df['Total Assets'] - df['Prev Month Assets']
    #AH
    df['% Change'] = np.where(df['Prev Month Assets'] > 0, df['$ Change'] / df['Prev Month Assets'], 0)
    #AI
    mode_series = df.loc[df['Prev Month Assets'] > 0, '% Change'].round(4).mode()
    market_benchmark = mode_series.iloc[0] if not mode_series.empty else 0
    df['Mode.Sngl'] = market_benchmark
    #AJ
    df['Flow (MODE)'] = df['$ Change'] - (df['Prev Month Assets'] * df['Mode.Sngl'])
    #AK
    df['Type'] = np.where(df['Flow (MODE)'] < 10000, '', np.where(df['Prev Month Assets'] > 0, 'Addition', 'Open'))
    #AL
    df['True State'] = df['Rep Name'].apply(lambda x: rep_lookup(x).True_State if rep_lookup(x) else '')
    #AM
    df['AE'] = df['Rep Name'].apply(lambda x: rep_lookup(x).AE if rep_lookup(x) else '')
    #AN
    df['Territory'] = df['Rep Name'].apply(lambda x: rep_lookup(x).Territory if rep_lookup(x) else '')
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    clean_name = os.path.splitext(os.path.basename(thisMonth))[0]
    raw_filename = f'{clean_name} - New and Additions.xlsx'
    output_path = get_unique_filename(os.path.join(base_dir, raw_filename))
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        sheet_name = "Primerica Div Model"
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # --- Define Formats ---
        yellow_bg = workbook.add_format({'bg_color': '#FFFF00'})
        orange_bg = workbook.add_format({'bg_color': '#FFC000'})
        
        # Money/Percent with Yellow/Orange overrides
        money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
        percent_fmt = workbook.add_format({'num_format': '0.00%'})
        money_yellow = workbook.add_format({'num_format': '$#,##0.00', 'bg_color': '#FFFF00'})
        
        # Type/Territory Formats
        green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'font': ''})
        purple_fmt = workbook.add_format({'bg_color': '#E1D5E7', 'font_color': '#400080'})

        # --- 1. Basic Setup ---
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        worksheet.freeze_panes(1, 0)

        # --- 2. Static Column Highlights (M, P, V, AA, AB, AJ, AL in Yellow | W in Orange) ---
        # Map Excel letters to 0-based indices for the loop
        yellow_cols = ['M', 'P', 'V', 'AA', 'AB', 'AJ', 'AL']
        orange_cols = ['W']

        yellow_indices = [col_to_idx(c) for c in yellow_cols]
        orange_indices = [col_to_idx(c) for c in orange_cols]

        # --- 3. The Main Formatting Loop ---
        for i, col in enumerate(df.columns):
            max_len = max(df[col].fillna('').astype(str).map(len).max(), len(str(col))) + 2
            max_len = min(max_len, 50)
            
            # Determine Base Format
            fmt = None
            if i in yellow_indices:
                fmt = money_yellow if any(x in col for x in ['Assets', 'Change', 'Flow (MODE)']) else yellow_bg
            elif i in orange_indices:
                fmt = orange_bg
            elif any(x in col for x in ['$ Change', 'Assets', 'Flow (MODE)']):
                fmt = money_fmt
            elif any(x in col for x in ['% Change', 'Mode.Sngl']):
                fmt = percent_fmt
            
            worksheet.set_column(i, i, max_len, fmt)

        # --- 4. Conditional Formatting (Type & Territory) ---
        last_row = len(df)
        Type_idx = df.columns.get_loc('Type')
        ae_idx = df.columns.get_loc('AE')
        terr_idx = df.columns.get_loc('Territory')

        # Type: Open (Green) & Addition (Purple)
        worksheet.conditional_format(1, Type_idx, last_row, Type_idx, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Open"', 'format': green_fmt})
        worksheet.conditional_format(1, Type_idx, last_row, Type_idx, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Addition"', 'format': purple_fmt})

        # Territory: West (Purple) & East (Green)
        # We apply this to both the AE and Territory columns
        for idx in [ae_idx, terr_idx]:
            worksheet.conditional_format(1, idx, last_row, idx, 
                                         {'type': 'formula', 'criteria': f'=$AN2="West"', 'format': purple_fmt})
            worksheet.conditional_format(1, idx, last_row, idx, 
                                         {'type': 'formula', 'criteria': f'=$AN2="East"', 'format': green_fmt})
    
    print(f"SUCCESS: Detailed report with color-coding saved to {os.path.abspath(output_path)}")
    return os.path.abspath(output_path)

def GenT_GenM_New_And_Addition(thisMonth, thisMonthSheet, lastMonth, lastMonthSheet):
    # --- 1. Initial Load ---
    df_raw = pd.read_excel(thisMonth, sheet_name=thisMonthSheet)
    df_raw.columns = df_raw.columns.str.strip()
    df_raw['ModelName'] = df_raw['ModelName'].astype(str).str.strip()

    df_prev_raw = pd.read_excel(lastMonth, sheet_name=lastMonthSheet)
    df_prev_raw.columns = df_prev_raw.columns.str.strip()
    
    df_prev_raw['accountid'] = df_prev_raw['accountid'].astype(str).str.strip()
    prev_assets_map = dict(zip(df_prev_raw['accountid'], df_prev_raw['Total Assets']))

    models_to_process = [
        'Genter Capital Balanced Growth with GENT', 'Genter Capital Balanced Growth with GENM',
        'Genter Capital Balanced Income with GENT', 'Genter Capital Balanced Income with GENM',
        'Genter Capital Balanced with GENT', 'Genter Capital Balanced with GENM'
    ]

    all_model_dfs = []

    # --- 2. Processing Loop ---
    for model_name in models_to_process:
        df = df_raw[(df_raw['ModelName'] == model_name) & 
                    (df_raw['IBD/Sponsor Name'] == 'Primerica Brokerage Services')].copy()
        
        if df.empty:
            print(f"DEBUG: No data found for model: {model_name}")
            continue

        # Insert 'Rep ID' to the right of 'Rep Name'
        if 'Rep Name' in df.columns:
            rep_name_idx = df.columns.get_loc('Rep Name')
            df.insert(rep_name_idx + 1, 'Rep ID', df['Rep Name'].apply(
                lambda x: rep_lookup(x).Primary_Rep_ID if rep_lookup(x) else 'N/A'
            ))

        # Financial Calculations
        df['accountid'] = df['accountid'].astype(str).str.strip()
        df['Prev Month Assets'] = df['accountid'].map(prev_assets_map).fillna(0)
        df['Total Assets'] = pd.to_numeric(df['Total Assets'], errors='coerce').fillna(0)
        df['$ Change'] = df['Total Assets'] - df['Prev Month Assets']
        df['% Change'] = np.where(df['Prev Month Assets'] > 0, df['$ Change'] / df['Prev Month Assets'], 0)
        
        # Mode/Flow/Type
        mode_val = df.loc[df['Prev Month Assets'] > 0, '% Change'].round(4).mode()
        df['Mode.Sngl'] = mode_val.iloc[0] if not mode_val.empty else 0
        df['Flow (MODE)'] = df['$ Change'] - (df['Prev Month Assets'] * df['Mode.Sngl'])
        #TODO: Get BD performance in somehow so I can use this, until then fall back on Flow from Mode
        #df['Flow (BD)'] = df['$ Change'] - (df['Prev Month Assets']*df['BD Performance'])
        df['Type'] = np.where(df['Flow (MODE)'] < 1000, '', np.where(df['Prev Month Assets'] > 0, 'Addition', 'Open'))
        
        # Meta Data
        df['AE'] = df['Rep Name'].apply(lambda x: rep_lookup(x).AE if rep_lookup(x) else '')
        df['Territory'] = df['Rep Name'].apply(lambda x: rep_lookup(x).Territory if rep_lookup(x) else '')

        all_model_dfs.append(df)

    # --- 3. Consolidation & Check ---
    if not all_model_dfs:
        print("\n[!] ERROR: all_model_dfs is empty. No data matched the model names or institution.")
        print(f"Check if models like 'Balanced Growth with GENT' exist in the '{thisMonthSheet}' sheet.")
        return

    final_df = pd.concat(all_model_dfs, ignore_index=True)
    output_path = get_unique_filename(f'Consolidated_Models_{datetime.now().strftime("%b_%Y")}.xlsx')

    # --- 4. Export & Formatting ---
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, sheet_name="GENT and GENM", index=False)
        workbook, worksheet = writer.book, writer.sheets["GENT and GENM"]

        # --- 1. THE FOUNDATION: Shared Font (No Borders) ---
        base_style = {'font_name': 'Aptos Narrow', 'font_size': 11, 'border': 0}

        # --- 2. DEFINE FORMATS ---
        header_fmt = workbook.add_format({**base_style, 'bold': True, 'bottom': 1, 'bg_color': '#D9D9D9', 'align': 'left'})
        
        # Data Formats
        default_fmt = workbook.add_format(base_style)
        yellow_bg = workbook.add_format({**base_style, 'bg_color': '#FFFF00'})
        
        # Numeric formats
        money_fmt = workbook.add_format({**base_style, 'num_format': '$#,##0.00'})
        money_yellow = workbook.add_format({**base_style, 'num_format': '$#,##0.00', 'bg_color': '#FFFF00'})
        percent_fmt = workbook.add_format({**base_style, 'num_format': '0.00%'})
        
        # Conditional formats (Regional/Type) - Removed borders here too
        green_fmt = workbook.add_format({**base_style, 'bg_color': '#C6EFCE', 'font_color': '#006100'})
        purple_fmt = workbook.add_format({**base_style, 'bg_color': '#E1D5E7', 'font_color': '#400080'})

        # --- 3. APPLY BOLD HEADERS ---
        for col_num, value in enumerate(final_df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # --- 4. UPDATED YELLOW TARGETS ---
        yellow_target_cols = [
            'ModelCode', 'accountid', 'Total Assets', 'Rep Name', 
            'Rep City', 'Rep State', 'Flow (MODE)', 'AE', 'Territory'
        ]
        
        for i, col in enumerate(final_df.columns):
            # Safe width calculation using .str.len()
            max_val_len = final_df[col].astype(str).str.len().max()
            max_len = min(max(max_val_len, len(str(col))) + 2, 40)
            
            # Formatting Decision
            if col in yellow_target_cols:
                fmt = money_yellow if any(x in col for x in ['Assets', 'Flow (MODE)']) else yellow_bg
            elif any(x in col for x in ['Change', 'Assets', 'Flow (MODE)']):
                fmt = money_fmt
            elif any(x in col for x in ['Change', 'Mode']):
                fmt = percent_fmt
            else:
                fmt = default_fmt 

            # Applying fmt here forces Calibri on the whole column
            worksheet.set_column(i, i, max_len, fmt)

        # --- 5. CONDITIONAL FORMATTING (Regional Sync) ---
        last_row = len(final_df)
        Type_idx = final_df.columns.get_loc('Type')
        ae_idx = final_df.columns.get_loc('AE')
        terr_idx = final_df.columns.get_loc('Territory')
        
        # Find Excel Column Letter (e.g., 'AN') for the Territory column
        def get_excel_letter(n):
            res = ""
            while n >= 0:
                res = chr(n % 26 + 65) + res
                n = n // 26 - 1
            return res
        
        t_letter = get_excel_letter(terr_idx)

        # Type rules
        worksheet.conditional_format(1, Type_idx, last_row, Type_idx, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Open"', 'format': green_fmt})
        worksheet.conditional_format(1, Type_idx, last_row, Type_idx, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Addition"', 'format': purple_fmt})

        # Regional rules: Applying to AE and Territory based on Territory value
        for idx in [ae_idx, terr_idx]:
            worksheet.conditional_format(1, idx, last_row, idx, 
                                         {'type': 'formula', 
                                          'criteria': f'=${t_letter}2="West"', 
                                          'format': purple_fmt})
            worksheet.conditional_format(1, idx, last_row, idx, 
                                         {'type': 'formula', 
                                          'criteria': f'=${t_letter}2="East"', 
                                          'format': green_fmt})

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, last_row, len(final_df.columns) - 1)

    print(f"SUCCESS: Report saved at {os.path.abspath(output_path)}")
    return os.path.abspath(output_path)