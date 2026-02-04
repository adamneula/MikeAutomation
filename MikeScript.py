import pandas as pd
import os
from datetime import datetime, timedelta
import numpy as np
from dateutil.relativedelta import relativedelta
from tqdm import tqdm

tqdm.pandas()

reps = {}
IDtoName = {}
States = {
    'AL', 'AK', 'AZ', 'AR', 'AB', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 
    'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD', 
    'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 
    'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'PR', 'RI', 'SC', 
    'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY', 'DC'
}
NamesNotFound = {}

class Representatives():
    def __init__(self, name: str, ID: str, state: str, email: str, AE: str, territory: str, total: float):
        self.Advisor_Name = name
        self.Primary_Rep_ID = ID
        self.True_State = state
        self.Email = email
        self.AE = AE
        self.Territory = territory
        self.Ranking = None
        self.Sum_of_Total_Assets = 0
        self.Previous_Month_AUM = None
        self.MoM_Change = None
        self.Dollar_Val_Change = None
        self.Lifetime_Total = total
        
    def __hash__(self):
        return hash(self.Primary_Rep_ID)

    def __eq__(self, other):
        if not isinstance(other, Representatives):
            return False
        return self.Primary_Rep_ID == other.Primary_Rep_ID
    
    def __str__(self):
        return f'{self.Advisor_Name} {self.True_State} {self.Primary_Rep_ID} {self.Email} {self.Territory} AE: {self.AE} Balance: {self.Sum_of_Total_Assets}'

    def add_account(self, amount):
        self.Sum_of_Total_Assets += amount
        
def load_reps_from_xlsx(Fit_List_Dir, Fit_List_Sheet_Name):
    global reps
    
    # header=1 skips the 'Owned...' row and uses the 'Code, Mutual...' row as headers
    df = pd.read_excel(Fit_List_Dir, sheet_name=Fit_List_Sheet_Name, header=1)
    
    # Standardize column names to remove any accidental spaces
    df.columns = df.columns.str.strip()
    
    # Drop rows where ID is missing
    df = df.dropna(subset=['ID'])
    
    for _, row in df.iterrows():
        first_name = str(row['First']).strip()
        last_name = str(row['Last']).strip()
        if first_name.lower() == 'christophe': first_name = 'CHRISTOPHER'
        elif first_name.lower() == 'theodore' and last_name.lower() == 'lund': first_name = 'TED'
        elif first_name.lower() == 'danny' and last_name.lower() == 'creswell': first_name = 'DANIEL'
        full_name = f"{first_name} {last_name}"
        clean_ID = str(row['ID']).replace(' ', '').strip()        
        IDtoName[clean_ID] = full_name.lower()
        
        total = float(row['LifeTime'])
        if full_name.lower() in reps:
            if reps[full_name.lower()].Lifetime_Total > total: continue
        state = str(row['State']).strip()
        email = str(row['Pol Email']).strip()
        territory = str(row['Territory']).strip()
        AE = ''
        #Sets central to East region and assigns AE accordingly
        if state.lower() in ['ok', 'ks']:
            territory = 'East'
            AE = 'Rob Hunt'
        if territory.lower() == 'central':
            territory = 'East' 
            AE = 'Rob Hunt'
        elif territory == 'East': AE = 'Rob Hunt'
        elif territory == 'West': AE = 'MeiWah Wong'
        reps[full_name.lower()] = Representatives(full_name, clean_ID, state, email, AE, territory, total)
            
def attribute_accounts(Primerica_Dir, Primerica_Sheet_Name):
    global reps

    df = pd.read_excel(Primerica_Dir, sheet_name=Primerica_Sheet_Name, header=0)
    df.columns = df.columns.str.strip()
    
    for index, row in df.iterrows():
        clean_Name = str(row['Rep Name']).strip()
        if clean_Name == 'nan': continue
        elif clean_Name.lower().split()[0] == 'christophe':
            clean_Name = " ".join(['CHRISTOPHER'] + clean_Name.upper().split()[1:])
        elif clean_Name.lower() == 'danny creswell': clean_Name = 'DANIEL CRESWELL'
        
        if clean_Name.lower() in reps:
            reps[clean_Name.lower()].add_account(row['Total Assets'])
        elif clean_Name[:5] in IDtoName:
            reps[IDtoName[clean_Name[:5]]].add_account(row['Total Assets'])

        elif clean_Name.replace(' ', '') in IDtoName:
            reps[IDtoName[clean_Name.replace(' ', '')]].add_account(row['Total Assets'])

        else:
            NamesNotFound[clean_Name] = index
    
    for rep in reps:
        #assign ranking
        if reps[rep].Sum_of_Total_Assets == 0: continue
        elif reps[rep].Sum_of_Total_Assets < 250000: reps[rep].Ranking = 'C'
        elif reps[rep].Sum_of_Total_Assets < 1000000: reps[rep].Ranking = 'B'
        elif reps[rep].Sum_of_Total_Assets < 2000000: reps[rep].Ranking = 'BB'
        elif reps[rep].Sum_of_Total_Assets < 5000000: reps[rep].Ranking = 'A'
        elif reps[rep].Sum_of_Total_Assets < 10000000: reps[rep].Ranking = 'AA'
        else: reps[rep].Ranking = 'AAA'
    
def load_previous_month_data(prev_month_file, prev_month_sheet):
    try:
        df_prev = pd.read_excel(prev_month_file, sheet_name=prev_month_sheet, header=0)
        df_prev.columns = df_prev.columns.str.strip()
    except Exception as e:
        print(f"Error loading {prev_month_file}: {e}")
        return

    for _, row in df_prev.iterrows():
        # 1. Look up the user's ID and clean it
        raw_id = str(row['Primary Rep ID']).replace(' ', '').strip().upper()
        if raw_id == 'NAN': continue
        
        # 2. Find the proper name from the ID lookup list
        proper_name_lower = IDtoName.get(raw_id)
        
        # 3. Use that name to go into the advisor object and add attributes
        if proper_name_lower and proper_name_lower in reps:
            advisor = reps[proper_name_lower]
            
            # Fill the instance variables
            prev_bal = float(row['Sum of Total Assets']) if pd.notna(row['Sum of Total Assets']) else 0.0
            advisor.Previous_Month_AUM = prev_bal
            
            # Calculate changes automatically
            advisor.Dollar_Val_Change = advisor.Sum_of_Total_Assets - prev_bal
            if prev_bal > 0:
                advisor.MoM_Change = advisor.Dollar_Val_Change / prev_bal
            else:
                advisor.MoM_Change = 0.0
                           
def export_to_pivot(fit_path='', fit_sheet='', details_path='', details_sheet='', pivot_path='', pivot_sheet=''):
    global reps
    
    # 1. Prepare Data
    data = []
    sorted_reps = sorted(reps.values(), key=lambda x: x.Advisor_Name)
    
    for r in sorted_reps:
        if r.Sum_of_Total_Assets == 0 and (r.Previous_Month_AUM is None or r.Previous_Month_AUM == 0):
            continue
            
        row_data = {
            'Row Labels': r.Advisor_Name.upper(),
            'Primary Rep ID': r.Primary_Rep_ID,
            'True State': r.True_State,
            'AE': r.AE,
            'Territory': r.Territory,
            'Sum of Total Assets ': r.Sum_of_Total_Assets,
            'Spacer_1': '', 'Spacer_2': '',
            'Advisor Name': r.Advisor_Name.upper(),
            'True State.1': r.True_State,
            'AE.1': r.AE,
            'Territory.1': r.Territory,
            'Email': r.Email,
            'Primary Rep ID.1': r.Primary_Rep_ID,
            'Ranking.1': r.Ranking,
            'Sum of Total Assets .1': r.Sum_of_Total_Assets,
            'Previous Month AUM': r.Previous_Month_AUM,
            'MoM Change': r.MoM_Change,
            'Dollar Val Change': r.Dollar_Val_Change
        }
        data.append(row_data)

    df_output = pd.DataFrame(data).rename(columns={'Spacer_1': '', 'Spacer_2': ' '})
    base_dir = os.path.dirname(os.path.abspath(__file__))
    clean_name = os.path.splitext(os.path.basename(details_path))[0]
    raw_filename = f'{clean_name}_Pivot.xlsx'
    output_path = get_unique_filename(os.path.join(base_dir, raw_filename))
    
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_output.to_excel(writer, sheet_name=f'AUM Pivot - {(datetime.now() - relativedelta(months=1)).strftime("%b %y")}', index=False)
            workbook  = writer.book
            worksheet = writer.sheets[f'AUM Pivot - {(datetime.now() - relativedelta(months=1)).strftime("%b %y")}']

            # --- Define Styles ---
            dark_blue = workbook.add_format({'bg_color': '#1F4E78', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center'})
            light_blue = workbook.add_format({'bg_color': "#C7E5F3", 'font_color': 'black', 'bold': True, 'border': 1, 'align': 'center'})
            money_fmt = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
            percent_fmt = workbook.add_format({'num_format': '0.0%', 'border': 1})
            border_fmt = workbook.add_format({'border': 1})
            
            # Ranking Legend Hex Codes (Updated with your swatch codes)
            rank_colors = {
                'AAA': '#00B050', 'AA': '#92D050', 'A': '#E2F0D9',
                'BB': '#00B0F0', 'B': '#B4C6E7', 'C': '#FFFF00'
            }

            # --- Apply Formatting ---
            # 1. Header Styling & Active Sort Buttons
            # We explicitly set the autofilter across all columns (0 to end)
            num_rows = len(df_output)
            num_cols = len(df_output.columns)
            #This puts the sorting box on the appropriate columns
            worksheet.autofilter(0, 0, num_rows, 1)
            worksheet.autofilter(0, 8, num_rows, 18)

            for col_num, value in enumerate(df_output.columns.values):
                if col_num < 6: # Left side group
                    worksheet.write(0, col_num, value, light_blue)
                elif col_num > 7: # Right side group
                    worksheet.write(0, col_num, value, dark_blue)
                else: # Spacers G and H
                    worksheet.write(0, col_num, "", border_fmt)

            # 2. Column-Specific Formatting (Widths & Ranking Colors)
            apply_excel_highlighting(workbook, worksheet, df_output)
            for i, col in enumerate(df_output.columns):
                # This safely calculates length even if there are floats/NaNs
                # We convert every item to a string, find its length, and take the max
                column_data = df_output[col].fillna('') # Fill NaNs with empty strings first
                max_len = max(
                    column_data.astype(str).map(len).max(), 
                    len(str(col))
                ) + 2
                
                max_len = min(max_len, 50) # Keep it reasonable
                
                # Apply column specific formatting
                if any(x.lower() in col.lower() for x in ['assets', 'aum', 'dollar']):
                    worksheet.set_column(i, i, 21, money_fmt)
                elif 'Change' in col:
                    worksheet.set_column(i, i, 12, percent_fmt)
                else:
                    worksheet.set_column(i, i, max_len, border_fmt)

            # --- Final Source Tabs ---
            try:
                pd.read_excel(fit_path, sheet_name=fit_sheet, header=1).to_excel(writer, sheet_name='Source_FIT', index=False)
                pd.read_excel(details_path, sheet_name=details_sheet).to_excel(writer, sheet_name='Source_Details', index=False)
                pd.read_excel(pivot_path, sheet_name=pivot_sheet).to_excel(writer, sheet_name='Source_Pivots', index=False)
            except:
                print(f"\nSOURCE TAB ERROR: {e}")
            
                            
    #PRINT LEGEND HERE

            legend_start_row = 1
            legend_col_label = 20 # Column U
            legend_col_val = 21   # Column V
            
            worksheet.write(0, legend_col_label, "Ranking Legend (minimum):", workbook.add_format({'bold': True}))
            
            legend_items = [
                ('', '', '#FFFFFF'),
                ('AAA', 10000000, '#00B050'), # Green
                ('AA', 5000000, '#92D050'),   # Light Green
                ('A', 2000000, '#FCE4D6'),    # Peach/Tan
                ('BB', 1000000, '#00B0F0'),   # Blue
                ('B', 250000, '#B4C6E7'),     # Light Blue
                ('C', 0, '#FFFF00'),          # Yellow
            ]

            for i, (rank, val, color) in enumerate(legend_items):
                row = legend_start_row + i
                fmt = workbook.add_format({'bg_color': color, 'border': 1})
                money_fmt_legend = workbook.add_format({'bg_color': color, 'border': 1, 'num_format': '$#,##0.00'})
                
                worksheet.write(row, legend_col_label, rank, fmt)
                worksheet.write(row, legend_col_val, val, money_fmt_legend)

            # 2. Add the Summary Totals (Total Assets, East, West, etc.)
            summary_start_row = legend_start_row + len(legend_items) + 1
            bold_border = workbook.add_format({'bold': True, 'border': 1})
            money_bold = workbook.add_format({'bold': True, 'border': 1, 'num_format': '$#,##0.00'})
            percent_bold = workbook.add_format({'bold': True, 'border': 1, 'num_format': '0.00%'})

            # Calculations for the summary
            total_assets = df_output['Sum of Total Assets '].sum()
            east_assets = df_output[df_output['Territory'] == 'East']['Sum of Total Assets '].sum()
            west_assets = df_output[df_output['Territory'] == 'West']['Sum of Total Assets '].sum()
            prev_aum = df_output['Previous Month AUM'].sum()
            mom_growth = (total_assets - prev_aum) / prev_aum if prev_aum > 0 else 0

            summary_data = [
                (f"Total Assets {(datetime.now().replace(day=1) - timedelta(days=1)).strftime('%m/%d/%Y')}:", total_assets, money_bold),
                ("East", east_assets, money_bold),
                ("West", west_assets, money_bold),
                ("", None, None),
                (f"{(datetime.now() - relativedelta(months=2)).strftime('%b %y')} AUM", prev_aum, money_bold),
                ("", None, None),
                ("MoM Growth", mom_growth, percent_bold)
            ]

            for i, (label, value, fmt) in enumerate(summary_data):
                curr_row = summary_start_row + i
                if label:
                    worksheet.write(curr_row, legend_col_label, label, bold_border)
                if value is not None:
                    worksheet.write(curr_row, legend_col_val, value, fmt)

        print(f"\nSUCCESS: Report generated at {output_path}")
    
    except Exception as e:
        print(f"\nERROR: {e}")
    
def apply_excel_highlighting(workbook, worksheet, df):
    # 1. Define the Ranking Legend Hex Codes
    rank_colors = {
        'AAA': '#00B050', 'AA': '#92D050', 'A': '#E2F0D9',
        'BB': '#00B0F0', 'B': '#B4C6E7', 'C': '#FFFF00'
    }

    # 2. Pre-create formats
    # Note: Positive and Negative are independent of the AAA-C ranking
    pos_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1, 'num_format': '0.00%', 'align': 'center', 'font_color': '#000000'})
    neg_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1, 'num_format': '0.00%', 'align': 'center', 'font_color': '#9C0006'})
    
    rank_formats = {}
    for rank, hex_code in rank_colors.items():
        rank_formats[rank] = {
            'text': workbook.add_format({'bg_color': hex_code, 'border': 1, 'align': 'center'}),
            'money': workbook.add_format({'bg_color': hex_code, 'border': 1, 'num_format': '$#,##0.00'})
        }

    # 3. Identify column indices safely
    try:
        rank_col_idx = df.columns.get_loc('Ranking.1')
        assets_col_idx = df.columns.get_loc('Sum of Total Assets ')
        assets1_col_idx = df.columns.get_loc('Sum of Total Assets .1')
        mom_change_idx = df.columns.get_loc('MoM Change')
    except KeyError as e:
        print(f"Warning: Could not find column {e} for highlighting.")
        return

    # 4. Loop through every row
    for row_num in range(len(df)):
        # --- Handle Ranking and Assets ---
        rank_val = str(df.iloc[row_num]['Ranking.1']).strip()
        
        if rank_val in rank_formats:
            # Highlight Ranking
            worksheet.write(row_num + 1, rank_col_idx, rank_val, rank_formats[rank_val]['text'])
            
            # Highlight both Asset columns with the same rank color
            asset_val = df.iloc[row_num]['Sum of Total Assets .1']
            worksheet.write(row_num + 1, assets_col_idx, asset_val, rank_formats[rank_val]['money'])
            worksheet.write(row_num + 1, assets1_col_idx, asset_val, rank_formats[rank_val]['money'])
        
        # --- Handle MoM Change Highlighting ---
        mom_change_val = df.iloc[row_num]['MoM Change']
        
        # Check if it's a number and not NaN
        if pd.notna(mom_change_val):
            if mom_change_val > 0:
                worksheet.write(row_num + 1, mom_change_idx, mom_change_val, pos_fmt)
            elif mom_change_val < 0:
                worksheet.write(row_num + 1, mom_change_idx, mom_change_val, neg_fmt)
    
def Primerica_Div_Model(thisMonth, thisMonthSheet, lastMonth, lastMonthSheet):
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
    df['Flow'] = df['$ Change'] - (df['Prev Month Assets'] * df['Mode.Sngl'])
    #AK
    df['Status'] = np.where(df['Flow'] < 10000, '', np.where(df['Prev Month Assets'] > 0, 'Addition', 'Open'))
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
        yellow_bg = workbook.add_format({'bg_color': '#FFFF00', 'border': 1})
        orange_bg = workbook.add_format({'bg_color': '#FFC000', 'border': 1})
        
        # Money/Percent with Yellow/Orange overrides
        money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
        percent_fmt = workbook.add_format({'num_format': '0.00%'})
        money_yellow = workbook.add_format({'num_format': '$#,##0.00', 'bg_color': '#FFFF00', 'border': 1})
        
        # Status/Territory Formats
        green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        purple_fmt = workbook.add_format({'bg_color': '#E1D5E7', 'font_color': '#400080'})

        # --- 1. Basic Setup ---
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        worksheet.freeze_panes(1, 0)

        # --- 2. Static Column Highlights (M, P, V, AA, AB, AJ, AL in Yellow | W in Orange) ---
        # Map Excel letters to 0-based indices for the loop
        yellow_cols = ['M', 'P', 'V', 'AA', 'AB', 'AJ', 'AL']
        orange_cols = ['W']
        
        def col_to_idx(col_letter):
            # Converts 'A' -> 0, 'B' -> 1, etc.
            num = 0
            for c in col_letter:
                num = num * 26 + (ord(c.upper()) - ord('A') + 1)
            return num - 1

        yellow_indices = [col_to_idx(c) for c in yellow_cols]
        orange_indices = [col_to_idx(c) for c in orange_cols]

        # --- 3. The Main Formatting Loop ---
        for i, col in enumerate(df.columns):
            max_len = max(df[col].fillna('').astype(str).map(len).max(), len(str(col))) + 2
            max_len = min(max_len, 50)
            
            # Determine Base Format
            fmt = None
            if i in yellow_indices:
                fmt = money_yellow if any(x in col for x in ['Assets', 'Change', 'Flow']) else yellow_bg
            elif i in orange_indices:
                fmt = orange_bg
            elif any(x in col for x in ['$ Change', 'Assets', 'Flow']):
                fmt = money_fmt
            elif any(x in col for x in ['% Change', 'Mode.Sngl']):
                fmt = percent_fmt
            
            worksheet.set_column(i, i, max_len, fmt)

        # --- 4. Conditional Formatting (Status & Territory) ---
        last_row = len(df)
        status_idx = df.columns.get_loc('Status')
        ae_idx = df.columns.get_loc('AE')
        terr_idx = df.columns.get_loc('Territory')

        # Status: Open (Green) & Addition (Purple)
        worksheet.conditional_format(1, status_idx, last_row, status_idx, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Open"', 'format': green_fmt})
        worksheet.conditional_format(1, status_idx, last_row, status_idx, 
                                     {'type': 'cell', 'criteria': 'equal to', 'value': '"Addition"', 'format': purple_fmt})

        # Territory: West (Purple) & East (Green)
        # We apply this to both the AE and Territory columns
        for idx in [ae_idx, terr_idx]:
            worksheet.conditional_format(1, idx, last_row, idx, 
                                         {'type': 'formula', 'criteria': f'=$AN2="West"', 'format': purple_fmt})
            worksheet.conditional_format(1, idx, last_row, idx, 
                                         {'type': 'formula', 'criteria': f'=$AN2="East"', 'format': green_fmt})

    print(f"SUCCESS: Detailed report with color-coding saved to {os.path.abspath(output_path)}")

#HELPER FUNCTIONS
def get_unique_filename(file_path):
    """Checks if a file exists and appends a numeric suffix if it does."""
    if not os.path.exists(file_path):
        return file_path

    # Split into file path/name and the .xlsx extension
    base, extension = os.path.splitext(file_path)
    counter = 1
    
    # Try 'FileName 1.xlsx', 'FileName 2.xlsx', etc.
    new_path = f"{base} {counter}{extension}"
    while os.path.exists(new_path):
        counter += 1
        new_path = f"{base} {counter}{extension}"
        
    return new_path

def rep_lookup(input_str) -> Representatives:
    global reps, IDtoName
    
    if not input_str or str(input_str).lower() == 'nan':
        return None

    input_clean = str(input_str).strip().upper()
    resolved_name_lower = IDtoName.get(input_clean)
    
    if not resolved_name_lower:
        resolved_name_lower = IDtoName.get(input_clean.replace(" ", ""))
    if not resolved_name_lower:
        resolved_name_lower = IDtoName.get(input_clean[:5])
    target_name_lower = resolved_name_lower if resolved_name_lower else input_clean.lower()

    parts = target_name_lower.split()
    if parts:
        first = parts[0]
        last = " ".join(parts[1:])
        
        if first == 'christophe':
            target_name_lower = " ".join(['christopher', last])
        elif first == 'danny' and last == 'creswell':
            target_name_lower = 'daniel creswell'
        elif first == 'theodore' and last == 'lund':
            target_name_lower = 'ted lund'

    return reps.get(target_name_lower)

def main():
    while True:
        print("\n" + "="*40)
        print("      GENTER CAPITAL AUTOMATION")
        print("="*40)
        print("1. Generate Primerica AUM Pivot Table")
        print("2. Run Primerica Div Model Pipeline (additions + opens)")
        print("3. Run Both Pipelines")
        print("Q. Quit")
        print("-" * 40)
        
        choice = input("Select an option: ").strip().upper()
        fitlist = input('Enter FULL PATH of the fit list (<MONTH>-<YEAR): ').strip().replace('"', '')
        fitlist_sheet = input('Enter the name of the fit list sheet within that excel file (FIT): ')
        thisMonth = input('Enter FULL PATH of the Primerica excel file (ModelProvider_AUM_RNC_<MONTH><YEAR>.xlsx): ').strip().replace('"', '')
        thisMonthSheet = input('Enter the name of the Primerica sheet within that excel file (Account-Rep Details): ')
        lastMonth = input("Enter FULL PATH of last month's pivot table excel file (ModelProvider_AUM_RNC_<MONTH><YEAR>_Pivot.xlsx): ").strip().replace('"', '')
        
        if choice == '1':
            lastMonthTableSheet = input('Enter the name of the pivot table sheet on that excel file (AUM Pivot - <month> <year>): ')

            load_reps_from_xlsx(fitlist, fitlist_sheet)
            attribute_accounts(thisMonth, thisMonthSheet)
            load_previous_month_data(lastMonth, lastMonthTableSheet)
            export_to_pivot(fitlist, fitlist_sheet, thisMonth, thisMonthSheet, lastMonth, lastMonthTableSheet)
        elif choice == '2':
            lastMonthAccountSheet = input("Enter the name of the sheet on last month's Primerica table's file (Account-Rep Details): ")
            
            load_reps_from_xlsx(fitlist, fitlist_sheet)
            Primerica_Div_Model(thisMonth, thisMonthSheet, lastMonth, lastMonthAccountSheet)
        elif choice == '3':
            lastMonthTableSheet = input('Enter the name of the pivot table sheet on that excel file (AUM Pivot - <month> <year>): ')
            lastMonthAccountSheet = input("Enter the name of the sheet on last month's Primerica table's file (Account-Rep Details): ")
            
            load_reps_from_xlsx(fitlist, fitlist_sheet)
            attribute_accounts(thisMonth, thisMonthSheet)
            load_previous_month_data(lastMonth, lastMonthTableSheet)
            export_to_pivot(fitlist, fitlist_sheet, thisMonth, thisMonthSheet, lastMonth, lastMonthTableSheet)
            Primerica_Div_Model(thisMonth, thisMonthSheet, lastMonth, lastMonthAccountSheet)
        elif choice == 'Q':
            print("Closing application. Have a good one!")
            break
        else:
            print("Invalid selection. Please enter 1, 2, 3, or Q.")

if __name__ == "__main__":
    main()  