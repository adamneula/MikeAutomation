import pandas as pd
import os
from datetime import datetime, timedelta
import numpy as np
from dateutil.relativedelta import relativedelta
from tqdm import tqdm

tqdm.pandas()

eastBalance = 0
westBalance = 0

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
    global reps, eastBalance, westBalance

    df = load_dynamic_df(Primerica_Dir, Primerica_Sheet_Name, 'Rep Name')
    df.columns = df.columns.str.strip()
    
    for index, row in df.iterrows():
        clean_Name = str(row['Rep Name']).strip()
        if clean_Name == 'nan': continue
        elif clean_Name.lower().split()[0] == 'christophe':
            clean_Name = " ".join(['CHRISTOPHER'] + clean_Name.upper().split()[1:])
        elif clean_Name.lower() == 'danny creswell': clean_Name = 'DANIEL CRESWELL'
        
        if clean_Name.lower() in reps:
            reps[clean_Name.lower()].add_account(row['Total Assets'])
            if reps[clean_Name.lower()].Territory == 'East': eastBalance += row['Total Assets']
            else: westBalance += row['Total Assets']
        elif clean_Name[:5] in IDtoName:
            reps[IDtoName[clean_Name[:5]]].add_account(row['Total Assets'])
            if reps[IDtoName[clean_Name[:5]]].Territory == 'East': eastBalance += row['Total Assets']
            else: westBalance += row['Total Assets']

        elif clean_Name.replace(' ', '') in IDtoName:
            reps[IDtoName[clean_Name.replace(' ', '')]].add_account(row['Total Assets'])
            if reps[IDtoName[clean_Name.replace(' ', '')]].Territory == 'East': eastBalance += row['Total Assets']
            else: westBalance += row['Total Assets']

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
        df_prev = load_dynamic_df(prev_month_file, prev_month_sheet, 'Primary Rep ID')
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
    global reps, eastBalance, westBalance
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
            'True State ': r.True_State,
            'AE ': r.AE,
            'Territory ': r.Territory,
            'Email': r.Email,
            'Primary Rep ID ': r.Primary_Rep_ID,
            'Ranking ': r.Ranking,
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
            workbook  = writer.book
            
                        # --- Define Styles ---
            font_settings = {'font_name': 'Aptos Narrow', 'font_size': 11}

            dark_blue = workbook.add_format({**font_settings, 'bg_color': '#1F4E78', 'font_color': 'white', 'bold': True, 'border': 0, 'align': 'left'})
            light_blue = workbook.add_format({**font_settings, 'bg_color': "#C7E5F3", 'font_color': 'black', 'bold': True, 'border': 0, 'align': 'left'})
            money_fmt = workbook.add_format({**font_settings, 'num_format': '$#,##0.00', 'border': 0})
            percent_fmt = workbook.add_format({**font_settings, 'num_format': '0.0%', 'border': 0, 'align': 'left'})
            border_fmt = workbook.add_format({**font_settings, 'border': 0})
            bold_border = workbook.add_format({**font_settings, 'bold': True, 'border': 0})
            left_align_fmt = workbook.add_format({**font_settings, 'font_size': 11, 'align': 'left', 'border': 0})
            no_border_fmt = workbook.add_format({**font_settings, 'border': 0})
            
            #Source Sheets
            try:
                df_details = pd.read_excel(details_path, sheet_name=details_sheet)
                df_details.to_excel(writer, sheet_name='Account-Rep Details', index=False)
                
                details_ws = writer.sheets['Account-Rep Details']
                
                header_bold = workbook.add_format({'bold': True, 'font_name': 'Aptos Narrow'})
                
                details_ws.set_column(0, len(df_details.columns) - 1, 25)
                
                for col_num, value in enumerate(df_details.columns.values):
                    details_ws.write(0, col_num, value, header_bold)
                
                details_ws.autofilter(0, 0, len(df_details), len(df_details.columns) - 1)
                
            except:
                print(f"\ACCOUNT-REP DETAIL TAB ERROR: {e}")
                
            try:
                # 1. Load the raw details
                df_details_raw = pd.read_excel(details_path, sheet_name=details_sheet)
                df_details_raw.columns = df_details_raw.columns.str.strip()
                
                primmy_data = []
                
                # 2. Iterate and build the records
                for _, row in df_details_raw.iterrows():
                    # Look up the rep object to get the 'True State', 'AE', 'Territory', etc.
                    rep_name_raw = str(row.get('Rep Name', '')).strip()
                    rep_obj = rep_lookup(rep_name_raw)
                    
                    # Build the dictionary for this row
                    record = {
                        'ModelName': row.get('ModelName', ''),
                        'accountid': row.get('accountid', ''),
                        'Total Assets Excluding Cash': row.get('Total Assets Excluding Cash', 0),
                        'Total Cash': row.get('Total Cash', 0),
                        'Total Assets': row.get('Total Assets', 0),
                        'Target %': row.get('Target %', ''),
                        'Target Value': row.get('Target Value', ''),
                        'Cust Rep #': row.get('Cust Rep #', ''),
                        'Per Rep #': row.get('Per Rep #', ''),
                        'AccountState': row.get('AccountState', ''),
                        'Rep Name': rep_name_raw,
                        'Rep Address 1': row.get('Rep Address 1', ''),
                        'Rep Address 2': row.get('Rep Address 2', ''),
                        'Rep Address 3': row.get('Rep Address 3', ''),
                        'Primary Rep ID': rep_obj.Primary_Rep_ID if rep_obj else '',
                        'Secondary Rep Name': row.get('Secondary Rep Name', ''),
                        'Secondary Rep ID': row.get('Secondary Rep ID', ''),
                        'Rep City': row.get('Rep City', ''),
                        'Rep State': row.get('Rep State', ''),
                        'Rep Zip': row.get('Rep Zip', ''),
                        'Rep Email (via Fit List)': rep_obj.Email if rep_obj else '',
                        'Rep Phone': row.get('Rep Phone', ''),
                        'True State': rep_obj.True_State if rep_obj else '',
                        'AE': rep_obj.AE if rep_obj else '',
                        'Territory': rep_obj.Territory if rep_obj else ''
                    }
                    primmy_data.append(record)

                # 3. Create DataFrame and Write to Sheet
                df_primmy = pd.DataFrame(primmy_data)
                sheet_name_primmy = "Data for Primmy AUM"
                df_primmy.to_excel(writer, sheet_name=sheet_name_primmy, index=False, startrow=1, header=False)
                primmy_ws = writer.sheets[sheet_name_primmy]
                
                # 4. Formatting
                header_bold_fmt = workbook.add_format({**font_settings, 'bold': True, 'font_color': 'black', 'bottom': 1})
                horiz_data_fmt = workbook.add_format({**font_settings, 'bottom': 1, 'top': 1, 'border_color': '#D9D9D9'})
                money_horiz_fmt = workbook.add_format({**font_settings, 'num_format': '$#,##0.00', 'bottom': 1, 'top': 1, 'border_color': '#D9D9D9'})

                for i, col_name in enumerate(df_primmy.columns):
                    # Dynamic Width
                    column_data = df_primmy[col_name].fillna('')
                    max_len = max(column_data.astype(str).map(len).max(), len(str(col_name))) + 3
                    max_len = min(max_len, 60)
                    
                    # Apply money format to asset columns, horizontal border to others
                    if any(x in col_name for x in ['Assets', 'Cash', 'Value']):
                        primmy_ws.set_column(i, i, max_len, money_horiz_fmt)
                    else:
                        primmy_ws.set_column(i, i, max_len, horiz_data_fmt)
                        
                    primmy_ws.write(0, i, col_name, header_bold_fmt)

                primmy_ws.autofilter(0, 0, len(df_primmy), len(df_primmy.columns) - 1)
                primmy_ws.freeze_panes(1, 0)

            except Exception as e:
                print(f"\nPRIMMY DATA SHEET ERROR: {e}")
            
           
            rank_colors = {
                'AAA': '#4FAD5B', 'AA': '#9FCE63', 'A': '#DFF1D3',
                'BB': '#79ADEA', 'B': '#ADC8E9', 'C': '#FFFF54'
            }

            num_rows = len(df_output)
            num_cols = len(df_output.columns)
            
            df_output.to_excel(writer, sheet_name=f'AUM Pivot - {(datetime.now() - relativedelta(months=1)).strftime("%b %y")}', index=False, startrow=3, header=False)
            worksheet = writer.sheets[f'AUM Pivot - {(datetime.now() - relativedelta(months=1)).strftime("%b %y")}']    
            worksheet.autofilter(2, 0, num_rows + 2, 1)
            worksheet.autofilter(2, 8, num_rows + 2, 18)

            for col_num, value in enumerate(df_output.columns.values):
                if col_num < 6: # Left side group
                    worksheet.write(2, col_num, value, light_blue)
                elif col_num > 7: # Right side group
                    worksheet.write(2, col_num, value, dark_blue)
                else: # Spacers G and H
                    worksheet.write(2, col_num, "", no_border_fmt)

            apply_excel_highlighting(workbook, worksheet, df_output)
            
            for i, col in enumerate(df_output.columns):

                column_data = df_output[col].fillna('')
                max_len = max(column_data.astype(str).map(len).max(), len(str(col)))
                
                max_len = min(max_len, 100) 
                
                if any(x.lower() in col.lower() for x in ['assets', 'aum', 'dollar']):
                    worksheet.set_column(i, i, 21, money_fmt)
                elif 'Change' in col:
                    worksheet.set_column(i, i, 12, percent_fmt)
                else:
                    worksheet.set_column(i, i, max_len, border_fmt)

            worksheet.set_column(6, 7, 10)
            worksheet.set_column(19, 20, 10)

            legend_start_row = 4
            legend_col_label = 21 # Column U
            legend_col_val = 22   # Column V
            
            worksheet.write(2, legend_col_label, "Ranking Legend (minimum):", workbook.add_format({'font': 'Aptos Narrow', 'bold': True}))
            
            legend_items = [
                ('AAA', 10000000, '#00B050'), # Green
                ('AA', 5000000, '#92D050'),   # Light Green
                ('A', 2000000, '#FCE4D6'),    # Peach/Tan
                ('BB', 1000000, '#00B0F0'),   # Blue
                ('B', 250000, '#B4C6E7'),     # Light Blue
                ('C', 0, '#FFFF00'),          # Yellow
            ]

            for i, (rank, val, color) in enumerate(legend_items):
                row = legend_start_row + i
                fmt = workbook.add_format({'font': 'Aptos Narrow', 'bg_color': color, 'border': 0})
                money_fmt_legend = workbook.add_format({'font': 'Aptos Narrow', 'bg_color': color, 'border': 0, 'num_format': '$#,##0.00'})
                
                worksheet.write(row, legend_col_label, rank, fmt)
                worksheet.write(row, legend_col_val, val, money_fmt_legend)

            # 2. Add the Summary Totals (Total Assets, East, West, etc.)
            summary_start_row = legend_start_row + len(legend_items) + 1
            bold_border = workbook.add_format({'font': 'Aptos Narrow', 'bold': True, 'border': 0})
            money_bold = workbook.add_format({'font': 'Aptos Narrow', 'bold': True, 'border': 0, 'num_format': '$#,##0.00'})
            percent_bold = workbook.add_format({'font': 'Aptos Narrow', 'bold': True, 'border': 0, 'num_format': '0.00%'})
            
            worksheet.set_column(21, 21, 26)
            worksheet.set_column(22, 22, 17)

            # Calculations for the summary
            # --- Pull Previous Month Total from hardcoded cell W16 ---
            prev_aum = "" 
            try:
                # Read the previous pivot sheet
                df_prev_grid = pd.read_excel(pivot_path, sheet_name=pivot_sheet, header=None)
                
                # Excel 'W16' corresponds to row index 15, column index 22 (0-indexed)
                # pandas uses [row, col] format
                val = df_prev_grid.iloc[11, 22] 
                
                if pd.notna(val) and isinstance(val, (int, float)):
                    prev_aum = float(val)
                else:
                    print(f"\n[!] WARNING: Cell W16 in '{pivot_sheet}' is empty or not a number.")
                    print("    Please manually check last month's pivot table.")

            except Exception:
                print(f"\n[!] WARNING: Could not access cell W16 in {os.path.basename(pivot_path)}.")
                print("    THE PREVIOUS AUM AND MOM GROWTH FIELDS WILL BE BLANK.")
                print("    Please manually copy these values from last month's pivot table.")

            # Ensure MoM Growth only calculates if we have a valid number
            if isinstance(prev_aum, (int, float)) and prev_aum != 0:
                mom_growth = ((eastBalance + westBalance) - prev_aum) / prev_aum
            else:
                mom_growth = ""
                
            summary_data = [
                (f"Total Assets {(datetime.now().replace(day=1) - timedelta(days=1)).strftime('%m/%d/%Y')}:", eastBalance + westBalance, money_bold),
                ("East", eastBalance, money_bold),
                ("West", westBalance, money_bold),
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
    font_base = {'font_name': 'Aptos Narrow', 'font_size': 11}

    # 2. Pre-create formats with Aptos Narrow
    pos_fmt = workbook.add_format({**font_base, 'bg_color': '#CEEED0', 'border': 1, 'num_format': '0.00%', 'align': 'right', 'font_color': '#285F17'})
    neg_fmt = workbook.add_format({**font_base, 'bg_color': '#F6C9CE', 'border': 1, 'num_format': '0.00%', 'align': 'right', 'font_color': '#8F1B15'})
        
    rank_formats = {}
    for rank, hex_code in rank_colors.items():
        rank_formats[rank] = {
            'text': workbook.add_format({'font_name': 'Aptos Narrow', 'bg_color': hex_code, 'border': 1, 'align': 'left'}),
            'money': workbook.add_format({'font_name': 'Aptos Narrow', 'bg_color': hex_code, 'border': 1, 'num_format': '$#,##0.00'})
        }

    # 3. Identify column indices safely
    try:
        rank_col_idx = df.columns.get_loc('Ranking ')
        assets_col_idx = df.columns.get_loc('Sum of Total Assets ')
        assets1_col_idx = df.columns.get_loc('Sum of Total Assets .1')
        mom_change_idx = df.columns.get_loc('MoM Change')
    except KeyError as e:
        print(f"Warning: Could not find column {e} for highlighting.")
        return

    # 4. Loop through every row
    for row_num in range(len(df)):
        excel_row = row_num + 3
        
        # --- Handle Ranking and Assets ---
        rank_val = str(df.iloc[row_num]['Ranking ']).strip()
        
        if rank_val in rank_formats:
            # Highlight Ranking
            worksheet.write(excel_row, rank_col_idx, rank_val, rank_formats[rank_val]['text'])
            
            # Highlight both Asset columns with the same rank color
            asset_val = df.iloc[row_num]['Sum of Total Assets .1']
            worksheet.write(excel_row, assets_col_idx, asset_val, rank_formats[rank_val]['money'])
            worksheet.write(excel_row, assets1_col_idx, asset_val, rank_formats[rank_val]['money'])
        
        # --- Handle MoM Change Highlighting ---
        mom_change_val = df.iloc[row_num]['MoM Change']
        
        # Check if it's a number and not NaN
        if pd.notna(mom_change_val):
            if mom_change_val > 0:
                worksheet.write(excel_row, mom_change_idx, mom_change_val, pos_fmt)
            elif mom_change_val < 0:
                worksheet.write(excel_row, mom_change_idx, mom_change_val, neg_fmt)
    
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
        worksheet.autofilter(2, 0, 2 + len(df), len(df.columns) - 1)
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

def load_dynamic_df(path, sheet, target_col, max_search=10):
    """Searches for the header row within the first max_search rows."""
    for i in range(max_search + 1):
        try:
            df = pd.read_excel(path, sheet_name=sheet, header=i)
            df.columns = df.columns.str.strip()
            if target_col in df.columns:
                return df
        except Exception:
            continue
    raise KeyError(f"Could not find header with column '{target_col}' in the first {max_search} rows of {path}")

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
        
        if choice == 'Q':
            print("Closing application. Have a good one!")
            break
        
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
        else:
            print("Invalid selection. Please enter 1, 2, 3, or Q.")

if __name__ == "__main__":
    main()  