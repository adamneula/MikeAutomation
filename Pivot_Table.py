from Utils import *
from Rep_Objects import *
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

eastBalance = 0
westBalance = 0

def attribute_accounts(Primerica_Dir, Primerica_Sheet_Name):
    global eastBalance, westBalance

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

    for rep in reps:
        #assign ranking
        if reps[rep].Sum_of_Total_Assets == 0: continue
        elif reps[rep].Sum_of_Total_Assets < 250000: reps[rep].Ranking = 'C'
        elif reps[rep].Sum_of_Total_Assets < 1000000: reps[rep].Ranking = 'B'
        elif reps[rep].Sum_of_Total_Assets < 2000000: reps[rep].Ranking = 'BB'
        elif reps[rep].Sum_of_Total_Assets < 5000000: reps[rep].Ranking = 'A'
        elif reps[rep].Sum_of_Total_Assets < 10000000: reps[rep].Ranking = 'AA'
        else: reps[rep].Ranking = 'AAA'
        
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
                print(f"\\ACCOUNT-REP DETAIL TAB ERROR: {e}")
                
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
                        'Total Assets': row.get('Total Assets', 0),
                        'AccountState': row.get('AccountState', ''),
                        'Rep Name': rep_name_raw,
                        'Primary Rep ID': rep_obj.Primary_Rep_ID if rep_obj else '',
                        'Secondary Rep Name': row.get('Secondary Rep Name', ''),
                        'Secondary Rep ID': row.get('Secondary Rep ID', ''),
                        'Rep City': row.get('Rep City', ''),
                        'Rep State': row.get('Rep State', ''),
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
                    max_len = min(max_len, 40)
                    
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
                
                # Excel 'W12' corresponds to row index 11, column index 22 (0-indexed)
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