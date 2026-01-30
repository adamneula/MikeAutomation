import pandas as pd
import os

# First edit on corp machine
# Global dictionary to store objects
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
    
    # Setup pathing
    #base_dir = os.path.dirname(os.path.abspath(__file__))
    #sheet_name = input('Enter sheet name (it should be stored in the same directory as this script. Give the name and filetype, such as 12-25.xlsx): ')
    #file_path = os.path.join(base_dir, Fit_List_Dir)
    
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
    
    # base_dir = os.path.dirname(os.path.abspath(__file__))
    # file_path = os.path.join(base_dir, Primerica_Dir)
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
    
def assign_ranking():
    global reps
    for rep in reps:
        #assign ranking
        if reps[rep].Sum_of_Total_Assets == 0: continue
        elif reps[rep].Sum_of_Total_Assets < 250000: reps[rep].Ranking = 'C'
        elif reps[rep].Sum_of_Total_Assets < 1000000: reps[rep].Ranking = 'B'
        elif reps[rep].Sum_of_Total_Assets < 2000000: reps[rep].Ranking = 'BB'
        elif reps[rep].Sum_of_Total_Assets < 5000000: reps[rep].Ranking = 'A'
        elif reps[rep].Sum_of_Total_Assets < 10000000: reps[rep].Ranking = 'AA'
        else: reps[rep].Ranking = 'AAA'
        
def validation():
    global reps
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    #sheet_name = input('Enter sheet name (it should be stored in the same directory as this script. Give the name and filetype, such as 12-25.xlsx): ')
    file_path = os.path.join(base_dir, 'ModelProvider_AUM_RNC_DEC2025_Pivot.xlsx')
    df = pd.read_excel(file_path, sheet_name='AUM Pivot - Dec 25', header=2)
    df.columns = df.columns.str.strip()
    
    missing_count = 0
    for _, row in df.iterrows():
        clean_name = ' '.join(str(row['Advisor Name']).split()).lower()
        if clean_name == 'nan': continue
        if clean_name not in reps:
            missing_count += 1
            print(f'missed {row["Advisor Name"]} from my dataset (miss {missing_count})')

def load_previous_month_data(filename, sheetname):
    global reps, IDtoName
    
    # base_dir = os.path.dirname(os.path.abspath(__file__))
    # file_path = os.path.join(base_dir, filename)
    
    # Load the previous month's pivot (assuming headers on row 2 as in your validation)
    try:
        df_prev = pd.read_excel(filename, sheet_name=sheetname, header=2)
        df_prev.columns = df_prev.columns.str.strip()
    except Exception as e:
        print(f"Error loading {filename}: {e}")
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
        else:
            # Optional: handle IDs found in the file that aren't in your current master list
            pass
            
                
def export_to_pivot(output_filename, fit_path='', fit_sheet='', details_path='', details_sheet='', pivot_path='', pivot_sheet=''):
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
    output_path = os.path.join(base_dir, output_filename)

    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_output.to_excel(writer, sheet_name='AUM Pivot - Dec 25', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['AUM Pivot - Dec 25']

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

        print(f"\nSUCCESS: Report generated at {output_path}")
    
    except Exception as e:
        # print(f"\nERROR: {e}")
        pass
        
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
    
fitlist = input('Enter FULL PATH of the fit list: ').strip().replace('"', '')
fitlist_sheet = input('Enter the name of the fit list sheet within that excel file (case sensitive): ')
primerica_xlsx = input('Enter FULL PATH of the Primerica excel file: ').strip().replace('"', '')
primerica_sheet = input('Enter the name of the Primerica sheet within that excel file (case sensitive): ')
prev_table = input("Enter FULL PATH of last month's pivot excel file: ").strip().replace('"', '')
prev_sheet = input('Enter the name of the pivot table sheet on that excel file (case sensitive): ')
to_make = input('Enter name for the new file (include file extension): ').strip().replace('"', '')

load_reps_from_xlsx(fitlist, fitlist_sheet)
attribute_accounts(primerica_xlsx, primerica_sheet)
assign_ranking()
load_previous_month_data(prev_table, prev_sheet)
export_to_pivot(to_make, fitlist, fitlist_sheet, primerica_xlsx, primerica_sheet, prev_table, prev_sheet)

while True:
    toPrint = input('Type Advisor name: ')
    if toPrint.lower() in reps:
        print(reps[toPrint.lower()])
    elif toPrint.lower() in IDtoName:
        print(IDtoName[toPrint.lower])
    else:
        print('Not found')
