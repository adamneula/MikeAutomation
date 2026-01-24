import pandas as pd
import numpy as np
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
        
def load_reps_from_xlsx():
    global reps
    
    # Setup pathing
    base_dir = os.path.dirname(os.path.abspath(__file__))
    #sheet_name = input('Enter sheet name (it should be stored in the same directory as this script. Give the name and filetype, such as 12-25.xlsx): ')
    file_path = os.path.join(base_dir, '12-25.xlsx')
    
    # header=1 skips the 'Owned...' row and uses the 'Code, Mutual...' row as headers
    df = pd.read_excel(file_path, sheet_name='FIT', header=1)
    
    # Standardize column names to remove any accidental spaces
    df.columns = df.columns.str.strip()
    
    # Drop rows where ID is missing
    df = df.dropna(subset=['ID'])
    
    for _, row in df.iterrows():
        if str(row['First']).strip().lower() == 'cristophe': firstName = 'Christopher'
        else: firstName = str(row['First']).strip()
        full_name = f"{str(row['First']).strip()} {str(row['Last']).strip()}"
        clean_ID = str(row['ID']).replace(' ', '').strip()
        # Map ID to name immediately so secondary IDs are captured
        IDtoName[clean_ID] = full_name.lower()
        
        total = float(row['LifeTime'])
        if full_name.lower() in reps:
            if reps[full_name.lower()].Lifetime_Total > total: continue
        state = str(row['State']).strip()
        email = str(row['Pol Email']).strip()
        territory = str(row['Territory']).strip()
        AE = ''
        #Sets central to East region and assigns AE accordingly
        if territory == 'Central':
            territory = 'East' 
            AE = 'Rob Hunt'
        elif territory == 'East': AE = 'Rob Hunt'
        elif territory == 'West': AE = 'MeiWah Wong'
        reps[full_name.lower()] = Representatives(full_name, clean_ID, state, email, AE, territory, total)
            
def attribute_accounts():
    global reps
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    #sheet_name = input('Enter sheet name (it should be stored in the same directory as this script. Give the name and filetype, such as 12-25.xlsx): ')
    file_path = os.path.join(base_dir, 'ModelProvider_AUM_RNC_DEC2025_Pivot.xlsx')
    df = pd.read_excel(file_path, sheet_name='Account-Rep Details', header=0)
    df.columns = df.columns.str.strip()
    
    for index, row in df.iterrows():
        clean_Name = str(row['Rep Name']).strip()
        if clean_Name == 'nan': continue
        
        if clean_Name.lower() in reps:
            reps[clean_Name.lower()].add_account(row['Total Assets'])
            
        elif clean_Name[:5] in IDtoName:
            reps[IDtoName[clean_Name[:5]]].add_account(row['Total Assets'])

        # Check full ID with spaces removed (Fixes "BCU 27" vs "BCU27")
        elif clean_Name.replace(' ', '') in IDtoName:
            reps[IDtoName[clean_Name.replace(' ', '')]].add_account(row['Total Assets'])

        else:
            # print(f"Error, name not found for {row['Rep Name']} on row {index}")
            NamesNotFound[clean_Name] = index
    
def final_processing():
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
     
load_reps_from_xlsx()
attribute_accounts()
validation()
print(f'loaded {len(reps)} reps')
while True:
    toPrint = input('Type Advisor name: ')
    if toPrint.lower() in reps:
        print(reps[toPrint.lower()])
    else:
        print('Not found')