import pandas as pd
import numpy as np
import os

# First edit on corp machine
# Global dictionary to store objects
reps = {}
nameToID = {}
States = {
    'AL', 'AK', 'AZ', 'AR', 'AB', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 
    'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD', 
    'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 
    'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'PR', 'RI', 'SC', 
    'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY', 'DC'
}

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
        full_name = f"{str(row['First']).strip()} {str(row['Last']).strip()}"
        total = float(row['LifeTime'])
        if full_name.lower() in reps:
            if reps[full_name.lower()].Lifetime_Total > total: continue
        clean_ID = str(row['ID']).replace(' ', '').strip()
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
        nameToID[full_name.lower()] = clean_ID
            
def attribute_accounts():
    global reps
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    #sheet_name = input('Enter sheet name (it should be stored in the same directory as this script. Give the name and filetype, such as 12-25.xlsx): ')
    file_path = os.path.join(base_dir, 'ModelProvider_AUM_RNC_DEC2025_Pivot.xlsx')
    df = pd.read_excel(file_path, sheet_name='Account-Rep Details', header=0)
    df.columns = df.columns.str.strip()
    
    for _, row in df.iterrows():
        if str(row['Rep Name']).lower() in reps:
            reps[row['Rep Name'].lower()].add_account(row['Total Assets'])
    
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
        
        
load_reps_from_xlsx()
attribute_accounts()
print(f'loaded {len(reps)} reps')
while True:
    toPrint = input('Type Advisor name: ')
    if toPrint.lower() in reps:
        print(reps[toPrint.lower()])
    else:
        print('Not found')